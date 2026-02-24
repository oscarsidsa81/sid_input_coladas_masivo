# -*- coding: utf-8 -*-

import base64
import re
from io import BytesIO

from odoo import _, fields, models
from odoo.exceptions import UserError

try:
    import openpyxl
except ImportError:
    openpyxl = None


class StockMoveColadas(models.Model):
    _inherit = "stock.move"

    sid_coladas_masivo = fields.Char(
        string="Introduce coladas",
        store=True,
        help=(
            "Este campo requiere pares de datos 'Colada'/'Cantidad hecha' "
            "para realizar entradas múltiples en stock.move.line de cada stock.move"
        ),
    )
    sid_coladas_procesado = fields.Boolean(
        string="Coladas procesadas",
        default=False,
        copy=False,
        help="Marca técnica para evitar reprocesar coladas masivas ya aplicadas.",
    )


class StockPicking(models.Model):
    _inherit = "stock.picking"

    sid_coladas_excel = fields.Binary(
        string="Excel coladas",
        attachment=False,
        help="Sube la plantilla de coladas para precargar el campo sid_coladas_masivo.",
    )
    sid_coladas_excel_filename = fields.Char(string="Nombre archivo coladas")

    def _check_openpyxl(self):
        if openpyxl is None:
            raise UserError(_("Falta la librería 'openpyxl' en el servidor."))

    def _validate_coladas_format(self, coladas_raw, move_id):
        """Valida formato string/float/string/float y devuelve valor normalizado."""
        if not coladas_raw:
            return ""

        partes = [p.strip() for p in str(coladas_raw).split(";") if p and p.strip()]

        if len(partes) % 2 != 0:
            raise UserError(
                _(
                    "Movimiento %(move_id)s: estructura inválida en sid_coladas_masivo. "
                    "Debe seguir pares Colada;Cantidad."
                )
                % {"move_id": move_id}
            )

        normalizadas = []
        for i in range(0, len(partes), 2):
            colada = partes[i]
            cantidad_txt = partes[i + 1].replace(",", ".")

            if not colada:
                raise UserError(
                    _("Movimiento %(move_id)s: colada vacía en el par %(par)s.")
                    % {"move_id": move_id, "par": (i // 2) + 1}
                )

            try:
                cantidad = float(cantidad_txt)
            except (TypeError, ValueError):
                raise UserError(
                    _(
                        "Movimiento %(move_id)s: '%(valor)s' no es una cantidad válida. "
                        "Formato esperado: string;float;string;float..."
                    )
                    % {"move_id": move_id, "valor": partes[i + 1]}
                )

            normalizadas.extend([colada, str(cantidad)])

        return ";".join(normalizadas)

    def _get_export_xmlid(self, record):
        """Devuelve el XMLID tal como lo exporta Odoo en Identificación externa."""
        if not record:
            return ""
        data = record.sudo().export_data(["id"])
        if data.get("datas"):
            return data["datas"][0][0] or ""
        return ""

    def _get_move_from_row(self, picking, move_id_value=False, move_xmlid=False):
        move = self.env["stock.move"]

        # Permite id numérico o xmlid en cualquiera de las dos columnas.
        for candidate in (move_id_value, move_xmlid):
            if not candidate:
                continue

            if isinstance(candidate, str):
                val = candidate.strip()
            else:
                val = candidate

            if not val:
                continue

            try:
                move = self.env["stock.move"].browse(int(val))
                if move.exists():
                    break
            except (TypeError, ValueError):
                pass

            # Soporta xmlids tipo __export__.stock_move_<id>_<hash> incluso si el xmlid no existe.
            if isinstance(val, str):
                match = re.search(r"stock_move_(\d+)(?:_|$)", val)
                if match:
                    move = self.env["stock.move"].browse(int(match.group(1)))
                    if move.exists():
                        break

            if isinstance(val, str) and "." in val:
                try:
                    move = self.env.ref(val, raise_if_not_found=False)
                except ValueError:
                    move = self.env["stock.move"]
                if move and move._name == "stock.move":
                    break

        if not move or move._name != "stock.move" or move.picking_id != picking:
            raise UserError(
                _(
                    "No se pudo resolver el movimiento de la fila (move_id='%(move_id)s', move_external_id='%(xmlid)s') "
                    "para el albarán %(picking)s."
                )
                % {
                    "move_id": move_id_value or "",
                    "xmlid": move_xmlid or "",
                    "picking": picking.display_name,
                }
            )
        return move

    def action_cargar_coladas_desde_excel(self):
        self._check_openpyxl()
        for picking in self:
            if not picking.sid_coladas_excel:
                raise UserError(_("Debes subir un archivo Excel de coladas."))

            wb = openpyxl.load_workbook(
                filename=BytesIO(base64.b64decode(picking.sid_coladas_excel)),
                data_only=True,
            )
            ws = wb.active

            header_values = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
            header_map = {
                str(col).strip(): idx
                for idx, col in enumerate(header_values)
                if col is not None and str(col).strip()
            }

            rows = list(ws.iter_rows(min_row=2, values_only=True))
            if not rows:
                raise UserError(_("El Excel no contiene filas de datos para procesar."))

            def get_cell(row_values, field_name, fallback_idx=None, aliases=None):
                names = [field_name] + (aliases or [])
                for name in names:
                    idx = header_map.get(name)
                    if idx is not None and idx < len(row_values):
                        return row_values[idx]
                if fallback_idx is not None and fallback_idx < len(row_values):
                    return row_values[fallback_idx]
                return False

            updated = 0
            for row in rows:
                if not any(row):
                    continue

                # Soporta plantilla técnica y exportaciones de UI con cabeceras traducidas.
                move_id_value = get_cell(
                    row,
                    "move_id",
                    1,
                    aliases=["Movimientos de stock/ID", "Movimientos de stock/Identificación externa"],
                )
                move_xmlid = get_cell(
                    row,
                    "move_external_id",
                    None,
                    aliases=["Movimientos de stock/Identificación externa", "External ID", "Identificación externa"],
                )
                coladas_value = get_cell(
                    row,
                    "sid_coladas_masivo",
                    2 if len(row) <= 3 else (11 if len(header_map) > 10 else 9),
                    aliases=["Movimientos de stock/Introduce coladas", "Introduce coladas"],
                )

                move = picking._get_move_from_row(
                    picking, move_id_value=move_id_value, move_xmlid=move_xmlid
                )
                move.sid_coladas_masivo = picking._validate_coladas_format(coladas_value, move.id)
                move.sid_coladas_procesado = False
                updated += 1

            if not updated:
                raise UserError(
                    _("No se actualizó ningún movimiento. Revisa el contenido del Excel.")
                )

            picking.message_post(
                body=_(
                    "✅ Se ha cargado el Excel de coladas y validado la estructura "
                    "string/float en %(count)s movimientos."
                )
                % {"count": updated}
            )

    def action_procesar_coladas(self):
        for picking in self:
            if picking.state in ("done", "cancel"):
                raise UserError(
                    _("No se pueden procesar coladas en un albarán cerrado o cancelado.")
                )

            errores = []
            bloques = []

            for move in picking.move_ids_without_package:
                if not move.sid_coladas_masivo:
                    continue

                if move.sid_coladas_procesado:
                    continue

                if move.product_id.tracking == "none":
                    errores.append(
                        f"Producto {move.product_id.display_name} no está configurado con seguimiento por lote."
                    )
                    continue

                partes = [p.strip() for p in move.sid_coladas_masivo.split(";") if p.strip()]

                if len(partes) % 2 != 0:
                    errores.append(f"Movimiento {move.id}: número impar de elementos.")
                    continue

                product = move.product_id
                lotes_registrados = []

                nombres_lote = partes[0::2]
                lotes_existentes = self.env["stock.production.lot"].search(
                    [("product_id", "=", product.id), ("name", "in", nombres_lote)]
                )
                lot_dict = {l.name: l for l in lotes_existentes}

                for i in range(0, len(partes), 2):
                    lote_nombre = partes[i]
                    cantidad_str = partes[i + 1].replace(",", ".")

                    try:
                        qty = float(cantidad_str)
                    except Exception:
                        errores.append(
                            f"Movimiento {move.id}: cantidad inválida '{cantidad_str}'."
                        )
                        continue

                    if qty <= 0:
                        continue

                    lot = lot_dict.get(lote_nombre)
                    if not lot:
                        lot = self.env["stock.production.lot"].create(
                            {
                                "name": lote_nombre,
                                "product_id": product.id,
                                "company_id": picking.company_id.id,
                            }
                        )
                        lot_dict[lote_nombre] = lot

                    existing_line = self.env["stock.move.line"].search(
                        [("move_id", "=", move.id), ("lot_id", "=", lot.id)], limit=1
                    )

                    if existing_line:
                        existing_line.qty_done += qty
                    else:
                        self.env["stock.move.line"].create(
                            {
                                "move_id": move.id,
                                "picking_id": picking.id,
                                "product_id": product.id,
                                "lot_id": lot.id,
                                "qty_done": qty,
                                "location_id": move.location_id.id,
                                "location_dest_id": move.location_dest_id.id,
                                "product_uom_id": move.product_uom.id,
                            }
                        )

                    lotes_registrados.append((lote_nombre, qty))

                if lotes_registrados:
                    move.sid_coladas_procesado = True
                    bloques.append(
                        {
                            "producto": product.display_name,
                            "demanda": move.product_uom_qty,
                            "lineas": lotes_registrados,
                        }
                    )

            if bloques:
                mensaje = "✔️ Procesamiento de coladas completado:\n\n"
                total_hecho_global = 0.0
                total_demanda_global = 0.0

                for bloque in bloques:
                    mensaje += f"{bloque['producto']}\n"
                    mensaje += "-" * 30 + "\n"

                    total_item = 0.0
                    for lote, qty in bloque["lineas"]:
                        total_item += qty
                        mensaje += f"{lote:<15}{qty:.2f}\n"

                    diferencia = total_item - bloque["demanda"]
                    icono = "🟢" if diferencia > 0 else "🔴" if diferencia < 0 else "⚪"

                    mensaje += f"{'Total hecho:':<15}{total_item:.2f}\n"
                    mensaje += f"{'Demanda:':<15}{bloque['demanda']:.2f}\n"
                    mensaje += f"{'Diferencia:':<15}{diferencia:+.2f} {icono}\n\n"

                    total_hecho_global += total_item
                    total_demanda_global += bloque["demanda"]

                diferencia_global = total_hecho_global - total_demanda_global
                icono_global = "🟢" if diferencia_global > 0 else "🔴" if diferencia_global < 0 else "⚪"

                mensaje += "Totales globales:\n"
                mensaje += f"{'Total hecho:':<18}{total_hecho_global:.2f}\n"
                mensaje += f"{'Total demanda:':<18}{total_demanda_global:.2f}\n"
                mensaje += f"{'Diferencia total:':<18}{diferencia_global:+.2f} {icono_global}"

                picking.message_post(body=f"<pre>{mensaje}</pre>")

            if errores:
                picking.message_post(
                    body="<b>⚠️ Se detectaron incidencias:</b><br/><br/>" + "<br/>".join(errores)
                )

    def action_descargar_plantilla_coladas(self):
        self._check_openpyxl()
        self.ensure_one()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Coladas"

        ws.append(
            [
                "picking_id",
                "move_id",
                "move_external_id",
                "picking_external_id",
                "reference",
                "item",
                "family",
                "desc_picking",
                "location_external_id",
                "producto",
                "demanda",
                "uom",
                "sid_coladas_masivo",
            ]
        )

        moves = self.move_ids_without_package

        for mv in moves:
            product = mv.product_id
            ws.append(
                [
                    mv.picking_id.id,
                    mv.id,
                    self._get_export_xmlid(mv),
                    self._get_export_xmlid(mv.picking_id),
                    mv.reference,
                    mv.item or "",
                    mv.family or "",
                    mv.desc_picking or "",
                    self._get_export_xmlid(mv.location_id),
                    product.display_name or "",
                    mv.product_uom_qty or 0.0,
                    mv.product_uom.name if mv.product_uom else "",
                    mv.sid_coladas_masivo or "",
                ]
            )

        widths = [14, 14, 40, 40, 30, 10, 20, 50, 35, 35, 10, 10, 45]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        filename = f"coladas_{(self.name or self.id)}.xlsx"
        attachment = self.env["ir.attachment"].create(
            {
                "name": filename,
                "type": "binary",
                "datas": base64.b64encode(buf.getvalue()),
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "res_model": self._name,
                "res_id": self.id,
            }
        )

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/{attachment.id}?download=true",
            "target": "self",
        }
