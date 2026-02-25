# -*- coding: utf-8 -*-

import base64
import re
from io import BytesIO

from odoo import _, api, fields, models
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

    sid_has_coladas = fields.Boolean(
        string="Tiene coladas",
        compute="_compute_sid_has_coladas",
    )

    @api.depends("move_ids_without_package.sid_coladas_masivo")
    def _compute_sid_has_coladas(self):
        for picking in self:
            picking.sid_has_coladas = any(
                bool(m.sid_coladas_masivo and str(m.sid_coladas_masivo).strip())
                for m in picking.move_ids_without_package
            )

    def _check_openpyxl(self):
        if openpyxl is None:
            raise UserError(_("Falta la librería 'openpyxl' en el servidor."))

    def _get_export_xmlid(self, record):
        """Devuelve el XMLID tal como lo exporta Odoo en Identificación externa."""
        if not record:
            return ""
        data = record.sudo().export_data(["id"])
        if data.get("datas"):
            return data["datas"][0][0] or ""
        return ""

    def _get_first_attr(self, record, names, default=""):
        """Devuelve el primer atributo existente con valor no vacío."""
        for name in names:
            if hasattr(record, name):
                value = getattr(record, name)
                if value not in (False, None):
                    return value
        return default

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
                "Identificación externa",
                "Movimientos de stock/Identificación externa",
                "Movimientos de stock/item",
                "Movimientos de stock/Descripción de picking",
                "Movimientos de stock/Introduce coladas",
            ]
        )

        moves = self.move_ids_without_package

        for mv in moves:
            ws.append(
                [
                    self._get_export_xmlid(mv.picking_id),
                    self._get_export_xmlid(mv),
                    self._get_first_attr(mv, ["item"], ""),
                    self._get_first_attr(mv, ["description_picking"], ""),
                    mv.sid_coladas_masivo or "",
                ]
            )

        widths = [40, 50, 20, 60, 45]
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
