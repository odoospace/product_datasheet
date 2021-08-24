# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError, AccessError
from datetime import datetime, date
from io import BytesIO
import openpyxl
import string
import base64
import tempfile
import binascii


class ProductDatasheetTemplateWizard(models.TransientModel):
    _name = 'product.datasheet.template.wizard'

    template_id = fields.Many2one('product.datasheet.template', string='Template', required=True)
    product_id = fields.Many2one('product.product', string='Product')
    file_generated = fields.Binary(string='Generated Excel file', attachment=False)

    def action_download_excel(self):
        print(f'You have chosen {self.template_id.name}')

        fp = tempfile.NamedTemporaryFile(suffix=".xlsx")
        fp.write(binascii.a2b_base64(self.template_id.file))  # self.xls_file is your binary field
        fp.seek(0)
        # path = '/home/file.xlsx'
        wb_obj = openpyxl.load_workbook(fp.name)
        sheet_obj = wb_obj.active
        general_dict = {
            'o.name': self.product_id.name,
            'o.image_1920': self.product_id.image_1920,
            'h.date': datetime.now().strftime('%Y/%m/%d'),
            'h.logo': self.env.user.company_id.logo,
            'h.regulation_footer': self.env['ir.config_parameter'].sudo().get_param(
                'product_datasheet.regulation_footer'),
            'h.text_footer': self.env['ir.config_parameter'].sudo().get_param(
                'product_datasheet.text_footer'),
        }
        for row in sheet_obj.iter_rows():
            for cell in row:
                cell_value = cell.value
                if cell_value is not None:
                    print(cell_value)
                    if cell_value.startswith('{{'):
                        cell_value = cell_value.replace(' ', '').replace('{{', '').replace('}}', '')
                        if 'i.' in cell_value:
                            if '|' in cell_value:
                                label = cell_value.split('|')[1]
                                cell_value_splitted = cell_value.split('|')[0].split('.')
                                model_obj = cell_value_splitted[1]
                                code = cell_value_splitted[2]
                                model_env = f'product.datasheet.{model_obj}'
                                if label == 'name':
                                    product_datasheet = self.env[model_env].search([('code', '=', code)])
                                    if product_datasheet:
                                        cell.value = product_datasheet.name
                                else:
                                    id_model = cell_value_splitted[3]
                                    product_datasheet = self.env[model_env].search([('id', '=', id_model)])
                                    if product_datasheet:
                                        if product_datasheet.value:
                                            uom = ''
                                            if product_datasheet.uom:
                                                uom = _(
                                                    dict(self.env[model_env].fields_get(
                                                        allfields=['uom'])['uom'][
                                                             'selection'])[
                                                        product_datasheet.uom])
                                            if product_datasheet.field_id and product_datasheet.field_id.type == 'number':
                                                info_display = str(round(float(product_datasheet.value), 2))
                                            elif product_datasheet.field_id and product_datasheet.field_id.type == 'boolean':
                                                if product_datasheet.value in ('0', 'True'):
                                                    info_display = _('Yes')
                                                else:
                                                    info_display = _('No')
                                            else:
                                                info_display = product_datasheet.value
                                            info_display += ' ' + uom if uom else ''
                                        else:
                                            info_display = '-'
                                        cell.value = info_display
                            else:
                                cell_value_splitted = cell_value.split('.')
                                model_obj = cell_value_splitted[1]
                                code = cell_value_splitted[2]
                                model_env = f'product.datasheet.{model_obj}'
                                product_datasheet = self.env[model_env].search([('code', '=', code)])
                                if product_datasheet:
                                    cell.value = product_datasheet.name

                                # IMAGES SECTION
                                if 'section' in cell_value:
                                    letter_column = list(string.ascii_uppercase)  # Array from A to Z
                                    columns_section = product_datasheet.column_ids.filtered(
                                        lambda m: m.group_id.export in [True])
                                    index_start = len(columns_section) + 1 if product_datasheet.column_ids else 2
                                    for product_image in self.product_id.image_ids.filtered(
                                            lambda m: m.section_id.id == product_datasheet.id):
                                        if product_image.image:
                                            buf_product_image = BytesIO(base64.b64decode(product_image.image))
                                            image_drawing = openpyxl.drawing.image.Image(buf_product_image)
                                            image_drawing.height = 300
                                            image_drawing.width = 300
                                            sheet_obj.add_image(image_drawing, letter_column[index_start] + str(
                                                cell.row + 1))  # Insert section image
                                            index_start += 1
                        else:
                            if cell_value in general_dict and any(st in cell_value for st in ['image', 'logo']):
                                if general_dict[cell_value]:
                                    buf_product_image = BytesIO(base64.b64decode(general_dict[cell_value]))
                                    image_drawing = openpyxl.drawing.image.Image(buf_product_image)
                                    image_drawing.height = 150
                                    image_drawing.width = 150
                                    sheet_obj.add_image(image_drawing, cell.coordinate)  # Insert product image
                                cell.value = None
                            else:
                                cell.value = general_dict[cell_value] if cell_value in general_dict else None
                    print(cell_value)

        # wb_obj.save('/home/file.xlsx')
        output = BytesIO()
        wb_obj.save(output)
        wb_obj.close()

        # output.seek(0)
        # data = output.read()
        data = output.getvalue()
        self.file_generated = base64.encodestring(data)

        return {
            'context': self.env.context,
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'product.datasheet.template.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }
