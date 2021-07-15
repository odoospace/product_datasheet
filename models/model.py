from datetime import datetime, date

from odoo import models, fields, api, _
from io import BytesIO
import shortuuid
import xlsxwriter
import base64
import json


# TODO: write historic info
class Section(models.Model):
    _name = 'product.datasheet.section'
    _description = "Product Datasheet Section"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    active = fields.Boolean(default=True)
    timestamp = fields.Datetime(default=fields.Datetime.now)
    export = fields.Boolean('Is it exported?')

    group_ids = fields.One2many('product.datasheet.group', 'section_id')


class Group(models.Model):
    _name = 'product.datasheet.group'
    _rec_name = 'fullname'
    _description = "Product Datasheet Group"

    # @api.depends('name', 'section_id')
    def _get_fullname(self):
        for record in self:
            res = f'{record.name} ({record.section_id.name})'
            record.fullname = res

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    fullname = fields.Text(compute=_get_fullname, store=True)
    timestamp = fields.Datetime(default=fields.Datetime.now)
    active = fields.Boolean(default=True)
    export = fields.Boolean('Is it exported?')

    section_id = fields.Many2one('product.datasheet.section')


class Field(models.Model):
    _name = "product.datasheet.field"
    _description = "Product Datasheet Field"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    type = fields.Selection(
        [
            ("integer", "Integer"),
            ("string", "String"),
            ("html", "HTML"),
            ("selection", "Selection"),  # comma separated values or so
        ], required=True, translate=True)
    uom = fields.Selection(
        [
            ("gr", _("gr")),
            ("cfu_g", _("cfu/gr")),
            ("m3", _("m³")),
            ("cm", _("cm")),
            ("cm3", _("cm³")),
            ("mm", _("mm")),
            ("µg", _("µg")),
            ("box", _("caja")),
            ("mg", _("mg")),
            ("kcal", _("kcal")),
            ("KJ", _("kj")),
            ("ud", _("unidades")),
            ("kg", _("kg")),
            ("l", _("l")),
            ("min", _("min")),
            ("seg", _("seg")),
            ("day", _("day")),
            ("month", _("month")),
        ])
    export = fields.Boolean('Is it exported?')

    info_ids = fields.One2many('product.datasheet.info', 'field_id')


class Info(models.Model):
    _name = 'product.datasheet.info'
    _description = 'Product Datasheet Info'
    _order = 'sequence'

    @api.depends('value')
    def _compute_value_name(self):
        for record in self:
            # This will be called every time the value field changes
            if record.value and len(record.value) > 50:
                record.value_display = record.value[:47] + '...'
            else:
                record.value_display = record.value

    field_id = fields.Many2one('product.datasheet.field', required=True)
    value = fields.Text(translate=True)
    value_display = fields.Text(compute=_compute_value_name)
    timestamp = fields.Datetime(default=fields.Datetime.now)
    active = fields.Boolean(default=True)

    product_id = fields.Many2one('product.product')
    group_id = fields.Many2one('product.datasheet.group', required=True)
    # related fields
    group_name = fields.Char(string=_('Group'), related='group_id.name')
    section_id = fields.Many2one(related='group_id.section_id')
    uom = fields.Selection(related='field_id.uom')
    sequence = fields.Integer(default=1)


class ProductProduct(models.Model):
    _inherit = 'product.product'

    def filter_by_name(self):
        res = []
        if self.filter_field:
            res = [('field_id.name', 'ilike', self.filter_field)]
        return res

    info_ids = fields.One2many('product.datasheet.info', 'product_id', domain=filter_by_name)

    datasheet_note = fields.Text()
    country_ids = fields.Many2many('res.country', 'product_ids')

    # filters
    filter_field = fields.Char('Field')
    filter_section = fields.Many2one('product.datasheet.section')
    filter_group = fields.Many2one('product.datasheet.group')

    # add the field itself to onchange to trigger this method in edit mode too
    @api.onchange('filter_field')
    def onchange_filter_field(self):
        print('***')
        domain = []
        if self.filter_field:
            domain.append(('field_id.name', 'ilike', self.filter_field))
        if self.filter_section:
            domain.append(('field_id.section_id', 'ilike', self.filter_section))
        if self.filter_group:
            domain.append(('field_id.group_id', 'ilike', self.filter_group))
        res = {'domain': {'info_ids': domain}}
        print(res)
        self.info_ids = []
        return res

    def download_xlsx(self):
        # TODO: reload page to refresh attachments
        filename = f'{self.name}.xlsx'
        output = BytesIO()

        _info = {
            'code': 'DataSheet of Product',
            'created': datetime.now().strftime('%Y/%m/%d')
        }

        workbook = xlsxwriter.Workbook(output)

        product_name_format = workbook.add_format({
            'bold': True,
            'font_color': 'black',
            'bg_color': 'white',
            'border': 1
        })
        product_name_format.set_font_size(20)
        product_name_format.set_align('center')
        product_name_format.set_align('vcenter')
        black_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': 'black'
        })
        gray_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': 'gray'
        })
        normal_format = workbook.add_format({
            'font_color': 'black',
            'bg_color': 'white',
            'border': 1
        })
        normal_center_format = workbook.add_format({
            'font_color': 'black',
            'bg_color': 'white',
            'border': 1
        })
        normal_center_format.set_align('center')
        normal_center_format.set_align('vcenter')
        footer_format = workbook.add_format({
            'font_color': 'black',
            'bg_color': 'white'
        })

        # TAB NAME
        worksheet = workbook.add_worksheet(self.display_name)  # Tab with display_name of product

        # COMMENTS
        # info = _info
        # code = shortuuid.uuid()
        # info['worksheet'] = code
        #
        # worksheet.write_comment('A1', json.dumps(info))

        # INFO COMPANY HEADER
        if self.env.user.company_id and self.env.user.company_id.logo:
            buf_image_company = BytesIO(base64.b64decode(self.env.user.company_id.logo))
            worksheet.insert_image('A1', "image_company.png", {
                'image_data': buf_image_company,
                'x_scale': 0.03,
                'y_scale': 0.03
            })
        worksheet.set_row(0, 70)  # Set height of first row
        worksheet.set_column('A:A', 100)  # Set width column A
        worksheet.set_column('B:B', 50)  # Set width column B
        worksheet.set_column('C:C', 50)  # Set width column C
        worksheet.write(0, 0, self.name, product_name_format)
        worksheet.write(0, 1, datetime.now().strftime('%Y/%m/%d'), normal_center_format)

        # DATA OF SUPPLIER
        if self._context['lang'] == 'es_ES':
            title_data_supplier = ['Datos del Proveedor', 'Nombre Empresa', 'CIF', 'Registro Sanitario',
                                   'Dirección Fiscal', 'Contacto', 'Página Web']
        else:
            title_data_supplier = ['Supplier Data', 'Company Name', 'CIF', 'Health Register',
                                   'Fiscal Address', 'Contact', 'Website']

        row_start = 2
        row_title_supplier = row_start
        row_data_supplier = row_start + 1

        worksheet.write(row_title_supplier, 1, '', black_format)

        for title in title_data_supplier:
            if row_title_supplier == 2:
                format_title = black_format
            else:
                format_title = normal_format
            worksheet.write(row_title_supplier, 0, title, format_title)
            row_title_supplier += 1

        foodsfortomorrow_company = self.env['res.company'].search([('id', '=', 1)])
        if foodsfortomorrow_company:
            direction = foodsfortomorrow_company.street + ' - ' + foodsfortomorrow_company.zip + ', ' + foodsfortomorrow_company.state_id.display_name
            data_supplier = [foodsfortomorrow_company.name, foodsfortomorrow_company.vat,
                             foodsfortomorrow_company.company_registry, direction,
                             'calidad@heurafoods.com', foodsfortomorrow_company.website]
            for data in data_supplier:
                worksheet.write(row_data_supplier, 1, data, normal_format)
                row_data_supplier += 1

        # IMAGE PRODUCT
        if self.image_1920:
            row_data_product = row_start + 1
            buf_image_product = BytesIO(base64.b64decode(self.image_1920))
            worksheet.insert_image('C' + str(row_data_product), "image_product.png", {
                'image_data': buf_image_product,
                'x_scale': 0.3,
                'y_scale': 0.3
            })  # Insert image product

        # DATA OF PRODUCT
        row_start = row_title_supplier + 1  # Space between tables
        row_start_micro_analysis = 0
        enc_row_start_micro_analysis = False
        row_start_nut_information = 0
        enc_row_start_nut_information = False
        for section in self.env['product.datasheet.section'].search([('export', '=', True)]):
            is_header_section = True
            for group in section.group_ids.filtered(lambda m: m.export in [True]):
                is_header_group = True
                for info in self.env['product.datasheet.info'].search(
                        [('product_id', '=', self.id), ('section_id', '=', section.id), ('group_id', '=', group.id)]):
                    # HEADER NAME
                    if is_header_section:
                        # Space between tables
                        if row_start != row_title_supplier + 1:
                            row_start += 2
                        worksheet.write(row_start, 0, section.name, black_format)
                        worksheet.write(row_start, 1, '', black_format)

                        if section.name in ['Análisis Microbiológico', 'Información Nutricional']:
                            worksheet.write(row_start, 2, '', black_format)

                        is_header_section = False

                    # GROUP NAME
                    if is_header_group:
                        if (section.name not in ['Análisis Microbiológico', 'Modo Empleo',
                                                 'Información Nutricional']):
                            row_start += 1
                            worksheet.write(row_start, 0, group.name, gray_format)
                            worksheet.write(row_start, 1, '', gray_format)

                        # SUBGROUP ONLY CASES
                        if section.name == 'Alérgenos o intolerancias':
                            worksheet.write(row_start, 1, '', gray_format)
                            row_start += 1
                            worksheet.write(row_start, 1, 'Presencia - Puede contener (Trazas)', normal_center_format)
                        elif section.name == 'Análisis Microbiológico':
                            if group.name == 'Normal':
                                row_start += 1
                                worksheet.write(row_start, 2, 'Referencia laboratorio', normal_center_format)
                        elif section.name == 'Información Nutricional':
                            if group.name == 'Valores medios por 100g':
                                row_start += 1
                                worksheet.write(row_start, 1, 'Valores medios por 100gr de producto',
                                                normal_center_format)
                                worksheet.write(row_start, 2, 'CDR%', normal_center_format)

                        is_header_group = False

                    # FIELD NAME
                    if (info.field_id and info.field_id.export) and ((
                            section.name not in ['Análisis Microbiológico', 'Modo Empleo',
                                                 'Información Nutricional']) or (
                            section.name == 'Análisis Microbiológico' and group.name == 'Normal') or (
                            section.name == 'Modo Empleo' and group.name == '1') or (
                            section.name == 'Información Nutricional' and group.name == 'Valores medios por 100g')):
                        row_start += 1
                        worksheet.write(row_start, 0, info.field_id.name, normal_format)

                    def isfloat(value):
                        try:
                            float(value)
                            return True
                        except ValueError:
                            return False

                    # GET VALUE DISPLAY
                    if info.field_id and info.field_id.export:
                        if info.value_display and info.value_display != 'False':
                            uom = ''
                            if info.uom and group.name != 'Referencia Laboratorio':
                                uom = _(
                                    dict(self.env['product.datasheet.info'].fields_get(allfields=['uom'])['uom'][
                                             'selection'])[
                                        info.uom])
                            if isfloat(info.value_display):
                                info_display = str(round(float(info.value_display), 2)) + ' ' + uom
                            else:
                                info_display = info.value_display + ' ' + uom
                        else:
                            info_display = '-'

                        # PRINT VALUE DISPLAY WITH FORMAT COLUMN
                        if section.name == 'Alérgenos o intolerancias':
                            worksheet.write(row_start, 1, 'Sí - Sí' if info_display == 'True' else 'No - No', normal_format)
                        elif section.name == 'Análisis Microbiológico':
                            if group.name == 'Normal':
                                worksheet.write(row_start, 1, info_display, normal_format)
                                if not enc_row_start_micro_analysis:
                                    row_start_micro_analysis = row_start
                                    enc_row_start_micro_analysis = True
                            elif group.name == 'Referencia Laboratorio':
                                worksheet.write(row_start_micro_analysis, 2, info_display, normal_format)
                                row_start_micro_analysis += 1
                        elif section.name == 'Información Nutricional':
                            if group.name == 'Valores medios por 100g':
                                worksheet.write(row_start, 1, info_display, normal_format)
                                if not enc_row_start_nut_information:
                                    row_start_nut_information = row_start
                                    enc_row_start_nut_information = True
                            elif group.name == '%IR':
                                worksheet.write(row_start_nut_information, 2, info_display, normal_format)
                                row_start_nut_information += 1
                        else:
                            worksheet.write(row_start, 1, info_display, normal_format)

        # FOOTER
        worksheet.write(row_start + 2, 0, '*This is the footer text', footer_format)

        print('Saving excel...')
        workbook.close()
        output.seek(0)

        data = output.read()
        attachment = self.add_file_in_attachment(filename, data)
        # TODO: refresh page
        return {
            'type': 'ir.actions.act_url',
            'url': "web/content/?model=ir.attachment&id=" + str(attachment.id) +
                   f"&filename={self.name}.xlsx&field=datas&download=true&filename=" + attachment.name,
            'target': 'new',
        }

    def add_file_in_attachment(self, filename, output):
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'datas': base64.b64encode(output),
            'db_datas': filename,
            'res_model': 'product.product',
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        return attachment


class Country(models.Model):
    _inherit = 'res.country'

    product_ids = fields.Many2many('product.product', 'country_ids')
