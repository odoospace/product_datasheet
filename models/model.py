from datetime import datetime, date

from odoo import models, fields, api, _


# TODO: write historic info

class Product(models.Model):
    _inherit = 'product.product'

    info_ids = fields.One2many('product.datasheet.info', 'product_id')


class Section(models.Model):
    _name = 'product.datasheet.section'
    _description = "Product Datasheet Section"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    active = fields.Boolean(default=True)
    timestamp = fields.Datetime(default=fields.Datetime.now)

    group_ids = fields.One2many('product.datasheet.group', 'section_id')

class Group(models.Model):
    _name = 'product.datasheet.group'
    _description = "Product Datasheet Group"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    timestamp = fields.Datetime(default=fields.Datetime.now)
    active = fields.Boolean(default=True)

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
            ("selection", "Selection"), #comma separated values or so
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
        ])

    info_ids = fields.One2many('product.datasheet.info', 'field_id')

class Info(models.Model):
    _name = 'product.datasheet.info'
    _description = "Product Datasheet Info"

    @api.depends('value')
    def _compute_value_name(self):
        for record in self:
            # This will be called every time the value field changes
            if len(record.value) > 50:
                record.value_display = record.value[:47] + '...'
            else:
                record.value_display = record.value


    
    field_id = fields.Many2one('product.datasheet.field', required=True)
    value = fields.Text(translatable=True)
    value_display = fields.Text(compute=_compute_value_name)
    timestamp = fields.Datetime(default=fields.Datetime.now)
    active = fields.Boolean(default=True)

    product_id = fields.Many2one('product.product')
    group_id = fields.Many2one('product.datasheet.group', required=True)
    # related fields
    section_id = fields.Many2one(related='group_id.section_id')
    uom = fields.Selection(related='field_id.uom')


class ProductProduct(models.Model):
    _inherit = 'product.product'

    datasheet_note = fields.Text()
    country_ids = fields.Many2many('res.country', 'product_ids')

class Country(models.Model):
    _inherit = 'res.country'

    product_ids = fields.Many2many('product.product', 'country_ids')