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
    active = fields.Boolean()
    timestamp = fields.Datetime(default=fields.Datetime.now)

    group_ids = fields.One2many('product.datasheet.group', 'section_id')

    # user_ids = ...

    def write(self):
        pass


class Group(models.Model):
    _name = 'product.datasheet.group'
    _description = "Product Datasheet Group"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    timestamp = fields.Datetime(default=fields.Datetime.now)
    active = fields.Boolean()

    section_id = fields.Many2one('product.datasheet.section')

    # user_ids = ...

    def write(self):
        pass


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
            ("selection", "Selection"),
        ], required=True, translate=True)
    uom = fields.Selection(
        [
            ("gr", _("Gr")),
            ("cfu_g", _("Cfu/Gr")),
            ("m3", _("m3")),
            ("cm", _("Cm")),
            ("cm3", _("Cm3")),
            ("mm", _("Mm")),
            ("ud", _("Ud")),
            ("µg", _("µg")),
            ("box", _("Box")),
            ("mg", _("Mg")),
            ("kcal", _("Kcal")),
            ("KJ", _("KJ")),
            ("ud", _("Ud")),
            ("kg", _("Kg")),
            ("l", _("L")),
        ], translate=True)

    info_ids = fields.One2many('product.datasheet.info', 'field_id')


class Info(models.Model):
    _name = 'product.datasheet.info'
    _description = "Product Datasheet Info"

    value = fields.Char()
    timestamp = fields.Datetime(default=fields.Datetime.now)
    active = fields.Boolean()

    product_id = fields.Many2one('product.product')
    group_id = fields.Many2one('product.datasheet.group', required=True)
    field_id = fields.Many2one('product.datasheet.field')

    section_id = fields.Many2one(related='group_id.section_id')
    uom = fields.Selection(related='field_id.uom')

    # user_ids = ...

    # def write(self):
    #     pass
