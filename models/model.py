from odoo import models, fields, api, _
from datetime import datetime, date

# TODO: add config menu
# TODO: add forms and tree views
# TODO: write historic info  


class Product(models.Model):
    _inherit = 'product.product'

    group_ids = fields.One2many('product.datasheet.info', 'product_id')


class Section(models.Model):
    _name = 'product.datasheet.section'
    _description = "Product Datasheet Section"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    group_ids = fields.One2many('product.datasheet.group', 'section_id')
    timestamp = fields.Datetime(default=datetime.now)
    active = fields.Boolean()
    # user_ids = ...

    def write(self):
        pass


class Group(models.Model):
    _name = 'product.datasheet.group'
    _description = "Product Datasheet Group"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    timestamp = fields.Datetime(default=datetime.now)
    
    active = fields.Boolean()

    section_id = fields.Many2one('product.datasheet.section')
    # user_ids = ...

    def write(self):
        pass


class Info(models.Model):
    _name = 'product.datasheet.info'
    _description = "Product Datasheet Info"

    code = fields.Char(required=True)
    name = fields.Char(required=True, translate=True)
    value = fields.Char()
    timestamp = fields.Datetime(default=datetime.now)
    active = fields.Boolean()

    product_id = fields.Many2one('product.product', string='Product Datasheet info')
    group_id = fields.Many2one('product.datasheet.group')
    section_id = fields.Many2one(related='group_id.section_id')
    # user_ids = ...

    def write(self):
        pass

    
    




