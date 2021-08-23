# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError, AccessError


class ProductDatasheetTemplateWizard(models.TransientModel):
    _name = 'product.datasheet.template.wizard'

    template_id = fields.Many2one('product.datasheet.template', string='Template', required=True)
    product_id = fields.Many2one('product.product', string='Product')

    def action_download_excel(self):
        print(f'You have chosen {self.template_id.name}')
