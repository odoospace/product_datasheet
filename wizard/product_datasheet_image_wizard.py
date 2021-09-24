# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError, AccessError


class ProductDatasheetImageWizard(models.TransientModel):
    _name = 'product.datasheet.image.wizard'

    image = fields.Binary(string='Profile Image', required=True, attachment=False)
    product_ids = fields.Many2many('product.product', string='Products')

    @api.model
    def default_get(self, fields_list):
        product_ids = self.env['product.product'].browse(
            self.env.context['active_ids']
        )
        res = super(ProductDatasheetImageWizard, self).default_get(fields_list)
        res['product_ids'] = [(6, 0, product_ids.ids)]
        return res

    def action_add_image(self):
        for product in self.product_ids:
            product.product_tmpl_id.can_image_1024_be_zoomed = self.image
            product.product_tmpl_id.image_1024 = self.image
            product.product_tmpl_id.image_128 = self.image
            product.product_tmpl_id.image_1920 = self.image
            product.product_tmpl_id.image_256 = self.image
            product.product_tmpl_id.image_512 = self.image
