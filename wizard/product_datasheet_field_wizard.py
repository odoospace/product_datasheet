# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError, AccessError


class ProductDatasheetFieldWizard(models.TransientModel):
    _name = 'product.datasheet.field.wizard'

    action_type = fields.Selection(
        [
            ('add', _('Add')),
            ('remove', _('Remove')),
        ], string='Action Type', default='add', required=True, translate=True)
    product_ids = fields.Many2many('product.product', string='Products')
    group_id = fields.Many2one('product.datasheet.group', required=True)
    section_id = fields.Many2one(related='group_id.section_id')
    field_id = fields.Many2one('product.datasheet.field', required=True)
    group_after_id = fields.Many2one('product.datasheet.group')
    section_after_id = fields.Many2one(related='group_after_id.section_id')
    field_after_id = fields.Many2one('product.datasheet.field')

    @api.model
    def default_get(self, fields_list):
        product_ids = self.env['product.product'].browse(
            self.env.context['active_ids']
        )
        res = super(ProductDatasheetFieldWizard, self).default_get(fields_list)
        res['product_ids'] = [(6, 0, product_ids.ids)]
        return res

    def action_type_selected_confirm(self):
        print(f'You have chosen "{self.action_type}" Action Type with {len(self.product_ids)} Products')

        # Section, Group and Field Selected
        section_id = self.section_id
        group_id = self.group_id
        field_id = self.field_id

        # Place to put it
        section_after_id = self.section_after_id
        group_after_id = self.group_after_id
        field_after_id = self.field_after_id

        product_datasheet_info_obj = self.env['product.datasheet.info']

        if self.action_type == 'add':
            # if field_id and not field_id.related_field_product_id:
            #     raise UserError(_("The selected field is not related to the Product file"))

            for product in self.product_ids:
                product_datasheet_info = product_datasheet_info_obj.search(
                    [('section_id', '=', section_id.id), ('group_id', '=', group_id.id),
                     ('field_id', '=', field_id.id), ('product_id', '=', product.id)])
                if not product_datasheet_info:
                    product_datasheet_info_after = product_datasheet_info_obj.search(
                        [('section_id', '=', section_after_id.id), ('group_id', '=', group_after_id.id),
                         ('field_id', '=', field_after_id.id), ('product_id', '=', product.id)])
                    if product_datasheet_info_after:
                        sequence = product_datasheet_info_after.sequence + 1
                    else:
                        product_datasheet_info_group_after = product_datasheet_info_obj.search(
                            [('section_id', '=', section_after_id.id), ('group_id', '=', group_after_id.id),
                             ('product_id', '=', product.id)])
                        if product_datasheet_info_group_after:
                            sequence = product_datasheet_info_group_after[-1].sequence + 1
                        else:
                            product_datasheet_info_all_after = product_datasheet_info_obj.search([('product_id', '=', product.id)])
                            sequence = product_datasheet_info_all_after[-1].sequence + 1 if len(product_datasheet_info_all_after) != 0 else 0
                    product_datasheet_info.create({
                        'sequence': sequence,
                        'section_id': section_id.id,
                        'group_id': group_id.id,
                        'field_id': field_id.id,
                        'value': product.product_tmpl_id[field_id.related_field_product_id.name],
                        'product_id': product.id,
                    })
        elif self.action_type == 'remove':
            for product in self.product_ids:
                product_datasheet_info = product_datasheet_info_obj.search(
                    [('section_id', '=', section_id.id), ('group_id', '=', group_id.id),
                     ('field_id', '=', field_id.id), ('product_id', '=', product.id)])
                product_datasheet_info.unlink()
