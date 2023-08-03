from odoo import models
import base64
import pandas as pd
import requests

class ProductCategoryUpdate(models.Model):
    _inherit = 'product.category'

    def update_categories_and_products(self):
        # Relative excel file in domain + /products_update/data/products.xlsx
        domain = self.env['ir.config_parameter'].sudo().get_param('web.base.url')
        excel_url = domain + '/products_update/static/src/products.xls'

        try:
            response = requests.get(excel_url)
            if response.status_code == 200:
                data = response.content

                # Read the Excel file using pandas
                df = pd.read_excel(data, engine='xlrd')

                category_obj = self.env['product.category']
                product_obj = self.env['product.template']
                account_obj = self.env['account.account']
                company_obj = self.env['res.company']

                for index, row in df.iterrows():
                    product_id = int(row['product_id'])
                    category_name = row['category']
                    account_income_code = row['account_income']
                    company_name = row['company']

                    company = company_obj.search([('name', '=', company_name)], limit=1)
                    
                    if company:
                        account_income = account_obj.search([('code', '=', account_income_code), ('company_id', '=', company.id)], limit=1)

                        category = category_obj.search([('name', '=', category_name)], limit=1)
                        if category and account_income:
                            if category.property_account_income_categ_id != account_income.id:
                                category.write({
                                    'property_account_income_categ_id': account_income.id,
                                })

                        if not category and account_income:
                            category = category_obj.create({
                                'name': category_name,
                                'property_valuation': 'real_time',
                                'property_account_income_categ_id': account_income.id,
                                'property_cost_method': 'standard',
                                'company_id': company.id,
                            })

                        product = product_obj.search([('id', '=', product_id), ('company_id', '=', company.id)], limit=1)

                        if product:
                            if product.categ_id != category.id:
                                product.write({
                                    'categ_id': category.id,
                                })
            else:
                print(f"Error: Unable to fetch Excel file from {excel_url}")
        except requests.exceptions.RequestException as e:
            print(f"Error: {e}")
