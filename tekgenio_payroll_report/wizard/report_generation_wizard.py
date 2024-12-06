from odoo import models, fields, api, _
from odoo.exceptions import ValidationError


class PayrollMonthlyReport(models.TransientModel):
    _name = "payroll.monthly.report"

    year = fields.Char("Year", default='2023')
    select_month = fields.Selection(
        [('1', 'January'), ('2', 'February'), ('3', 'March'), ('4', 'April'), ('5', 'May'), ('6', 'June'),
         ('7', 'July'), ('8', 'August'), ('9', 'September'), ('10', 'October'), ('11', 'November'), ('12', 'December')],
        string="Month", required=True)

    @api.onchange('year')
    def check_select_month(self):
        if self.year:
            if len(self.year) > 4:
                raise ValidationError("The data does not exist for entered year")

    def generate_report(self):
        return self.env.ref('tekgenio_payroll_report.payroll_report_xlsx').report_action(self)
