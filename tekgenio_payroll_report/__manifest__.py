# -*- coding: utf-8 -*-
{
    'name': "Tekgenio Payroll Report Generation",
    'version': '17.0.0.2',
    'description': """Gives the consolidate reports of all the employee payslip in monthly wise.""",
    'author': "TekGenio",
    'website': "https://tekgenio.com",
    'license': 'OPL-1',
    'depends': ['hr_contract', 'hr_payroll', 'report_xlsx'],
    'currency': 'USD',
    'price': 40,
    'images': ['static/description/banner.gif'],
    'data': [
        'security/ir.model.access.csv',
        'report/report.xml',
        'wizard/report_generation_wizard.xml',
        'view/hr_payroll_config.xml',
    ],
    'installable': True,
    'application': True,
    'license': 'LGPL-3',

}
