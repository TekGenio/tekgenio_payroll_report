from odoo import models
from odoo.exceptions import ValidationError
import calendar


class PayrollReportXlsx(models.AbstractModel):
    _name = 'report.tekgenio_payroll_report.payroll_report'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, partners):

        bold = workbook.add_format({"bold": True, 'align': 'center', 'border': 2, 'font_size': '11px'})
        title_alignment = workbook.add_format({"bold": True, 'align': 'center', 'font_size': '24px'})
        align_center = workbook.add_format({'align': 'center', 'border': 2, 'font_size': '11px'})
        value_assign_table = workbook.add_format({'align': 'center', 'border': 1, 'font_size': '11px'})
        sheet = workbook.add_worksheet('Payroll Report')
        sheet.set_row(0, 50)
        sheet.set_column(2, 30, 30)

        sheet.merge_range('B2:D2', 'Employees', bold)

        sheet.merge_range('B3:B4', 'Id', align_center)
        sheet.merge_range('C3:C4', 'Name', align_center)
        sheet.merge_range('D3:D4', 'Nationality', align_center)
        sheet.merge_range('E3:E4', 'Basic Salary', align_center)

        selected_month_year = self.env['payroll.monthly.report'].search([])
        entered_data_month = selected_month_year[-1].select_month
        convert_data_to_int = int(entered_data_month)
        entered_data_year = selected_month_year[-1].year
        convert_data_to_int_for_year = int(entered_data_year)

        month_name = calendar.month_name[convert_data_to_int]

        sheet.merge_range('D1:K1', f"Payroll Report for the Month of {month_name}, {convert_data_to_int_for_year}",
                          title_alignment)

        sql = """SELECT * FROM
        hr_payslip
        WHERE
        EXTRACT('month'
        from date_from) = %s and EXTRACT('year'
        from date_from)=%s"""

        vals = self.env.cr.execute(sql, (convert_data_to_int, convert_data_to_int_for_year))
        dict_of_vals = self.env.cr.dictfetchall()

        hr_payslip_record = []
        final_allowance = []
        final_deduction = []
        gross_pay_cell_with_alw_ded = []
        basic_salary_cell = 5

        if dict_of_vals:
            for value in dict_of_vals:
                # print(value)
                serach_record_in_hr_payslip = self.env['hr.payslip'].search([('id', '=', value.get('id'))])
                if serach_record_in_hr_payslip:
                    hr_payslip_record.append(serach_record_in_hr_payslip)

        if not hr_payslip_record:
            raise ValidationError("Records dose not exist for entered month and year")
        else:
            allowance = []
            allowance_name = []
            deduction = []
            deduction_name = []
            for partner in hr_payslip_record:

                if partner.line_ids:
                    for category in partner.line_ids:
                        # if category.code == 'BASIC':
                        #     sheet.write(row, 4, category.total, value_assign_table)
                        if category.category_id.name == 'Allowance' and not category.code == 'OT':
                            allowance.append(category)
                            if category.name not in allowance_name:
                                allowance_name.append(category.name)
                        elif category.category_id.name == 'Deduction':
                            deduction.append(category)
                            if category.name not in deduction_name:
                                deduction_name.append(category.name)
                            # deduction_name.append(category.name)
            set_allowance = set(allowance_name)
            print(set_allowance)
            if set_allowance:
                final_allowance.append(allowance_name)
                # final_allowance.append(list(set_allowance))
                print(set_allowance)
            set_deduction = set(deduction_name)
            if set_deduction:
                final_deduction.append(deduction_name)
                # final_deduction.append(list(set_deduction))

            if final_deduction and final_allowance:
                allowance_cell = basic_salary_cell

                for allow in final_allowance[0]:
                    allowance_elements = allowance_cell
                    sheet.write(3, allowance_elements, allow, align_center)

                    allowance_cell += 1
                sheet.merge_range(2, basic_salary_cell, 2, allowance_elements, '', align_center)
                sheet.write(2, basic_salary_cell, 'Allowances', align_center)
                ovrtime_cell = allowance_cell
                sheet.merge_range(2, ovrtime_cell, 3, ovrtime_cell, '', bold)
                sheet.write(2, ovrtime_cell, 'Overtime', align_center)
                gross_pay_cell = ovrtime_cell + 1
                gross_pay_cell_with_alw_ded.append(gross_pay_cell)
                sheet.merge_range(2, gross_pay_cell, 3, gross_pay_cell, '', bold)
                sheet.write(2, gross_pay_cell, 'Gross pay', align_center)
                sheet.merge_range(1, basic_salary_cell - 1, 1, gross_pay_cell, '', bold)
                sheet.write(1, basic_salary_cell - 1, 'Salaries & Allowances', bold)

                decut_cell_start = gross_pay_cell_with_alw_ded[0] + 1

                increment_cell_on_loop = decut_cell_start

                for ded in final_deduction[0]:
                    incremented_loop = increment_cell_on_loop
                    sheet.merge_range(2, incremented_loop, 3, incremented_loop, '', align_center)
                    sheet.write(2, incremented_loop, ded, align_center)

                    increment_cell_on_loop += 1
                total_deduct = increment_cell_on_loop

                sheet.merge_range(2, total_deduct, 3, total_deduct, '', align_center)
                sheet.write(2, total_deduct, 'Total Deduction', align_center)

                # sheet.write(row, total_deduct, sum_total_dedct, value_assign_table)

                sheet.merge_range(1, decut_cell_start, 1, total_deduct, '', bold)
                sheet.write(1, decut_cell_start, 'Deduction', bold)

                net_salrary = total_deduct + 1
                sheet.merge_range(1, net_salrary, 3, net_salrary, '', bold)
                sheet.write(1, net_salrary, 'Net Salary Payable', bold)

                remark = net_salrary + 1
                sheet.merge_range(1, remark, 3, remark, '', bold)
                sheet.write(1, remark, 'Remarks', bold)

            if not final_deduction and final_allowance:
                allowance_cell = basic_salary_cell

                for allow in final_allowance[0]:
                    allowance_elements = allowance_cell
                    sheet.write(3, allowance_elements, allow, align_center)
                    # sheet.write(row, allowance_elements, allowance.total, value_assign_table)
                    allowance_cell += 1
                sheet.merge_range(2, basic_salary_cell, 2, allowance_elements, '', align_center)
                sheet.write(2, basic_salary_cell, 'Allowances', align_center)
                ovrtime_cell = allowance_cell
                sheet.merge_range(2, ovrtime_cell, 3, ovrtime_cell, '', bold)
                sheet.write(2, ovrtime_cell, 'Overtime', align_center)
                # sheet.write(3, ovrtime_cell, 'Overtime', align_center)
                gross_pay_cell = ovrtime_cell + 1
                gross_pay_cell_with_alw_ded.append(gross_pay_cell)
                sheet.merge_range(2, gross_pay_cell, 3, gross_pay_cell, '', bold)
                sheet.write(2, gross_pay_cell, 'Gross pay', align_center)
                sheet.merge_range(1, basic_salary_cell - 1, 1, gross_pay_cell, '', bold)
                sheet.write(1, basic_salary_cell - 1, 'Salaries & Allowances', bold)

                net_salrary = gross_pay_cell + 1
                sheet.merge_range(1, net_salrary, 3, net_salrary, '', bold)
                sheet.write(1, net_salrary, 'Net Salary Payable', bold)

                remark = net_salrary + 1
                sheet.merge_range(1, remark, 3, remark, '', bold)
                sheet.write(1, remark, 'Remarks', bold)

            if final_deduction and not final_allowance:
                sheet.merge_range(2, basic_salary_cell, 3, basic_salary_cell, '', bold)
                sheet.write(3, basic_salary_cell, 'Overtime', align_center)
                sheet.merge_range(2, basic_salary_cell+1, 3, basic_salary_cell+1, '', bold)
                sheet.write(2, basic_salary_cell + 1, 'Gross pay', align_center)

                sheet.merge_range(1, 4, 1, basic_salary_cell + 1, '', bold)
                sheet.write(1, 4, 'Salaries & Allowances', bold)

                decut_cell_start = basic_salary_cell + 2

                increment_cell_on_loop = decut_cell_start

                for ded in final_deduction[0]:
                    incremented_loop = increment_cell_on_loop
                    sheet.merge_range(2, incremented_loop, 3, incremented_loop, '', align_center)
                    sheet.write(2, incremented_loop, ded, align_center)

                    increment_cell_on_loop += 1
                total_deduct = increment_cell_on_loop

                sheet.merge_range(2, total_deduct, 3, total_deduct, '', align_center)
                sheet.write(2, total_deduct, 'Total Deduction', align_center)

                sheet.merge_range(1, decut_cell_start, 1, total_deduct, '', bold)
                sheet.write(1, decut_cell_start, 'Deduction', bold)

                net_salrary = total_deduct + 1
                sheet.merge_range(1, net_salrary, 3, net_salrary, '', bold)
                sheet.write(1, net_salrary, 'Net Salary Payable', bold)

                remark = net_salrary + 1
                sheet.merge_range(1, remark, 3, remark, '', bold)
                sheet.write(1, remark, 'Remarks', bold)

            if not final_deduction and not final_allowance:
                sheet.merge_range(2, basic_salary_cell, 3, basic_salary_cell, '', bold)
                sheet.write(2, basic_salary_cell, 'Overtime', align_center)
                sheet.merge_range(2, basic_salary_cell+1, 3, basic_salary_cell+1, '', bold)
                sheet.write(2, basic_salary_cell + 1, 'Gross pay', align_center)
                sheet.merge_range(1, 4, 1, basic_salary_cell + 1, '', bold)
                sheet.write(1, 4, 'Salaries & Allowances', bold)

                net_salrary = basic_salary_cell + 2
                sheet.merge_range(1, net_salrary, 3, net_salrary, '', bold)
                sheet.write(1, net_salrary, 'Net Salary Payable', bold)

                remark = net_salrary + 1
                sheet.merge_range(1, remark, 3, remark, '', bold)
                sheet.write(1, remark, 'Remarks', bold)

            total_basic = []
            total_gross = []
            total_net = []
            total_deduction_column = []
            row = 4
            if hr_payslip_record:
                sum_column = {}
                sum_ded_column = {}
                for partner in hr_payslip_record:

                    sheet.write(row, 1, partner.employee_id.id, value_assign_table)

                    sheet.write(row, 2, partner.employee_id.name, value_assign_table)

                    if partner.employee_id.country_id.name:
                        sheet.write(row, 3, partner.employee_id.country_id.name, value_assign_table)

                    category_name = []
                    if partner.line_ids:

                        for category in partner.line_ids:
                            category_name.append(category.name)
                            if final_allowance:
                                if category.name in final_allowance[0]:
                                    index = final_allowance[0].index(category.name)

                                    sheet.write(row, index + 5, category.total, value_assign_table)

                                    if sum_column.get(index + 5, False):
                                        sum_column[index + 5] += category.total
                                    else:
                                        sum_column[index + 5] = category.total

                            if final_deduction:
                                if category.name in final_deduction[0]:
                                    index = final_deduction[0].index(category.name)
                                    if gross_pay_cell_with_alw_ded:

                                        sheet.write(row, index + gross_pay_cell_with_alw_ded[0] + 1, category.total,
                                                    value_assign_table)
                                        if sum_ded_column.get(index + gross_pay_cell_with_alw_ded[0] + 1, False):
                                            sum_ded_column[index + gross_pay_cell_with_alw_ded[0] + 1] += category.total
                                        else:
                                            sum_ded_column[index + gross_pay_cell_with_alw_ded[0] + 1] = category.total
                                    else:
                                        sheet.write(row, index + basic_salary_cell + 2, category.total,
                                                    value_assign_table)

                                        if sum_ded_column.get(index + basic_salary_cell + 2, False):
                                            sum_ded_column[index + basic_salary_cell + 2] += category.total
                                        else:
                                            sum_ded_column[index + basic_salary_cell + 2] = category.total

                        # for category in partner.line_ids.filtered(
                        #         lambda x: x.code in ['OT']):
                        #     if category.code == 'OT':
                        #
                        #         if final_allowance:
                        #             sheet.write(row, len(final_allowance[0]) + 5, category.total, value_assign_table)
                        #         else:
                        #             sheet.write(row, 5, 'NA', value_assign_table)
                        # else:
                        #
                        #     if final_allowance:
                        #         sheet.write(row, len(final_allowance[0]) + 5, 'NAA', value_assign_table)
                        #     else:
                        #         sheet.write(row, 5, 'NAb', value_assign_table)

                        for category in partner.line_ids.filtered(
                                lambda x: x.code in ['BASIC', 'GROSS', 'OT', 'NET', 'REMARK']):
                            if category.code == 'BASIC':
                                sheet.write(row, 4, category.total, value_assign_table)
                                total_basic.append(category.total)

                            elif category.code == 'OT':

                                if final_allowance:
                                    sheet.write(row, len(final_allowance[0]) + 5, category.total, value_assign_table)
                                else:
                                    sheet.write(row, 5, 'NA', value_assign_table)


                            elif category.code == 'GROSS':
                                if final_allowance:
                                    sheet.write(row, len(final_allowance[0]) + 6, category.total, value_assign_table)
                                    total_gross.append(category.total)
                                else:
                                    sheet.write(row, 6, category.total, value_assign_table)
                                    total_gross.append(category.total)

                            elif category.code == 'NET':
                                if final_allowance and final_deduction:
                                    sheet.write(row, len(final_deduction[0]) + len(final_allowance[0]) + 5 + 3,
                                                category.total,
                                                value_assign_table)
                                    total_net.append(category.total)
                                elif final_allowance and not final_deduction:
                                    sheet.write(row, len(final_allowance[0]) + 5 + 2, category.total,
                                                value_assign_table)
                                    total_net.append(category.total)
                                elif not final_allowance and final_deduction:
                                    sheet.write(row, len(final_deduction[0]) + 5 + 3, category.total,
                                                value_assign_table)
                                    total_net.append(category.total)
                                elif not final_allowance and not final_deduction:
                                    sheet.write(row, 5 + 2, category.total, value_assign_table)
                                    total_net.append(category.total)

                        # if partner note is there then mapping with remark
                        if partner.note:
                            if final_allowance and final_deduction:
                                sheet.write(row, len(final_deduction[0]) + len(final_allowance[0]) + 5 + 4,
                                            partner.note,
                                            value_assign_table)
                            elif final_allowance and not final_deduction:
                                sheet.write(row, len(final_allowance[0]) + 5 + 3, partner.note,
                                            value_assign_table)
                            elif not final_allowance and final_deduction:
                                sheet.write(row, len(final_deduction[0]) + 5 + 4, partner.note,
                                            value_assign_table)
                            elif not final_allowance and not final_deduction:
                                sheet.write(row, 5 + 3, partner.note, value_assign_table)

                            # sheet.write(row, len(final_deduction) + len(final_allowance)  + 5 + 4, category.total,
                            #             value_assign_table)

                        total_partner_deduction = []
                        for category in partner.line_ids.filtered(lambda x: x.category_id.name == 'Deduction'):
                            total_partner_deduction.append(category.total)

                            new_list = [abs(num) for num in total_partner_deduction]

                            if final_deduction and final_allowance:
                                sheet.write(row, len(final_deduction[0]) + len(final_allowance[0]) + 5 + 2,
                                            sum(new_list), value_assign_table)
                                total_deduction_column.append(category.total)

                            elif not final_allowance and final_deduction:
                                sheet.write(row, len(final_deduction[0]) + 5 + 2,
                                            sum(new_list), value_assign_table)
                                total_deduction_column.append(category.total)

                        if final_allowance:
                            for allowance_uniq in final_allowance[0]:
                                if allowance_uniq not in category_name:
                                    index = final_allowance[0].index(allowance_uniq)

                                    sheet.write(row, index + 5, 0, value_assign_table)

                        if final_deduction:
                            for deduction_uniq in final_deduction[0]:
                                if deduction_uniq not in category_name:
                                    index = final_deduction[0].index(deduction_uniq)
                                    if gross_pay_cell_with_alw_ded:

                                        sheet.write(row, index + gross_pay_cell_with_alw_ded[0] + 1, 0,
                                                    value_assign_table)
                                    else:
                                        sheet.write(row, index + basic_salary_cell + 2, 0,
                                                    value_assign_table)

                    row += 1

            sheet.write(row, 2, 'Total', bold)
            if total_basic:
                sheet.write(row, 4, sum(total_basic), bold)

            if sum_column:

                for key, value in sum_column.items():
                    sheet.write(row, int(key), value, bold)

            if sum_ded_column:

                for key, value in sum_ded_column.items():
                    sheet.write(row, int(key), value, bold)

            if total_gross:
                if final_allowance:
                    sheet.write(row, len(final_allowance[0]) + 6, sum(total_gross), bold)
                else:

                    sheet.write(row, 6, sum(total_gross), bold)

            if total_net:

                if final_allowance and final_deduction:
                    sheet.write(row, len(final_deduction[0]) + len(final_allowance[0]) + 5 + 3, sum(total_net), bold)

                elif final_allowance and not final_deduction:
                    sheet.write(row, len(final_allowance[0]) + 5 + 2, sum(total_net),
                                bold)

                elif not final_allowance and final_deduction:
                    sheet.write(row, len(final_deduction[0]) + 5 + 3, sum(total_net),
                                bold)

                elif not final_allowance and not final_deduction:
                    sheet.write(row, 5 + 2, sum(total_net), bold)

            if total_deduction_column:
                new_list = [abs(num) for num in total_deduction_column]

                if final_deduction and final_allowance:
                    sheet.write(row, len(final_deduction[0]) + len(final_allowance[0]) + 5 + 2,
                                sum(new_list), bold)


                elif not final_allowance and final_deduction:
                    sheet.write(row, len(final_deduction[0]) + 5 + 2,
                                sum(new_list), bold)
