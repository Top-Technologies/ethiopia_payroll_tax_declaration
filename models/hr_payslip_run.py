import base64
from collections import defaultdict
import io

from odoo import api, fields, models, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

# Salary rule codes mapped to the Schedule A report columns.
# Each report key maps to the primary rule code followed by common variations.
RULE_CODE_MAP = {
    "basic_salary": ("BASIC", "BASE"),
    "non_taxable_allowance": ("NT_ALLOW", "NON_TAXABLE_ALLOWANCE", "NTA", "NTAL", "TR"),
    "taxable_allowance": ("TA_ALLOW", "TAXABLE_ALLOWANCE", "TA", "TAAL", "SALES_COMMISSION"),
    "overtime": ("OT", "OVERTIME", "OT1.5", "OT1.75", "OT2.00", "OT2.5"),
    "taxable_income": ("GROSS_TAXABLE", "TAXABLE", "TAXABLE_GROSS", "TAXG"),
    "gross_salary": ("GROSS", "GROSS_SALARY"),
    "income_tax": ("IIT", "IT", "TAX", "PAYE", "INCOME_TAX", "INCTAX"),
    "employer_pension": ("PENS_11", "PENSION_11", "PENSION_EMPLOYER", "PENS_EMP", "PEN11"),
    "employee_pension": ("PENS_7", "PENSION_7", "PENSION_EMPLOYEE", "PENS_EE", "PEN7"),
    "other_deductions": ("OTHER_DED", "OTHER_DEDUCTION", "DEDUCTION", "SLOF", "CostS", "PENALITY"),
    "total_deductions": ("TOTAL_DED", "TOTAL_DEDUCTION", "DEDUCTIONS", "TOTALD"),
    "net_pay": ("NET", "NET_SALARY"),
}

DEDUCTION_KEYS = {
    "income_tax",
    "employer_pension",
    "employee_pension",
    "other_deductions",
    "total_deductions",
}

VALID_SLIP_STATES = ("done", "paid", "verify")


class HrPayslipRun(models.Model):
    _inherit = "hr.payslip.run"

    def action_print_schedule_a_tax_report(self):
        """Launch the QWeb PDF report for the MoR Schedule A declaration."""
        self.ensure_one()
        return self.env.ref(
            "ethiopia_payroll_tax_declaration.action_report_mor_schedule_a"
        ).report_action(self)

    def action_open_erca_wizard(self):
        """Open the ERCA Schedule A Tax & Pension export wizard dialog."""
        self.ensure_one()
        return {
            "name": _("ERCA Tax Declaration - Export to Excel"),
            "type": "ir.actions.act_window",
            "res_model": "schedule.a.export.wizard",
            "view_mode": "form",
            "target": "new",
            "context": {
                "default_payslip_run_id": self.id,
            },
        }

    def action_export_schedule_a_excel(self):
        """Directly generate and download the ERCA Schedule A Excel declaration."""
        self.ensure_one()
        xlsx_bytes = self._generate_schedule_a_excel()
        clean_name = (
            (self.name or "Batch")
            .replace("/", "_")
            .replace("\\", "_")
        )
        filename = f"ERCA_Schedule_A_{clean_name}.xlsx"
        attachment = self.env["ir.attachment"].create({
            "name": filename,
            "type": "binary",
            "datas": base64.b64encode(xlsx_bytes),
            "res_model": self._name,
            "res_id": self.id,
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })
        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/{attachment.id}?download=true",
            "target": "self",
        }

    action_export_erca_excel = action_export_schedule_a_excel

    def _get_schedule_a_data(self):
        """Prepare all data required to render the Schedule A report.

        Iterates over `slip_ids` filtered by state in ['done', 'paid', 'verify']
        and maps salary rule totals into an employee dictionary.

        :return: dict with company header info, employee lines and totals.
        """
        self.ensure_one()
        company = self.company_id or self.env.company
        partner = company.partner_id

        slips = self.slip_ids.filtered(
            lambda slip: slip.state in VALID_SLIP_STATES
        )
        sort_key = lambda s: (
            s.employee_id.name or "" if s.employee_id else ""
        ).lower()

        if hasattr(slips, "sorted"):
            slips = slips.sorted(sort_key)
        else:
            slips = sorted(slips, key=sort_key)

        lines = [self._get_schedule_a_line_vals(slip) for slip in slips]
        for idx, line in enumerate(lines, 1):
            line["sn"] = idx

        totals = {key: 0.0 for key in RULE_CODE_MAP}
        for line in lines:
            for key in RULE_CODE_MAP:
                totals[key] += line.get(key, 0.0)

        for key in totals:
            totals[key] = round(totals[key], 2)
        totals["total_pension"] = round(
            totals["employee_pension"] + totals["employer_pension"], 2
        )

        company_tin = (
            getattr(company, "vat", "")
            or getattr(company, "tin", "")
            or getattr(partner, "vat", "")
            or ""
        )
        company_phone = (
            partner.phone
            or partner.mobile
            or getattr(company, "phone", "")
            or ""
        )
        company_region = partner.state_id.name or ""
        company_zone = partner.city or partner.street2 or ""

        return {
            "company": company,
            "company_name": company.name or "",
            "company_tin": company_tin,
            "company_phone": company_phone,
            "company_region": company_region,
            "company_zone": company_zone,
            "tax_period_from": self.date_start,
            "tax_period_to": self.date_end,
            "period_name": self.name or "",
            "currency_symbol": company.currency_id.symbol or "ETB",
            "lines": lines,
            "employees": lines,
            "employee_count": len(lines),
            "totals": totals,
        }

    def _get_schedule_a_line_vals(self, slip):
        """Map one payslip's salary rule totals into a report line dict."""
        employee = slip.employee_id
        amounts = defaultdict(float)

        for line in slip.line_ids:
            code = (line.code or "").strip().upper()
            for report_key, codes in RULE_CODE_MAP.items():
                if code in codes:
                    amount = (
                        abs(line.total)
                        if report_key in DEDUCTION_KEYS
                        else line.total
                    )
                    amounts[report_key] += amount
                    break

        # Fallback: derive taxable income when no GROSS_TAXABLE rule exists.
        if not amounts["taxable_income"]:
            amounts["taxable_income"] = (
                amounts["basic_salary"]
                + amounts["taxable_allowance"]
                + amounts["overtime"]
            )

        # Fallback: derive gross salary when no GROSS rule exists.
        if not amounts["gross_salary"]:
            amounts["gross_salary"] = (
                amounts["taxable_income"]
                + amounts["non_taxable_allowance"]
            )

        # Fallback: derive total deductions when no TOTAL_DED rule exists.
        if not amounts["total_deductions"]:
            amounts["total_deductions"] = (
                amounts["income_tax"]
                + amounts["employee_pension"]
                + amounts["other_deductions"]
            )

        # Fallback: derive net pay when no NET rule exists.
        if not amounts["net_pay"]:
            amounts["net_pay"] = (
                amounts["gross_salary"]
                - amounts["total_deductions"]
            )

        tin = self._get_employee_tin(employee)
        start_date = (
            getattr(employee, "first_contract_date", False)
            or (slip.contract_id and slip.contract_id.date_start)
            or (employee.create_date and employee.create_date.date())
            or False
        )

        return {
            "employee_id": employee.id if employee else False,
            "employee_tin": tin,
            "tin": tin,
            "employee_name": employee.name or "",
            "name": employee.name or "",
            "start_date": start_date,
            "hire_date": start_date,
            "basic_salary": round(amounts["basic_salary"], 2),
            "non_taxable_allowance": round(amounts["non_taxable_allowance"], 2),
            "taxable_allowance": round(amounts["taxable_allowance"], 2),
            "overtime": round(amounts["overtime"], 2),
            "taxable_income": round(amounts["taxable_income"], 2),
            "gross_salary": round(amounts["gross_salary"], 2),
            "income_tax": round(amounts["income_tax"], 2),
            "employer_pension": round(amounts["employer_pension"], 2),
            "employee_pension": round(amounts["employee_pension"], 2),
            "other_deductions": round(amounts["other_deductions"], 2),
            "total_deductions": round(amounts["total_deductions"], 2),
            "net_pay": round(amounts["net_pay"], 2),
            # Rule code aliases:
            "BASIC": round(amounts["basic_salary"], 2),
            "NT_ALLOW": round(amounts["non_taxable_allowance"], 2),
            "TA_ALLOW": round(amounts["taxable_allowance"], 2),
            "OT": round(amounts["overtime"], 2),
            "GROSS_TAXABLE": round(amounts["taxable_income"], 2),
            "GROSS": round(amounts["gross_salary"], 2),
            "IIT": round(amounts["income_tax"], 2),
            "PENS_11": round(amounts["employer_pension"], 2),
            "PENS_7": round(amounts["employee_pension"], 2),
            "OTHER_DED": round(amounts["other_deductions"], 2),
            "TOTAL_DED": round(amounts["total_deductions"], 2),
            "NET": round(amounts["net_pay"], 2),
        }

    def _get_employee_tin(self, employee):
        """Retrieve TIN from employee_id.tin, employee_id.tin_no, or employee_id.identification_id."""
        if not employee:
            return ""
        if hasattr(employee, "_fields") and "tin" in employee._fields and employee.tin:
            return str(employee.tin).strip()
        tin = getattr(employee, "tin", None)
        if tin:
            return str(tin).strip()
        tin_no = getattr(employee, "tin_no", None)
        if tin_no:
            return str(tin_no).strip()
        ident = getattr(employee, "identification_id", None)
        if ident:
            return str(ident).strip()
        work_contact = getattr(employee, "work_contact_id", None)
        if work_contact and getattr(work_contact, "vat", None):
            return str(work_contact.vat).strip()
        return ""

    def _generate_schedule_a_excel(self):
        """Generate the official Ethiopian Ministry of Revenues (MoR / ERCA)
        Schedule A Employment Income Tax and Pension Contribution Declaration
        in XLSX format.

        :return: bytes of the XLSX file.
        """
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(
                _("The 'xlsxwriter' Python library is required to export to Excel.")
            )

        data = self._get_schedule_a_data()
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        worksheet = workbook.add_worksheet("ERCA Schedule A")

        # Page setup: A4 Landscape, fit all columns to 1 page wide
        worksheet.set_landscape()
        worksheet.set_paper(9)
        worksheet.fit_to_pages(1, 0)
        worksheet.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)

        # Formats
        title_fmt = workbook.add_format({
            "bold": True,
            "font_size": 13,
            "align": "center",
            "valign": "vcenter",
        })
        subtitle_fmt = workbook.add_format({
            "bold": True,
            "font_size": 10.5,
            "align": "center",
            "valign": "vcenter",
        })
        subtext_fmt = workbook.add_format({
            "font_size": 8.5,
            "italic": True,
            "align": "center",
            "valign": "vcenter",
        })
        meta_label_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "bg_color": "#F2F2F2",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        })
        meta_val_fmt = workbook.add_format({
            "font_size": 8.5,
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        })
        th_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8,
            "bg_color": "#E7E6E6",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        td_text_fmt = workbook.add_format({
            "font_size": 8,
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        })
        td_center_fmt = workbook.add_format({
            "font_size": 8,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        td_num_fmt = workbook.add_format({
            "font_size": 8,
            "border": 1,
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0.00",
        })
        total_label_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "bg_color": "#E2EFDA",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        total_num_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "bg_color": "#E2EFDA",
            "border": 1,
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0.00",
        })
        sec3_header_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "bg_color": "#D9E1F2",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        sec3_label_fmt = workbook.add_format({
            "font_size": 8.5,
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        })
        sec3_val_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "border": 1,
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0.00",
        })
        sec3_val_int_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "border": 1,
            "align": "right",
            "valign": "vcenter",
        })
        sig_title_fmt = workbook.add_format({
            "bold": True,
            "font_size": 8.5,
            "align": "left",
        })
        sig_text_fmt = workbook.add_format({
            "font_size": 8,
            "align": "left",
        })

        # Columns definition (Header Title, Width)
        headers = [
            ("S/N", 5),
            ("TIN", 13),
            ("Employee Name", 24),
            ("Start Date", 11),
            ("Basic Salary", 13),
            ("Non-Taxable\nAllowance", 13),
            ("Taxable\nAllowance", 13),
            ("Overtime", 11),
            ("Taxable\nIncome", 13),
            ("Gross Salary", 13),
            ("Income Tax\nWithheld", 13),
            ("Pension 11%\n(Employer)", 13),
            ("Pension 7%\n(Employee)", 13),
            ("Other\nDeductions", 12),
            ("Total\nDeductions", 13),
            ("Net Pay", 13),
            ("Signature", 11),
        ]
        for col_idx, (_header_name, col_width) in enumerate(headers):
            worksheet.set_column(col_idx, col_idx, col_width)

        # Title Block
        worksheet.merge_range(
            0,
            0,
            0,
            16,
            "Federal Democratic Republic of Ethiopia - Ministry of Revenues (MoR) / ERCA",
            title_fmt,
        )
        worksheet.merge_range(
            1,
            0,
            1,
            16,
            "Schedule A - Employment Income Tax and Pension Contribution Declaration",
            subtitle_fmt,
        )
        worksheet.merge_range(
            2,
            0,
            2,
            16,
            "(Monthly Declaration of Employment Income Tax Withheld and Pension Contributions)",
            subtext_fmt,
        )

        # Meta Information Block
        tax_period_str = (
            f"{data['tax_period_from'] or '-'} to {data['tax_period_to'] or '-'}"
        )
        worksheet.write(4, 0, "Employer Name:", meta_label_fmt)
        worksheet.merge_range(4, 1, 4, 3, data["company_name"], meta_val_fmt)
        worksheet.write(4, 4, "Company TIN:", meta_label_fmt)
        worksheet.merge_range(4, 5, 4, 6, data["company_tin"], meta_val_fmt)
        worksheet.write(4, 7, "Tax Period:", meta_label_fmt)
        worksheet.merge_range(4, 8, 4, 10, tax_period_str, meta_val_fmt)
        worksheet.write(4, 11, "Phone:", meta_label_fmt)
        worksheet.merge_range(4, 12, 4, 13, data["company_phone"], meta_val_fmt)
        worksheet.write(4, 14, "Region:", meta_label_fmt)
        worksheet.merge_range(4, 15, 4, 16, data["company_region"], meta_val_fmt)

        worksheet.write(5, 0, "Zone / Woreda:", meta_label_fmt)
        worksheet.merge_range(5, 1, 5, 3, data["company_zone"], meta_val_fmt)
        worksheet.write(5, 4, "Batch Reference:", meta_label_fmt)
        worksheet.merge_range(5, 5, 5, 6, data["period_name"], meta_val_fmt)
        worksheet.write(5, 7, "Currency:", meta_label_fmt)
        worksheet.merge_range(5, 8, 5, 10, data["currency_symbol"], meta_val_fmt)
        worksheet.write(5, 11, "Employees:", meta_label_fmt)
        worksheet.merge_range(
            5, 12, 5, 13, str(data["employee_count"]), meta_val_fmt
        )
        worksheet.write(5, 14, "", meta_val_fmt)
        worksheet.merge_range(5, 15, 5, 16, "", meta_val_fmt)

        # Section 2 Headers
        header_row = 7
        worksheet.set_row(header_row, 28)
        for col_idx, (header_name, _col_width) in enumerate(headers):
            worksheet.write(header_row, col_idx, header_name, th_fmt)

        # Data Rows
        start_data_row = 8
        current_row = start_data_row
        for idx, line in enumerate(data["lines"], 1):
            worksheet.set_row(current_row, 18)
            start_date_str = (
                str(line["start_date"]) if line.get("start_date") else "-"
            )
            worksheet.write(current_row, 0, line.get("sn", idx), td_center_fmt)
            worksheet.write(
                current_row, 1, line["employee_tin"] or "-", td_center_fmt
            )
            worksheet.write(current_row, 2, line["employee_name"], td_text_fmt)
            worksheet.write(current_row, 3, start_date_str, td_center_fmt)
            worksheet.write(current_row, 4, line["basic_salary"], td_num_fmt)
            worksheet.write(
                current_row, 5, line["non_taxable_allowance"], td_num_fmt
            )
            worksheet.write(
                current_row, 6, line["taxable_allowance"], td_num_fmt
            )
            worksheet.write(current_row, 7, line["overtime"], td_num_fmt)
            worksheet.write(current_row, 8, line["taxable_income"], td_num_fmt)
            worksheet.write(current_row, 9, line["gross_salary"], td_num_fmt)
            worksheet.write(current_row, 10, line["income_tax"], td_num_fmt)
            worksheet.write(
                current_row, 11, line["employer_pension"], td_num_fmt
            )
            worksheet.write(
                current_row, 12, line["employee_pension"], td_num_fmt
            )
            worksheet.write(
                current_row, 13, line["other_deductions"], td_num_fmt
            )
            worksheet.write(
                current_row, 14, line["total_deductions"], td_num_fmt
            )
            worksheet.write(current_row, 15, line["net_pay"], td_num_fmt)
            worksheet.write(current_row, 16, "", td_center_fmt)
            current_row += 1

        # Totals Row
        worksheet.set_row(current_row, 20)
        worksheet.merge_range(
            current_row, 0, current_row, 3, "TOTAL", total_label_fmt
        )
        amount_keys = [
            "basic_salary",
            "non_taxable_allowance",
            "taxable_allowance",
            "overtime",
            "taxable_income",
            "gross_salary",
            "income_tax",
            "employer_pension",
            "employee_pension",
            "other_deductions",
            "total_deductions",
            "net_pay",
        ]
        for offset, key in enumerate(amount_keys):
            col_idx = 4 + offset
            total_val = data["totals"].get(key, 0.0)
            worksheet.write(current_row, col_idx, total_val, total_num_fmt)
        worksheet.write(current_row, 16, "", total_label_fmt)
        current_row += 1

        # Section 3: Summary of Declaration
        current_row += 1
        worksheet.merge_range(
            current_row,
            0,
            current_row,
            4,
            "Section 3 - Summary of Declaration",
            sec3_header_fmt,
        )
        current_row += 1
        worksheet.merge_range(
            current_row,
            0,
            current_row,
            3,
            "Total Number of Employees",
            sec3_label_fmt,
        )
        worksheet.write(
            current_row, 4, data["employee_count"], sec3_val_int_fmt
        )

        current_row += 1
        worksheet.merge_range(
            current_row,
            0,
            current_row,
            3,
            f"Total Taxable Income ({data['currency_symbol']})",
            sec3_label_fmt,
        )
        worksheet.write(
            current_row, 4, data["totals"]["taxable_income"], sec3_val_fmt
        )

        current_row += 1
        worksheet.merge_range(
            current_row,
            0,
            current_row,
            3,
            f"Total Income Tax Withheld ({data['currency_symbol']})",
            sec3_label_fmt,
        )
        worksheet.write(
            current_row, 4, data["totals"]["income_tax"], sec3_val_fmt
        )

        current_row += 1
        worksheet.merge_range(
            current_row,
            0,
            current_row,
            3,
            f"Total Pension Contribution (Employee 7% + Employer 11%) ({data['currency_symbol']})",
            sec3_label_fmt,
        )
        worksheet.write(
            current_row, 4, data["totals"]["total_pension"], sec3_val_fmt
        )

        # Section 4: Sign-offs
        current_row += 2
        worksheet.write(
            current_row,
            1,
            "Prepared By (Accountant / HR Officer):",
            sig_title_fmt,
        )
        worksheet.write(
            current_row,
            9,
            "Taxpayer Certification (Authorized Representative):",
            sig_title_fmt,
        )
        current_row += 1
        worksheet.write(
            current_row,
            9,
            "I hereby declare that the information given on this declaration is true, complete and correct.",
            subtext_fmt,
        )
        current_row += 1
        worksheet.write(
            current_row,
            1,
            "Name: __________________________________________",
            sig_text_fmt,
        )
        worksheet.write(
            current_row,
            9,
            "Name: __________________________________________",
            sig_text_fmt,
        )
        current_row += 1
        worksheet.write(
            current_row,
            1,
            "Title: ___________________________________________",
            sig_text_fmt,
        )
        worksheet.write(
            current_row,
            9,
            "Title: ___________________________________________",
            sig_text_fmt,
        )
        current_row += 1
        worksheet.write(
            current_row,
            1,
            "Signature: ____________________ Date: ____________",
            sig_text_fmt,
        )
        worksheet.write(
            current_row,
            9,
            "Signature: ____________________ Date: ____________",
            sig_text_fmt,
        )

        # ================================================================
        # SHEET 2: ERCA Pension Declaration Form P1/2011 (Page 4)
        # ================================================================
        self._build_pension_sheet(workbook, data)

        workbook.close()
        return output.getvalue()

    def _build_pension_sheet(self, workbook, data):
        """Build Sheet 2: Official ERCA Pension Contribution Declaration
        matching the printed P1/2011 form layout (page 4 of the declaration).

        Sections:
          ክፍል -1 : Employer information header
          ክፍል -2 : Employee pension breakdown table (7% / 11% / 18%)
          ክፍል -3 : Summary lines (10-50)
          ክፍል -5 : Declaration signature block
        """
        ws = workbook.add_worksheet("ERCA Pension (P1-2011)")
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)

        # ── Shared formats ────────────────────────────────────────────────
        bold_center = workbook.add_format({
            "bold": True, "font_size": 10, "align": "center",
            "valign": "vcenter", "text_wrap": True,
        })
        logo_fmt = workbook.add_format({
            "bold": True, "font_size": 8.5, "align": "left",
            "valign": "vcenter", "text_wrap": True,
        })
        title_fmt = workbook.add_format({
            "bold": True, "font_size": 12, "align": "center",
            "valign": "vcenter",
        })
        sec_hdr_fmt = workbook.add_format({
            "bold": True, "font_size": 9, "bg_color": "#D9E1F2",
            "border": 1, "align": "center", "valign": "vcenter",
            "text_wrap": True,
        })
        meta_lbl_fmt = workbook.add_format({
            "bold": True, "font_size": 8.5, "bg_color": "#F2F2F2",
            "border": 1, "align": "left", "valign": "vcenter",
        })
        meta_val_fmt = workbook.add_format({
            "font_size": 8.5, "border": 1, "align": "left",
            "valign": "vcenter",
        })
        th_fmt = workbook.add_format({
            "bold": True, "font_size": 8, "bg_color": "#E7E6E6",
            "border": 1, "align": "center", "valign": "vcenter",
            "text_wrap": True,
        })
        td_center = workbook.add_format({
            "font_size": 8, "border": 1, "align": "center",
            "valign": "vcenter",
        })
        td_left = workbook.add_format({
            "font_size": 8, "border": 1, "align": "left",
            "valign": "vcenter",
        })
        td_num = workbook.add_format({
            "font_size": 8, "border": 1, "align": "right",
            "valign": "vcenter", "num_format": "#,##0.00",
        })
        total_lbl = workbook.add_format({
            "bold": True, "font_size": 8.5, "bg_color": "#E2EFDA",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        total_num = workbook.add_format({
            "bold": True, "font_size": 8.5, "bg_color": "#E2EFDA",
            "border": 1, "align": "right", "valign": "vcenter",
            "num_format": "#,##0.00",
        })
        sec3_lbl = workbook.add_format({
            "font_size": 8.5, "border": 1, "align": "left",
            "valign": "vcenter",
        })
        sec3_num = workbook.add_format({
            "bold": True, "font_size": 8.5, "border": 1,
            "align": "right", "valign": "vcenter",
            "num_format": "#,##0.00",
        })
        sec3_int = workbook.add_format({
            "bold": True, "font_size": 8.5, "border": 1,
            "align": "right", "valign": "vcenter",
        })
        sig_title_fmt = workbook.add_format({
            "bold": True, "font_size": 8.5, "align": "left",
        })
        sig_text_fmt = workbook.add_format({
            "font_size": 8, "align": "left",
        })
        italic_fmt = workbook.add_format({
            "italic": True, "font_size": 7.5, "align": "center",
        })

        # ── Column widths ─────────────────────────────────────────────────
        # S/N | Name | Start Date | Basic Salary | 7% | 11% | 18% | blank
        col_widths = [5, 28, 12, 15, 15, 15, 15, 10]
        for i, w in enumerate(col_widths):
            ws.set_column(i, i, w)

        # ── Row 0-1: Institution header ───────────────────────────────────
        ws.merge_range(
            0, 0, 1, 2,
            "Federal Democratic Republic of Ethiopia\nEthiopian Revenues and Customs Authority (ERCA)",
            logo_fmt,
        )
        ws.merge_range(
            0, 3, 1, 4,
            (
                "ቀ.ዕ.ቁ 2/2003 ቀ.ቁ.ቁ 715(2003) የሠራተኞች ጡረታ\n"
                "ወጪ ድርሻ መዋጮ ማስታወቂያ ቅጽ"
            ),
            logo_fmt,
        )
        ws.merge_range(0, 5, 1, 7, "", logo_fmt)

        # ── Row 2: Main title ─────────────────────────────────────────────
        ws.set_row(2, 20)
        ws.merge_range(
            2, 0, 2, 7,
            "ክፍል -1  የጡረታ መዋጮ ማስታወቂያ ደርጅት ዝርዝር መረጃ",
            title_fmt,
        )

        # ── Row 3-5: ክፍል -1 Employer meta ───────────────────────────────
        ws.set_row(3, 18)
        ws.set_row(4, 18)
        ws.set_row(5, 18)

        tax_period_str = (
            f"{data['tax_period_from'] or '-'} - {data['tax_period_to'] or '-'}"
        )

        # Row 3
        ws.write(3, 0, "1. የድርጅቱ ስም:-ደና ግብኝኝ  ታደስ (አንበሳ ቢራ ወ/አከፋፋይ)", meta_lbl_fmt)
        ws.merge_range(3, 1, 3, 3, data["company_name"], meta_val_fmt)
        ws.write(3, 4, "4. የግብር ሂሳብ ቁጥር (ወ-መስከርም)", meta_lbl_fmt)
        ws.merge_range(3, 5, 3, 7, data["company_tin"], meta_val_fmt)

        # Row 4
        ws.write(4, 0, "2a. ክልል", meta_lbl_fmt)
        ws.write(4, 1, data["company_region"], meta_val_fmt)
        ws.write(4, 2, "2b. ዞን/ክፍለ ከተማ", meta_lbl_fmt)
        ws.write(4, 3, "", meta_val_fmt)
        ws.write(4, 4, "3. የቁጥር ዝርዝር ጊዜ / ዓ.ም", meta_lbl_fmt)
        ws.merge_range(4, 5, 4, 7, tax_period_str, meta_val_fmt)

        # Row 5
        ws.write(5, 0, "2c ወረዳ", meta_lbl_fmt)
        ws.write(5, 1, data["company_zone"], meta_val_fmt)
        ws.write(5, 2, "2d ቀበሌ/ንብረ ማህበር:- አዳማ", meta_lbl_fmt)
        ws.write(5, 3, "", meta_val_fmt)
        ws.write(5, 4, "2e. የቤት ቁጥር", meta_lbl_fmt)
        ws.write(5, 5, "", meta_val_fmt)
        ws.write(5, 6, "7. ፋክስ ቁጥር", meta_lbl_fmt)
        ws.write(5, 7, data["company_phone"], meta_val_fmt)

        # ── Row 7: ክፍል -2 Title ──────────────────────────────────────────
        ws.set_row(7, 18)
        ws.merge_range(
            7, 0, 7, 7,
            "ክፍል -2  የሚስታወቂያ ዝርዝር መረጃ",
            sec_hdr_fmt,
        )

        # ── Row 8: Table column headers ───────────────────────────────────
        ws.set_row(8, 32)
        headers_p4 = [
            ("ሀ) የቁጥር\nፍሎ+ፎ\nፕ) ት/ቄ የፕ7ዓC ሃ4ይ", 5),
            ("ሁ) ሠራተኛው ስም፡ የእናት ስም እና\nያስ ስም", 28),
            ("ሐ) የተቀጠሩበት ቀን\n/ቀን/ወር/ዓ.ም/", 12),
            ("ሠ) የወር ደመወዝ /ብር/", 15),
            ("ረ) የሠራተኛው መዋጮ መጠን\nወጠን  7--%/ብር/", 15),
            ("ሰ) የአሰሪው መዋጮ መጠን\n11--%/ብር/", 15),
            ("ሸ) በአሰሪው የሚጣ ጥቅሉ ሰ\n18--%/ብር/ (ረ + ሰ)", 15),
            ("", 10),
        ]
        for col, (hdr, _) in enumerate(headers_p4):
            ws.write(8, col, hdr, th_fmt)

        # ── Rows 9+: Employee data ────────────────────────────────────────
        lines = data.get("lines", [])
        totals = data.get("totals", {})
        current_row = 9

        for idx, line in enumerate(lines, 1):
            ws.set_row(current_row, 16)
            start_date_str = str(line["start_date"]) if line.get("start_date") else "-"
            total_pension_18 = round(
                line.get("employee_pension", 0.0) + line.get("employer_pension", 0.0), 2
            )
            ws.write(current_row, 0, line.get("sn", idx), td_center)
            ws.write(current_row, 1, line.get("employee_name", ""), td_left)
            ws.write(current_row, 2, start_date_str, td_center)
            ws.write(current_row, 3, line.get("basic_salary", 0.0), td_num)
            ws.write(current_row, 4, line.get("employee_pension", 0.0), td_num)
            ws.write(current_row, 5, line.get("employer_pension", 0.0), td_num)
            ws.write(current_row, 6, total_pension_18, td_num)
            ws.write(current_row, 7, "", td_center)
            current_row += 1

        # Totals row
        ws.set_row(current_row, 18)
        ws.merge_range(current_row, 0, current_row, 2, "ሕጣሬ ቅጹች  (Page Total)", total_lbl)
        ws.write(current_row, 3, totals.get("basic_salary", 0.0), total_num)
        ws.write(current_row, 4, totals.get("employee_pension", 0.0), total_num)
        ws.write(current_row, 5, totals.get("employer_pension", 0.0), total_num)
        ws.write(current_row, 6, totals.get("total_pension", 0.0), total_num)
        ws.write(current_row, 7, "", total_lbl)
        current_row += 1

        ws.set_row(current_row, 18)
        ws.merge_range(current_row, 0, current_row, 2, "የወጣ ድምር  (Grand Total)", total_lbl)
        ws.write(current_row, 3, totals.get("basic_salary", 0.0), total_num)
        ws.write(current_row, 4, totals.get("employee_pension", 0.0), total_num)
        ws.write(current_row, 5, totals.get("employer_pension", 0.0), total_num)
        ws.write(current_row, 6, totals.get("total_pension", 0.0), total_num)
        ws.write(current_row, 7, "", total_lbl)
        current_row += 2

        # ── ክፍል -3: Summary lines 10-50 ─────────────────────────────────
        ws.merge_range(
            current_row, 0, current_row, 7,
            "ክፍል -3 — የወና የተጠቃሰው ሃሃሀ  (Section 3 — Summary)",
            sec_hdr_fmt,
        )
        current_row += 1

        summary_rows = [
            ("10", "በዚህ ወር ደሞዝዝ የሚስኒፌሎቸው የሠራተኞች ቁጥር  (Number of Employees)",
             data.get("employee_count", 0), "int"),
            ("20", "የወና ብቅሳሳ የሠራተኞች ጀምኦ ደምሮ (ሃሳይ ካሰው-ሩዋ.ሰ.ፈ(ሰ))",
             totals.get("basic_salary", 0.0), "num"),
            ("30", "የወና ብቅሳሳ የሰራተኞች መዋጮ መጠን ሰ (ሃሳይ ካሰው-ሩዋ.ሩዋ.ዝሩ3ሩ( ዘ))",
             totals.get("employee_pension", 0.0), "num"),
            ("40", "የወና ብቅሳሳ የአሰሪው መዋጮ መጠን  (ሃሳይ ካሰው-ሩዋ.ሩዋ.ዝሩ3ሩ( ሰ))",
             totals.get("employer_pension", 0.0), "num"),
            ("50", "የወና ብቅሳሳ የሙሉዓ መዋጮ መጠን  (ሃሳይ ካሰው-ሩዋ.ሩዋ.ሩዋ.ዝሩ3ሩ( ሸ))",
             totals.get("total_pension", 0.0), "num"),
        ]
        for line_no, label, value, vtype in summary_rows:
            ws.set_row(current_row, 18)
            ws.write(current_row, 0, line_no, sec3_lbl)
            ws.merge_range(current_row, 1, current_row, 5, label, sec3_lbl)
            if vtype == "int":
                ws.write(current_row, 6, value, sec3_int)
            else:
                ws.write(current_row, 6, value, sec3_num)
            ws.write(current_row, 7, "", sec3_lbl)
            current_row += 1

        # Reference lines (line 20, 30, 40) note
        current_row += 1
        ws.merge_range(
            current_row, 0, current_row, 7,
            "(line 20)                        (line 30)                       (line 40)",
            italic_fmt,
        )
        current_row += 2

        # ── ክፍል -5: Declaration signature ───────────────────────────────
        ws.merge_range(
            current_row, 0, current_row, 7,
            "ክፍል -5  የትክክለኝነት ማረጋገጫ  (Section 5 — Taxpayer Declaration)",
            sec_hdr_fmt,
        )
        current_row += 1
        ws.merge_range(
            current_row, 0, current_row, 7,
            (
                "ሃሳይ የተጠበው ማስታወቂያ ማስቀሰ ትርምት መረጃ በሙሉ፡ የተሞሉና ትክክለኝነቸን ቅዳ ቼ ቼ\n"
                "አርጋማሳሱ:: ትክክሰኝ ያልሆነ መረጃ ማቅረብ ሀናቢር ሃቸሚም ሃ'ን ሀዐ።ጀዘሰኝ ወቅዉ ሁ::\n"
                "ሃዋ የሚደቀዘ ወሃሞ'ን አንቅዘሃዘሰሰ።  —  Ethiopian Revenue & Customs Authority (as of /5/2011)"
            ),
            italic_fmt,
        )
        current_row += 2

        ws.write(current_row, 0, "የድርጅቱ ሃ4ይ/ሩ.7ዋ ወሃሶ ስም", sig_title_fmt)
        ws.write(current_row, 4, "ስም", sig_title_fmt)
        ws.write(current_row, 5, "ፌርማ", sig_title_fmt)
        current_row += 1
        ws.write(current_row, 0, "Name: ________________________________________________", sig_text_fmt)
        ws.write(current_row, 4, "__________________________________", sig_text_fmt)
        ws.write(current_row, 5, "__________________________________", sig_text_fmt)
        current_row += 1
        ws.write(current_row, 0, "Signature: _____________ Date: _________________________", sig_text_fmt)
        current_row += 1
        ws.write(current_row, 0, "ERCA Form ---- (P1/2011)", italic_fmt)
