import base64
from odoo import api, fields, models, _
from odoo.exceptions import UserError


class ScheduleAExportWizard(models.TransientModel):
    _name = "schedule.a.export.wizard"
    _description = "ERCA Schedule A Tax Declaration - Excel Export Wizard"

    payslip_run_id = fields.Many2one(
        "hr.payslip.run",
        string="Payslip Batch",
        required=True,
        default=lambda self: self.env.context.get("active_id"),
    )
    batch_name = fields.Char(
        related="payslip_run_id.name",
        string="Batch Name",
        readonly=True,
    )
    company_id = fields.Many2one(
        related="payslip_run_id.company_id",
        string="Company",
        readonly=True,
    )
    currency_id = fields.Many2one(
        related="company_id.currency_id",
        string="Currency",
        readonly=True,
    )
    date_start = fields.Date(
        related="payslip_run_id.date_start",
        string="Tax Period Start",
        readonly=True,
    )
    date_end = fields.Date(
        related="payslip_run_id.date_end",
        string="Tax Period End",
        readonly=True,
    )

    employee_count = fields.Integer(
        string="Total Employees",
        compute="_compute_summary",
    )
    total_taxable_income = fields.Float(
        string="Total Taxable Income",
        compute="_compute_summary",
    )
    total_tax_withheld = fields.Float(
        string="Total Tax Withheld",
        compute="_compute_summary",
    )
    total_pension = fields.Float(
        string="Total Pension (7% + 11%)",
        compute="_compute_summary",
    )

    state = fields.Selection(
        [("choose", "Choose"), ("get", "Get")],
        default="choose",
    )
    excel_file = fields.Binary(
        string="Excel File",
        readonly=True,
    )
    excel_filename = fields.Char(
        string="Filename",
        readonly=True,
    )

    @api.depends("payslip_run_id")
    def _compute_summary(self):
        for wizard in self:
            if wizard.payslip_run_id:
                data = wizard.payslip_run_id._get_schedule_a_data()
                wizard.employee_count = data.get("employee_count", 0)
                wizard.total_taxable_income = data.get("totals", {}).get(
                    "taxable_income", 0.0
                )
                wizard.total_tax_withheld = data.get("totals", {}).get(
                    "income_tax", 0.0
                )
                wizard.total_pension = data.get("totals", {}).get(
                    "total_pension", 0.0
                )
            else:
                wizard.employee_count = 0
                wizard.total_taxable_income = 0.0
                wizard.total_tax_withheld = 0.0
                wizard.total_pension = 0.0

    def action_export_excel(self):
        """Generate and download the ERCA Schedule A Excel spreadsheet."""
        self.ensure_one()
        if not self.payslip_run_id:
            raise UserError(_("No Payslip Batch selected for export."))

        xlsx_bytes = self.payslip_run_id._generate_schedule_a_excel()
        clean_name = (
            (self.payslip_run_id.name or "Batch")
            .replace("/", "_")
            .replace("\\", "_")
        )
        filename = f"ERCA_Schedule_A_{clean_name}.xlsx"
        b64_data = base64.b64encode(xlsx_bytes)

        self.write({
            "excel_file": b64_data,
            "excel_filename": filename,
            "state": "get",
        })

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/?model=schedule.a.export.wizard&id={self.id}&field=excel_file&download=true&filename={filename}",
            "target": "self",
        }

    def action_print_pdf(self):
        """Print the official QWeb PDF declaration directly from the wizard."""
        self.ensure_one()
        return self.payslip_run_id.action_print_schedule_a_tax_report()

