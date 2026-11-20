from odoo import models, _
from odoo.exceptions import UserError


class HrPayslip(models.Model):
    _inherit = "hr.payslip"

    def action_print_schedule_a_tax_report(self):
        """Print the Schedule A Tax Report from an individual payslip."""
        self.ensure_one()
        if self.payslip_run_id:
            return self.payslip_run_id.action_print_schedule_a_tax_report()
        raise UserError(
            _(
                "This payslip is not part of a Payslip Batch. "
                "Please go to Payroll > Payslips > Batches to print the Schedule A Tax Declaration."
            )
        )

    def action_open_erca_wizard(self):
        """Open the ERCA Excel export wizard from an individual payslip."""
        self.ensure_one()
        if self.payslip_run_id:
            return self.payslip_run_id.action_open_erca_wizard()
        raise UserError(
            _(
                "This payslip is not part of a Payslip Batch. "
                "Please go to Payroll > Payslips > Batches to export the ERCA Tax Declaration."
            )
        )

