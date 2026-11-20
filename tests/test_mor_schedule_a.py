from datetime import date
from odoo.tests import tagged
from odoo.tests.common import TransactionCase


@tagged("post_install", "-at_install")
class TestEthiopiaPayrollTaxDeclaration(TransactionCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.company = cls.env.company
        cls.company.write({
            "vat": "0012345678",
            "phone": "+251 11 123 4567",
        })
        cls.company.partner_id.write({
            "city": "Addis Ababa",
            "street2": "Bole Subcity, Woreda 03",
        })

        cls.employee = cls.env["hr.employee"].create({
            "name": "Abebe Kebede",
            "identification_id": "EMP-ETH-001",
        })

        cls.batch = cls.env["hr.payslip.run"].create({
            "name": "Payroll Batch - January 2026",
            "date_start": date(2026, 1, 1),
            "date_end": date(2026, 1, 31),
        })

    def test_01_empty_batch_schedule_a_data(self):
        """Empty batch should return zero totals and empty lines without failing."""
        data = self.batch._get_schedule_a_data()
        self.assertEqual(data["employee_count"], 0)
        self.assertEqual(len(data["lines"]), 0)
        self.assertEqual(data["totals"]["basic_salary"], 0.0)
        self.assertEqual(data["totals"]["taxable_income"], 0.0)
        self.assertEqual(data["totals"]["income_tax"], 0.0)
        self.assertEqual(data["totals"]["total_pension"], 0.0)
        self.assertEqual(data["totals"]["net_pay"], 0.0)

    def test_02_report_action_returned(self):
        """action_print_schedule_a_tax_report returns the configured ir.actions.report."""
        action = self.batch.action_print_schedule_a_tax_report()
        self.assertEqual(action.get("type"), "ir.actions.report")
        self.assertEqual(
            action.get("report_name"),
            "ethiopia_payroll_tax_declaration.report_mor_schedule_a_template",
        )

    def test_03_open_erca_wizard_action(self):
        """action_open_erca_wizard opens the schedule.a.export.wizard in a new target window."""
        action = self.batch.action_open_erca_wizard()
        self.assertEqual(action.get("type"), "ir.actions.act_window")
        self.assertEqual(action.get("res_model"), "schedule.a.export.wizard")
        self.assertEqual(action.get("target"), "new")
        self.assertEqual(
            action.get("context", {}).get("default_payslip_run_id"),
            self.batch.id,
        )

    def test_04_wizard_summary_computation(self):
        """Wizard correctly computes summary totals from the payslip batch."""
        wizard = self.env["schedule.a.export.wizard"].create({
            "payslip_run_id": self.batch.id,
        })
        self.assertEqual(wizard.employee_count, 0)
        self.assertEqual(wizard.total_taxable_income, 0.0)
        self.assertEqual(wizard.total_tax_withheld, 0.0)
        self.assertEqual(wizard.total_pension, 0.0)
        self.assertEqual(wizard.state, "choose")

    def test_05_payslip_delegation_to_batch(self):
        """Payslip action_open_erca_wizard delegates cleanly to its parent batch."""
        slip = self.env["hr.payslip"].create({
            "name": "January 2026 - Abebe Kebede",
            "employee_id": self.employee.id,
            "payslip_run_id": self.batch.id,
        })
        action = slip.action_open_erca_wizard()
        self.assertEqual(action.get("res_model"), "schedule.a.export.wizard")
        self.assertEqual(
            action.get("context", {}).get("default_payslip_run_id"),
            self.batch.id,
        )
