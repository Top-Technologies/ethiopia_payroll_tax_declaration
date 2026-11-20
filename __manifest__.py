{
    "name": "Ethiopia Payroll Tax Declaration (Schedule A)",
    "version": "18.0.1.0.0",
    "category": "Human Resources/Payroll",
    "summary": "MoR / ERCA Schedule A Employment Income Tax and Pension Declaration",
    "description": """
Ethiopian Ministry of Revenues (MoR / ERCA) Schedule A Tax Declaration
======================================================================
Adds buttons on the Payslip Batch (hr.payslip.run) form view to generate
and print the official QWeb PDF report and export formatted Excel spreadsheets
matching the Ethiopian Ministry of Revenues (MoR / ERCA) Schedule A
Employment Income Tax Declaration and Pension Contribution Declaration.

Features:
---------
* Inherits hr.payslip.run to calculate all Schedule A values across payslips
  in states 'done', 'paid', or 'verify'.
* Mapped Ethiopian salary rule codes:
    - Basic Salary (BASIC)
    - Non-taxable Allowance (NT_ALLOW)
    - Taxable Allowance (TA_ALLOW)
    - Overtime (OT)
    - Taxable Income (GROSS_TAXABLE or derived)
    - Gross Salary (GROSS or derived)
    - Income Tax Withheld (IIT)
    - Employer Pension 11% (PENS_11)
    - Employee Pension 7% (PENS_7)
    - Other Deductions (OTHER_DED)
    - Total Deductions (TOTAL_DED or derived)
    - Net Pay (NET or derived)
* Official A4 Landscape layout with compact margins.
* Header with Company TIN, name, tax period, region, zone, phone.
* Complete employee breakdown with column totals and Section 3 summary.
* Official certification and sign-off blocks.
* ERCA button on Payslip Batch to export official Excel spreadsheet (.xlsx).
* Interactive wizard with preview summary and one-click download.
    """,
    "author": "TopTec",
    "website": "https://www.toptec.com.et",
    "license": "LGPL-3",
    "depends": [
        "hr_payroll",
        "web",
    ],
    "data": [
        "security/ir.model.access.csv",
        "reports/report_action.xml",
        "reports/report_mor_schedule_a.xml",
        "views/hr_payslip_run_views.xml",
        "wizard/schedule_a_export_wizard_views.xml",
    ],
    "installable": True,
    "application": False,
    "auto_install": False,
}
