from __future__ import annotations

from dataclasses import dataclass
from string import ascii_uppercase


def cell_range(column_start: str, column_end: str, start_row: int, end_row: int) -> list[str]:
    columns = []
    started = False
    for column in ascii_uppercase:
        if column == column_start:
            started = True
        if started:
            columns.append(column)
        if column == column_end:
            break
    return [f"{column}{row}" for row in range(start_row, end_row + 1) for column in columns]


FIELD_TARGETS: dict[str, tuple[str, str]] = {
    "profile.client_1.first_name": ("Data Input", "C6"),
    "profile.client_1.last_name": ("Data Input", "D6"),
    "profile.client_1.birthdate": ("Data Input", "E6"),
    "profile.client_1.retirement_age": ("Data Input", "G6"),
    "profile.client_1.tax_filing_status": ("Data Input", "H6"),
    "profile.client_1.state": ("Data Input", "I6"),
    "profile.client_1.employer": ("Data Input", "J6"),
    "profile.client_1.job_title": ("Data Input", "K6"),
    "profile.client_2.first_name": ("Data Input", "C7"),
    "profile.client_2.last_name": ("Data Input", "D7"),
    "profile.client_2.birthdate": ("Data Input", "E7"),
    "profile.client_2.retirement_age": ("Data Input", "G7"),
    "profile.client_2.tax_filing_status": ("Data Input", "H7"),
    "profile.client_2.state": ("Data Input", "I7"),
    "profile.client_2.employer": ("Data Input", "J7"),
    "profile.client_2.job_title": ("Data Input", "K7"),
    "profile.child_1.first_name": ("Data Input", "C10"),
    "profile.child_1.last_name": ("Data Input", "D10"),
    "profile.child_1.birthdate": ("Data Input", "E10"),
    "profile.child_2.first_name": ("Data Input", "C11"),
    "profile.child_2.last_name": ("Data Input", "D11"),
    "profile.child_2.birthdate": ("Data Input", "E11"),
    "profile.child_3.first_name": ("Data Input", "C12"),
    "profile.child_3.last_name": ("Data Input", "D12"),
    "profile.child_3.birthdate": ("Data Input", "E12"),
    "profile.child_4.first_name": ("Data Input", "C13"),
    "profile.child_4.last_name": ("Data Input", "D13"),
    "profile.child_4.birthdate": ("Data Input", "E13"),
    "income.client_1.pay_periods_per_year": ("Data Input", "D43"),
    "income.client_2.pay_periods_per_year": ("Data Input", "E43"),
    "income.client_1.base_salary_gross": ("Data Input", "D46"),
    "income.client_2.base_salary_gross": ("Data Input", "E46"),
    "income.client_1.bonus": ("Data Input", "D47"),
    "income.client_2.bonus": ("Data Input", "E47"),
    "income.client_1.commission": ("Data Input", "D48"),
    "income.client_2.commission": ("Data Input", "E48"),
    "income.client_1.rsus": ("Data Input", "D49"),
    "income.client_2.rsus": ("Data Input", "E49"),
    "income.client_1.severance": ("Data Input", "D50"),
    "income.client_2.severance": ("Data Input", "E50"),
    "income.client_1.self_employment": ("Data Input", "F53"),
    "income.client_2.self_employment": ("Data Input", "G53"),
    "income.client_1.scorp_distributions": ("Data Input", "F54"),
    "income.client_2.scorp_distributions": ("Data Input", "G54"),
    "income.client_1.investment_income_1099": ("Data Input", "F55"),
    "income.client_2.investment_income_1099": ("Data Input", "G55"),
    "income.client_1.rental_income": ("Data Input", "F56"),
    "income.client_2.rental_income": ("Data Input", "G56"),
    "deductions.client_1.pre_tax_retirement": ("Data Input", "D60"),
    "deductions.client_2.pre_tax_retirement": ("Data Input", "E60"),
    "deductions.client_1.medical": ("Data Input", "D61"),
    "deductions.client_2.medical": ("Data Input", "E61"),
    "deductions.client_1.dental": ("Data Input", "D62"),
    "deductions.client_2.dental": ("Data Input", "E62"),
    "deductions.client_1.vision": ("Data Input", "D63"),
    "deductions.client_2.vision": ("Data Input", "E63"),
    "deductions.client_1.dependent_care": ("Data Input", "D64"),
    "deductions.client_2.dependent_care": ("Data Input", "E64"),
    "deductions.client_1.fsa": ("Data Input", "D65"),
    "deductions.client_2.fsa": ("Data Input", "E65"),
    "deductions.client_1.hsa": ("Data Input", "D66"),
    "deductions.client_2.hsa": ("Data Input", "E66"),
    "deductions.client_1.roth": ("Data Input", "D70"),
    "deductions.client_2.roth": ("Data Input", "E70"),
    "deductions.client_1.mega_backdoor_roth": ("Data Input", "D71"),
    "deductions.client_2.mega_backdoor_roth": ("Data Input", "E71"),
    "deductions.client_1.espp": ("Data Input", "D72"),
    "deductions.client_2.espp": ("Data Input", "E72"),
    "deductions.client_1.life_insurance": ("Data Input", "D73"),
    "deductions.client_2.life_insurance": ("Data Input", "E73"),
    "deductions.client_1.disability_insurance": ("Data Input", "D74"),
    "deductions.client_2.disability_insurance": ("Data Input", "E74"),
    "deductions.client_1.post_tax_other": ("Data Input", "D75"),
    "deductions.client_2.post_tax_other": ("Data Input", "E75"),
    "savings.client_1.savings_account_annual": ("Data Input", "F80"),
    "savings.client_2.savings_account_annual": ("Data Input", "G80"),
    "savings.client_1.ira_annual": ("Data Input", "F81"),
    "savings.client_2.ira_annual": ("Data Input", "G81"),
    "savings.client_1.brokerage_annual": ("Data Input", "F82"),
    "savings.client_2.brokerage_annual": ("Data Input", "G82"),
    "savings.client_1.529_annual": ("Data Input", "F83"),
    "savings.client_2.529_annual": ("Data Input", "G83"),
    "taxes.client_1.federal_income_tax": ("Data Input", "D89"),
    "taxes.client_2.federal_income_tax": ("Data Input", "E89"),
    "taxes.client_1.social_security": ("Data Input", "D90"),
    "taxes.client_2.social_security": ("Data Input", "E90"),
    "taxes.client_1.medicare": ("Data Input", "D91"),
    "taxes.client_2.medicare": ("Data Input", "E91"),
    "taxes.client_1.state_income_tax": ("Data Input", "D95"),
    "taxes.client_2.state_income_tax": ("Data Input", "E95"),
    "taxes.client_1.state_sdi": ("Data Input", "D96"),
    "taxes.client_2.state_sdi": ("Data Input", "E96"),
    "taxes.client_1.state_pfl": ("Data Input", "D97"),
    "taxes.client_2.state_pfl": ("Data Input", "E97"),
    "taxes.client_1.state_ltc": ("Data Input", "D98"),
    "taxes.client_2.state_ltc": ("Data Input", "E98"),
    "taxes.client_1.city_income_tax": ("Data Input", "D102"),
    "taxes.client_2.city_income_tax": ("Data Input", "E102"),
}


EXPENSE_ROW_BLOCKS: dict[str, tuple[int, int]] = {
    "Home Expenses": (6, 12),
    "Auto / Commute": (14, 20),
    "Utilities": (22, 27),
    "Food": (29, 31),
    "Child Related": (33, 35),
    "Health / Personal Care": (37, 42),
    "Shopping": (44, 49),
    "Entertainment": (51, 53),
    "Travel": (55, 58),
    "Miscellaneous": (60, 66),
    "Cash": (68, 70),
}


DATA_INPUT_ACCOUNT_ROWS = list(range(18, 40))
NET_WORTH_ASSET_ROWS = list(range(6, 19))
NET_WORTH_LIABILITY_ROWS = list(range(22, 30))


@dataclass(frozen=True)
class PortfolioBlock:
    sheet_name: str
    owner_section: str
    account_row: int
    holding_start_row: int
    holding_end_row: int


RETIREMENT_BLOCKS = [
    PortfolioBlock("Retirement Accounts", "client_1", 6, 7, 14),
    PortfolioBlock("Retirement Accounts", "client_1", 18, 19, 26),
    PortfolioBlock("Retirement Accounts", "client_1", 30, 31, 38),
    PortfolioBlock("Retirement Accounts", "client_1", 42, 43, 50),
    PortfolioBlock("Retirement Accounts", "client_2", 57, 58, 65),
    PortfolioBlock("Retirement Accounts", "client_2", 69, 70, 77),
    PortfolioBlock("Retirement Accounts", "client_2", 81, 82, 89),
]

TAXABLE_BLOCKS = [
    PortfolioBlock("Taxable Accounts", "client_1", 6, 7, 14),
    PortfolioBlock("Taxable Accounts", "client_1", 18, 19, 26),
    PortfolioBlock("Taxable Accounts", "client_1", 30, 31, 38),
    PortfolioBlock("Taxable Accounts", "client_1", 42, 43, 50),
    PortfolioBlock("Taxable Accounts", "client_2", 57, 58, 65),
    PortfolioBlock("Taxable Accounts", "client_2", 69, 70, 77),
    PortfolioBlock("Taxable Accounts", "client_2", 81, 82, 89),
]

EDUCATION_BLOCKS = [
    PortfolioBlock("Education Accounts", "education", 6, 7, 14),
    PortfolioBlock("Education Accounts", "education", 18, 19, 26),
    PortfolioBlock("Education Accounts", "education", 30, 31, 38),
    PortfolioBlock("Education Accounts", "education", 42, 43, 50),
]


PORTFOLIO_BLOCKS = RETIREMENT_BLOCKS + TAXABLE_BLOCKS + EDUCATION_BLOCKS

TEMPLATE_SHEET_ORDER = [
    "Data Input",
    "Net Worth",
    "Transactions Raw",
    "Expenses",
    "Retirement Accounts",
    "Taxable Accounts",
    "Education Accounts",
]


ALLOWED_WRITE_CELLS_BY_SHEET: dict[str, set[str]] = {}

for sheet_name, cell in FIELD_TARGETS.values():
    ALLOWED_WRITE_CELLS_BY_SHEET.setdefault(sheet_name, set()).add(cell)

ALLOWED_WRITE_CELLS_BY_SHEET.setdefault("Data Input", set()).update(
    {
        f"{column}{row}"
        for row in DATA_INPUT_ACCOUNT_ROWS
        for column in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]
    }
)
ALLOWED_WRITE_CELLS_BY_SHEET.setdefault("Net Worth", set()).update(
    {
        f"{column}{row}"
        for row in [*NET_WORTH_ASSET_ROWS, *NET_WORTH_LIABILITY_ROWS]
        for column in ["B", "C"]
    }
)

for start_row, end_row in EXPENSE_ROW_BLOCKS.values():
    ALLOWED_WRITE_CELLS_BY_SHEET.setdefault("Expenses", set()).update(
        {
            f"{column}{row}"
            for row in range(start_row, end_row + 1)
            for column in ["B", "C", "D", "F", "G"]
        }
    )

for block in RETIREMENT_BLOCKS:
    ALLOWED_WRITE_CELLS_BY_SHEET.setdefault(block.sheet_name, set()).add(f"B{block.account_row}")
    ALLOWED_WRITE_CELLS_BY_SHEET[block.sheet_name].update(
        {
            f"{column}{row}"
            for row in range(block.holding_start_row, block.holding_end_row + 1)
            for column in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
        }
    )

for block in TAXABLE_BLOCKS + EDUCATION_BLOCKS:
    ALLOWED_WRITE_CELLS_BY_SHEET.setdefault(block.sheet_name, set()).add(f"B{block.account_row}")
    ALLOWED_WRITE_CELLS_BY_SHEET[block.sheet_name].update(
        {
            f"{column}{row}"
            for row in range(block.holding_start_row, block.holding_end_row + 1)
            for column in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
        }
    )


ALLOWED_FORMULA_CELLS_BY_SHEET: dict[str, set[str]] = {
    "Expenses": set(),
    "Retirement Accounts": set(),
    "Taxable Accounts": set(),
    "Education Accounts": set(),
}

for start_row, end_row in EXPENSE_ROW_BLOCKS.values():
    for row in range(start_row, end_row + 1):
        ALLOWED_FORMULA_CELLS_BY_SHEET["Expenses"].update({f"C{row}", f"D{row}"})

for block in RETIREMENT_BLOCKS:
    for row in range(block.holding_start_row, block.holding_end_row + 1):
        ALLOWED_FORMULA_CELLS_BY_SHEET["Retirement Accounts"].add(f"K{row}")

for block in TAXABLE_BLOCKS + EDUCATION_BLOCKS:
    for row in range(block.holding_start_row, block.holding_end_row + 1):
        ALLOWED_FORMULA_CELLS_BY_SHEET[block.sheet_name].update({f"K{row}", f"M{row}"})


def is_allowed_write(sheet_name: str, cell: str) -> bool:
    return cell in ALLOWED_WRITE_CELLS_BY_SHEET.get(sheet_name, set())


def is_allowed_formula(sheet_name: str, cell: str) -> bool:
    return cell in ALLOWED_FORMULA_CELLS_BY_SHEET.get(sheet_name, set())


def schema_reference_for_prompt() -> str:
    field_lines = [
        f"- {key} -> {sheet_name}!{cell}"
        for key, (sheet_name, cell) in sorted(FIELD_TARGETS.items())
    ]
    expense_lines = [
        f"- {category}: write only within rows {start_row}-{end_row} on Expenses"
        for category, (start_row, end_row) in EXPENSE_ROW_BLOCKS.items()
    ]
    return "\n".join(
        [
            "Locked workbook target keys:",
            *field_lines,
            "",
            "Expense categories:",
            *expense_lines,
            "",
            "Only map direct evidence. Leave unsupported fields out.",
            "If a page shows current vs suggested values and the workbook has one planning input, prefer the suggested/target figure and note that assumption.",
        ]
    )


def sheet_reference_for_prompt(sheet_name: str) -> str:
    lines = [f"Sheet: {sheet_name}"]
    allowed_cells = sorted(ALLOWED_WRITE_CELLS_BY_SHEET.get(sheet_name, set()))
    if allowed_cells:
        lines.append("Allowed writable cells:")
        lines.append(", ".join(allowed_cells))
    allowed_formulas = sorted(ALLOWED_FORMULA_CELLS_BY_SHEET.get(sheet_name, set()))
    if allowed_formulas:
        lines.append("Cells that must stay formula-driven when derived values are needed:")
        lines.append(", ".join(allowed_formulas))
    if sheet_name == "Expenses":
        lines.append("Expense row blocks:")
        lines.extend(
            f"- {category}: rows {start_row}-{end_row}"
            for category, (start_row, end_row) in EXPENSE_ROW_BLOCKS.items()
        )
    if sheet_name == "Net Worth":
        lines.append("Net worth blocks:")
        lines.append(
            f"- Assets: write account labels to B{NET_WORTH_ASSET_ROWS[0]}:B{NET_WORTH_ASSET_ROWS[-1]} "
            f"and balances to C{NET_WORTH_ASSET_ROWS[0]}:C{NET_WORTH_ASSET_ROWS[-1]}"
        )
        lines.append(
            f"- Liabilities: write account labels to B{NET_WORTH_LIABILITY_ROWS[0]}:B{NET_WORTH_LIABILITY_ROWS[-1]} "
            f"and balances to C{NET_WORTH_LIABILITY_ROWS[0]}:C{NET_WORTH_LIABILITY_ROWS[-1]}"
        )
        lines.append("- Use `accounts` candidates for this sheet and set `net_worth_section` when it is defensible.")
    if sheet_name == "Data Input":
        field_lines = [
            f"- {key} -> {cell}"
            for key, (field_sheet_name, cell) in sorted(FIELD_TARGETS.items())
            if field_sheet_name == sheet_name
        ]
        if field_lines:
            lines.append("Direct field targets:")
            lines.extend(field_lines)
    if sheet_name in {"Retirement Accounts", "Taxable Accounts", "Education Accounts"}:
        blocks = [block for block in PORTFOLIO_BLOCKS if block.sheet_name == sheet_name]
        lines.append("Portfolio blocks:")
        lines.extend(
            f"- {block.owner_section}: account row {block.account_row}, holdings {block.holding_start_row}-{block.holding_end_row}"
            for block in blocks
        )
    return "\n".join(lines)
