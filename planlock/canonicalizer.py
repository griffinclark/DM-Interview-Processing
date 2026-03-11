from __future__ import annotations

from collections import OrderedDict

from planlock.models import (
    CanonicalPlanDocument,
    ExpenseCandidate,
    FieldCandidate,
    HoldingCandidate,
    ImportWarning,
    PageMappingResult,
    Severity,
    Stage,
)
from planlock.template_schema import EXPENSE_ROW_BLOCKS, FIELD_TARGETS


def _value_signature(value: object) -> str:
    return str(value).strip().lower()


def merge_page_mappings(page_results: list[PageMappingResult]) -> tuple[CanonicalPlanDocument, list[ImportWarning]]:
    warnings: list[ImportWarning] = []
    fields: OrderedDict[str, FieldCandidate] = OrderedDict()
    expenses: OrderedDict[str, ExpenseCandidate] = OrderedDict()
    accounts = []
    holdings: list[HoldingCandidate] = []
    unmapped_items: list[str] = []
    assumptions: list[str] = []

    for result in page_results:
        for field in result.mapped_fields:
            if field.target_key not in FIELD_TARGETS:
                unmapped_items.append(
                    f"Page {result.page_number}: unsupported target key {field.target_key}"
                )
                warnings.append(
                    ImportWarning(
                        code="unsupported_target_key",
                        message=f"Unsupported target key returned by mapping stage: {field.target_key}",
                        severity=Severity.WARNING,
                        stage=Stage.DATA_ENTRY,
                        page_numbers=[result.page_number],
                    )
                )
                continue

            existing = fields.get(field.target_key)
            if not existing:
                fields[field.target_key] = field
                if field.comment:
                    assumptions.append(field.comment)
                continue

            if _value_signature(existing.value) == _value_signature(field.value):
                if field.confidence > existing.confidence:
                    fields[field.target_key] = field
                continue

            chosen = field if field.confidence >= existing.confidence else existing
            discarded = existing if chosen is field else field
            fields[field.target_key] = chosen
            warnings.append(
                ImportWarning(
                    code="conflicting_field_values",
                    message=(
                        f"Conflicting values for {field.target_key}; "
                        f"kept {chosen.value!r} over {discarded.value!r}."
                    ),
                    severity=Severity.WARNING,
                    stage=Stage.DATA_ENTRY,
                    page_numbers=sorted(
                        {page for page in [existing.page_number, field.page_number] if page}
                    ),
                )
            )

        for expense in result.expenses:
            if expense.category not in EXPENSE_ROW_BLOCKS:
                unmapped_items.append(
                    f"Page {result.page_number}: unsupported expense category {expense.category}"
                )
                continue

            existing = expenses.get(expense.category)
            if not existing:
                expenses[expense.category] = expense
                if expense.comment:
                    assumptions.append(expense.comment)
                continue

            same_monthly = existing.monthly_amount == expense.monthly_amount
            same_yearly = existing.yearly_amount == expense.yearly_amount
            if same_monthly and same_yearly:
                if expense.confidence > existing.confidence:
                    expenses[expense.category] = expense
                continue

            chosen = expense if expense.confidence >= existing.confidence else existing
            expenses[expense.category] = chosen
            warnings.append(
                ImportWarning(
                    code="conflicting_expense_values",
                    message=f"Conflicting values for expense category {expense.category}; highest confidence kept.",
                    severity=Severity.WARNING,
                    stage=Stage.DATA_ENTRY,
                    page_numbers=sorted(
                        {page for page in [existing.page_number, expense.page_number] if page}
                    ),
                )
            )

        for account in result.accounts:
            dedupe_key = (
                account.account_type,
                account.owner_name,
                account.account_identifier,
                account.institution,
                account.balance,
            )
            if dedupe_key not in {
                (
                    existing.account_type,
                    existing.owner_name,
                    existing.account_identifier,
                    existing.institution,
                    existing.balance,
                )
                for existing in accounts
            }:
                accounts.append(account)

        for holding in result.holdings:
            dedupe_key = (
                holding.sheet_name,
                holding.owner_section,
                holding.account_name,
                holding.holding_name,
                holding.symbol,
            )
            if dedupe_key not in {
                (
                    existing.sheet_name,
                    existing.owner_section,
                    existing.account_name,
                    existing.holding_name,
                    existing.symbol,
                )
                for existing in holdings
            }:
                holdings.append(holding)

        unmapped_items.extend(f"Page {result.page_number}: {item}" for item in result.unmapped_items)
        for warning in result.warnings:
            warnings.append(
                ImportWarning(
                    code="mapping_warning",
                    message=warning,
                    severity=Severity.WARNING,
                    stage=Stage.DATA_ENTRY,
                    page_numbers=[result.page_number],
                )
            )

    return (
        CanonicalPlanDocument(
            fields=dict(fields),
            expenses=list(expenses.values()),
            accounts=accounts,
            holdings=holdings,
            unmapped_items=sorted(dict.fromkeys(unmapped_items)),
            assumptions=sorted(dict.fromkeys(filter(None, assumptions))),
        ),
        warnings,
    )
