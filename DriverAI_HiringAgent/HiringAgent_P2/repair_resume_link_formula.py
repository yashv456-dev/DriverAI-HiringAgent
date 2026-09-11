"""Repair the live workbook's Resume Link calculated-column formula.

The Excel/Graph row API returns a HYPERLINK formula as its display text, so a
manual workbook edit or a full-row rewrite can leave Resume Link as a plain
filename. This script restores the uniform table formula that displays the
friendly Original Filename while opening the stored Resume URL.
"""

import hiring_agent.config  # noqa: F401 - loads .env and logging config
from hiring_agent.sharepoint_scoring import _RESUME_LINK_FORMULA
from sharepoint_client import SharePointClient


def _repair_table(client: SharePointClient, table_name: str | None) -> tuple[int, str]:
    if not table_name:
        return 0, ""
    cols = client._table_columns_of(table_name)
    if "Resume Link" not in cols:
        return 0, "Resume Link column missing"
    client.set_calculated_column(table_name, "Resume Link", _RESUME_LINK_FORMULA)
    rows = client._rows_paged(f"{client._wb_base()}/tables/{table_name}/rows")
    first_formula = ""
    if rows:
        first_idx = int(rows[0].get("index", 0) or 0)
        first_formula = client._column_formula_at(table_name, "Resume Link", first_idx)
    return len(rows), first_formula


def main() -> None:
    client = SharePointClient()
    print("Repairing Resume Link formulas...")

    main_rows, main_formula = _repair_table(client, client.table)
    print(f"  Main table ({client.table}): {main_rows} row(s)")
    print(f"    Formula: {main_formula or '(no data rows)'}")

    rejected_table = client._rejected_table_name_if_exists()
    rejected_rows, rejected_formula = _repair_table(client, rejected_table)
    if rejected_table:
        print(f"  Rejected table ({rejected_table}): {rejected_rows} row(s)")
        print(f"    Formula: {rejected_formula or '(no data rows)'}")
    else:
        print("  Rejected table: not present")

    print("Done.")


if __name__ == "__main__":
    main()
