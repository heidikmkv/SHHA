from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import re
from typing import Optional

import pandas as pd

ADDRESS_FILE = "addresses_export.csv"
USER_EXPORT_GLOB = "export-users-*.csv"

KNOWN_PRIMARY_STATUSES = {
    "owner occupied",
    "tenant occupied",
    "vacant",
}


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_email(value: object) -> str:
    return clean_text(value).lower()


def is_valid_email(email: str) -> bool:
    if not email or ("@" not in email):
        return False

    normalized = email.lower().strip()
    if "fake.fake" in normalized:
        return False

    domain = normalized.split("@", 1)[1].strip()
    if not domain:
        return False

    if domain == "fake" or domain.startswith("fake.") or domain.endswith(".fake"):
        return False

    return True


def split_user_emails(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []

    for delimiter in [",", ";"]:
        text = text.replace(delimiter, "|")

    emails = [normalize_email(part) for part in text.split("|") if clean_text(part)]
    return [email for email in emails if is_valid_email(email)]


def fmt(value: int) -> str:
    return f"{value:,}"


def pct(part: int, whole: int) -> str:
    if whole == 0:
        return "0.0%"
    return f"{(part / whole) * 100:.1f}%"


def signed_int(value: int) -> str:
    return f"{value:+,}"


def has_required_exports(folder: Path) -> bool:
    return (folder / ADDRESS_FILE).exists() and any(folder.glob(USER_EXPORT_GLOB))


def resolve_data_dir(data_dir: Path) -> Path:
    dated_candidates = [child for child in data_dir.iterdir() if child.is_dir() and has_required_exports(child)]
    if dated_candidates:
        return sorted(dated_candidates, key=lambda path: path.name)[-1]

    if has_required_exports(data_dir):
        return data_dir

    raise FileNotFoundError(
        "No export set found. Expected addresses_export.csv and one export-users-*.csv "
        f"in {data_dir} or dated subfolders under {data_dir}."
    )


def infer_snapshot_date(snapshot_date: Optional[str], resolved_data_dir: Path) -> str:
    if snapshot_date:
        return snapshot_date

    folder_name = resolved_data_dir.name
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", folder_name):
        return folder_name
    if re.fullmatch(r"\d{8}", folder_name):
        return f"{folder_name[0:4]}-{folder_name[4:6]}-{folder_name[6:8]}"

    return date.today().isoformat()


def add_analysis_flags(addresses: pd.DataFrame) -> pd.DataFrame:
    df = addresses.copy()

    df["status_clean"] = df["Status"].map(clean_text)
    df["status_norm"] = df["status_clean"].str.lower()

    df["is_vacant"] = df["status_norm"] == "vacant"
    df["is_blank_status"] = df["status_norm"] == ""
    df["is_other_status"] = ~df["status_norm"].isin(KNOWN_PRIMARY_STATUSES)
    df["is_other_or_vacant"] = df["is_vacant"] | df["is_other_status"]

    df["users_text"] = df["Users"].map(clean_text)
    df["emails_text"] = df["User Emails"].map(clean_text)
    df["tenants_text"] = df["Tenants"].map(clean_text)

    df["has_users"] = df["users_text"] != ""
    df["has_email_field"] = df["emails_text"] != ""
    df["valid_emails"] = df["User Emails"].map(split_user_emails)
    df["valid_email_count"] = df["valid_emails"].map(len)
    df["has_valid_email"] = df["valid_email_count"] > 0

    df["is_member"] = pd.to_numeric(df["Is Member"], errors="coerce").fillna(0).astype(int) == 1
    df["mail_grit"] = pd.to_numeric(df["Mail GRIT"], errors="coerce").fillna(0).astype(int) == 1
    df["has_tenant_text"] = df["tenants_text"] != ""

    df["issue_has_users"] = df["is_other_or_vacant"] & df["has_users"]
    df["issue_has_valid_email"] = df["is_other_or_vacant"] & df["has_valid_email"]
    df["issue_marked_member"] = df["is_other_or_vacant"] & df["is_member"]
    df["issue_receives_grit"] = df["is_other_or_vacant"] & df["mail_grit"]
    df["issue_has_tenant_data"] = df["is_other_or_vacant"] & df["has_tenant_text"]

    return df


def trend_lines(output_dir: Path, snapshot_date: str, total_vacant_other: int, open_issues: int) -> list[str]:
    trend_path = output_dir / "vacant_other_trend_history.csv"
    current = {
        "run_date": snapshot_date,
        "vacant_or_other_lots": total_vacant_other,
        "open_edge_case_flags": open_issues,
    }

    if trend_path.exists():
        trend_df = pd.read_csv(trend_path)
    else:
        trend_df = pd.DataFrame(columns=["run_date", "vacant_or_other_lots", "open_edge_case_flags"])

    for col in ["run_date", "vacant_or_other_lots", "open_edge_case_flags"]:
        if col not in trend_df.columns:
            trend_df[col] = pd.NA

    previous_snapshot = None
    prior_rows = trend_df[trend_df["run_date"].astype(str) < snapshot_date].copy()
    if not prior_rows.empty:
        prior_rows = prior_rows.sort_values("run_date")
        previous_snapshot = prior_rows.iloc[-1].to_dict()

    trend_df = trend_df[trend_df["run_date"].astype(str) != snapshot_date]
    trend_df = pd.concat([trend_df, pd.DataFrame([current])], ignore_index=True)
    trend_df = trend_df.sort_values("run_date")

    for col in ["vacant_or_other_lots", "open_edge_case_flags"]:
        trend_df[col] = pd.to_numeric(trend_df[col], errors="coerce").fillna(0).astype(int)

    trend_df.to_csv(trend_path, index=False)

    if previous_snapshot is None:
        return ["- No prior snapshot yet for trend comparison."]

    prev_vacant_other = int(float(previous_snapshot.get("vacant_or_other_lots", 0)))
    prev_open_issues = int(float(previous_snapshot.get("open_edge_case_flags", 0)))
    prior_date = str(previous_snapshot.get("run_date", "unknown"))

    return [
        f"Prior snapshot date: {prior_date}",
        f"- Vacant/other lots: {fmt(total_vacant_other)} ({signed_int(total_vacant_other - prev_vacant_other)})",
        f"- Open edge-case flags: {fmt(open_issues)} ({signed_int(open_issues - prev_open_issues)})",
    ]


def write_report(
    output_path: Path,
    snapshot_date: str,
    resolved_data_dir: Path,
    addresses_path: Path,
    analyzed: pd.DataFrame,
    status_counts: pd.Series,
    issue_counts: dict[str, int],
    trend_output: list[str],
) -> None:
    total_households = len(analyzed)
    target = analyzed[analyzed["is_other_or_vacant"]].copy()
    vacant = analyzed[analyzed["is_vacant"]].copy()
    other = analyzed[analyzed["is_other_status"] & ~analyzed["is_vacant"]].copy()

    lines: list[str] = []
    lines.append("# Vacant/Other Lot Edge-Case Report")
    lines.append("")
    lines.append(f"Data pull date: {snapshot_date}")
    lines.append(f"Source folder: {resolved_data_dir}")
    lines.append(f"Source file: {addresses_path}")
    lines.append("")

    lines.append("## Purpose")
    lines.append(
        "Provide a hand-review list of vacant lots and key contact/admin fields for office cleanup and follow-up."
    )
    lines.append("")

    lines.append("## Summary")
    lines.append(f"- Total households reviewed: {fmt(total_households)}")
    lines.append(
        f"- Vacant or other-status lots: {fmt(len(target))} ({pct(len(target), total_households)})"
    )
    lines.append(f"- Vacant lots: {fmt(len(vacant))}")
    lines.append(f"- Other-status lots: {fmt(len(other))}")
    lines.append("")

    lines.append("## Status breakdown (vacant + other)\n")
    if not status_counts.empty:
        for status, count in status_counts.items():
            label = status if status else "[blank]"
            lines.append(f"- {label}: {fmt(int(count))}")
    else:
        lines.append("- None")
    lines.append("")

    lines.append("## Admin review counts")
    lines.append(f"- Vacant/other lots with listed users (good contact coverage): {fmt(issue_counts['has_users'])}")
    lines.append(f"- Vacant/other lots with valid user email: {fmt(issue_counts['has_valid_email'])}")
    lines.append(f"- Vacant/other lots marked as member: {fmt(issue_counts['marked_member'])}")
    lines.append(f"- Vacant/other lots set to receive printed GRIT: {fmt(issue_counts['receives_grit'])}")
    lines.append(f"- Vacant/other lots with tenant data entered: {fmt(issue_counts['has_tenant_data'])}")
    lines.append("")

    lines.append("## Change since previous data pull")
    lines.extend(trend_output)
    lines.append("")

    lines.append("## Vacant lots table (for office admin)")
    lines.append("| Status | Address | Users | Membership status | Separate mailing address |")
    lines.append("|---|---|---|---|---|")

    vacant_for_table = vacant.sort_values(["Status", "Street", "Number", "Unit"], na_position="last")
    if vacant_for_table.empty:
        lines.append("| [none] |  |  |  |  |")
    else:
        for _, row in vacant_for_table.iterrows():
            status = clean_text(row.get("Status", "")) or "[blank]"
            address = clean_text(row.get("Address", "")) or "[blank]"
            users = clean_text(row.get("Users", "")) or "[none]"

            member_value = pd.to_numeric(clean_text(row.get("Is Member", "")), errors="coerce")
            membership_status = "Member" if pd.notna(member_value) and int(member_value) == 1 else "Non-member"

            mail_address = clean_text(row.get("Mail Address", ""))
            mail_city = clean_text(row.get("Mail City", ""))
            mail_state = clean_text(row.get("Mail State", ""))
            mail_zip = clean_text(row.get("Mail Postal Code", ""))

            mailing_full = ", ".join(part for part in [mail_address, mail_city, mail_state, mail_zip] if part)
            property_address = clean_text(row.get("Address", ""))
            separate_mailing = mailing_full if mailing_full and mailing_full.lower() != property_address.lower() else "No"

            status_md = status.replace("|", "\\|")
            address_md = address.replace("|", "\\|")
            users_md = users.replace("|", ", ")
            separate_md = separate_mailing.replace("|", "\\|")
            lines.append(f"| {status_md} | {address_md} | {users_md} | {membership_status} | {separate_md} |")
    lines.append("")

    lines.append("## Vacant/other lots set to receive GRIT")
    grit_lots = target[target["mail_grit"]].copy()
    lines.append(f"Count: {fmt(len(grit_lots))}")
    lines.append("| Address | Status | Users | Membership status |")
    lines.append("|---|---|---|---|")

    if grit_lots.empty:
        lines.append("| [none] |  |  |  |")
    else:
        grit_lots = grit_lots.sort_values(["Street", "Number", "Unit"], na_position="last")
        for _, row in grit_lots.iterrows():
            address = (clean_text(row.get("Address", "")) or "[blank]").replace("|", "\\|")
            status = (clean_text(row.get("Status", "")) or "[blank]").replace("|", "\\|")
            users = (clean_text(row.get("Users", "")) or "[none]").replace("|", ", ")

            member_value = pd.to_numeric(clean_text(row.get("Is Member", "")), errors="coerce")
            membership_status = "Member" if pd.notna(member_value) and int(member_value) == 1 else "Non-member"

            lines.append(f"| {address} | {status} | {users} | {membership_status} |")
    lines.append("")

    lines.append("## Definitions")
    lines.append("- Vacant lot: `Status = Vacant`.")
    lines.append(
        "- Other-status lot: any `Status` not equal to Owner Occupied, Tenant Occupied, or Vacant (includes Under Construction and blank statuses)."
    )
    lines.append("- Valid email excludes placeholder values such as `fake.fake` and `@fake` domains.")

    output_path.write_text("\n".join(lines), encoding="utf-8")


def run_analysis(data_dir: Path, output_dir: Path, snapshot_date: Optional[str]) -> None:
    resolved_data_dir = resolve_data_dir(data_dir)
    effective_snapshot_date = infer_snapshot_date(snapshot_date, resolved_data_dir)

    addresses_path = resolved_data_dir / ADDRESS_FILE
    addresses_df = pd.read_csv(addresses_path).fillna("")

    analyzed = add_analysis_flags(addresses_df)
    target = analyzed[analyzed["is_other_or_vacant"]].copy()

    status_counts = target["status_clean"].value_counts().sort_values(ascending=False)

    issue_counts = {
        "has_users": int(target["has_users"].sum()),
        "has_valid_email": int(target["has_valid_email"].sum()),
        "marked_member": int(target["is_member"].sum()),
        "receives_grit": int(target["mail_grit"].sum()),
        "has_tenant_data": int(target["has_tenant_text"].sum()),
    }

    issue_flag_total = int(
        (
            target[["has_users", "has_valid_email", "is_member", "mail_grit", "has_tenant_text"]]
            .any(axis=1)
            .sum()
        )
    )

    output_dir.mkdir(parents=True, exist_ok=True)

    trend_output = trend_lines(output_dir, effective_snapshot_date, len(target), issue_flag_total)

    report_path = output_dir / "vacant_other_lots_edge_case_report.md"
    write_report(
        output_path=report_path,
        snapshot_date=effective_snapshot_date,
        resolved_data_dir=resolved_data_dir,
        addresses_path=addresses_path,
        analyzed=analyzed,
        status_counts=status_counts,
        issue_counts=issue_counts,
        trend_output=trend_output,
    )

    print(f"Report written: {report_path}")
    print(f"Total households reviewed: {len(analyzed):,}")
    print(f"Vacant/other lots: {len(target):,}")
    print(f"Lots with at least one edge-case flag: {issue_flag_total:,}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate edge-case report for vacant/other lots from SHHA export snapshots."
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=Path("web_database_exports"),
        help="Directory containing dated export subfolders (YYYY-MM-DD) or a direct export folder.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("analysis/vacant_other_lots"),
        help="Directory to write markdown report and issue CSV files.",
    )
    parser.add_argument(
        "--snapshot-date",
        type=str,
        default=None,
        help="Optional snapshot date override (YYYY-MM-DD). Defaults to dated folder name.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    run_analysis(args.data_dir, args.output_dir, args.snapshot_date)


if __name__ == "__main__":
    main()
