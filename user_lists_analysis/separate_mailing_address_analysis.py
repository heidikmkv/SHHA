from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import re
from typing import Optional

import pandas as pd

ADDRESS_FILE = "addresses_export.csv"
USER_EXPORT_GLOB = "export-users-*.csv"


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


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


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def normalize_address(value: object) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"\b(\d+)-(\d+)\b", r"\1 unit \2", text)
    text = text.replace("#", " unit ")
    text = re.sub(r"\b(apartment|apt|unit|suite|ste)\b", " unit ", text)
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    text = normalize_spaces(text)
    text = re.sub(r"\b(northeast|ne)\b", "ne", text)
    text = re.sub(r"\b(northwest|nw)\b", "nw", text)
    text = re.sub(r"\b(southeast|se)\b", "se", text)
    text = re.sub(r"\b(southwest|sw)\b", "sw", text)
    return normalize_spaces(text)


def normalize_address_base_without_unit(value: object) -> str:
    text = normalize_address(value)
    text = re.sub(r"\bunit\s+[a-z0-9\-]+", "", text)
    return normalize_spaces(text)


def extract_number_tokens(value: object) -> list[str]:
    text = normalize_address(value)
    return re.findall(r"\b\d+\b", text)


def extract_alpha_tokens(value: object) -> list[str]:
    text = normalize_address(value)
    words = re.findall(r"\b[a-z]+\b", text)
    drop = {"unit", "ne", "nw", "se", "sw", "n", "s", "e", "w"}
    return [word for word in words if word not in drop]


def build_mailing_full(row: pd.Series) -> str:
    mail_address = clean_text(row.get("Mail Address", ""))
    mail_city = clean_text(row.get("Mail City", ""))
    mail_state = clean_text(row.get("Mail State", ""))
    mail_zip = clean_text(row.get("Mail Postal Code", ""))
    return ", ".join(part for part in [mail_address, mail_city, mail_state, mail_zip] if part)


def likely_same_property_address(mail_address_only: str, property_address: str) -> bool:
    if not mail_address_only or not property_address:
        return False

    mail_norm = normalize_address(mail_address_only)
    prop_norm = normalize_address(property_address)
    if mail_norm == prop_norm:
        return True

    mail_base = normalize_address_base_without_unit(mail_address_only)
    prop_base = normalize_address_base_without_unit(property_address)
    if bool(mail_base) and bool(prop_base) and mail_base == prop_base:
        return True

    mail_numbers = extract_number_tokens(mail_address_only)
    prop_numbers = extract_number_tokens(property_address)
    mail_alpha = extract_alpha_tokens(mail_address_only)
    prop_alpha = extract_alpha_tokens(property_address)

    if not mail_numbers or not prop_numbers:
        return False

    if mail_numbers[0] != prop_numbers[0]:
        return False

    if mail_alpha != prop_alpha:
        return False

    if len(mail_numbers) == 1 and len(prop_numbers) == 1:
        return True

    return set(mail_numbers[1:]) == set(prop_numbers[1:])


def add_flags(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    out["address_clean"] = out["Address"].map(clean_text)
    out["status_clean"] = out["Status"].map(clean_text)
    out["status_norm"] = out["status_clean"].str.lower()
    out["users_clean"] = out["Users"].map(clean_text)
    out["tenants_clean"] = out["Tenants"].map(clean_text)

    out["is_member"] = pd.to_numeric(out["Is Member"], errors="coerce").fillna(0).astype(int) == 1
    out["mail_grit"] = pd.to_numeric(out["Mail GRIT"], errors="coerce").fillna(0).astype(int) == 1

    out["mail_address_only"] = out["Mail Address"].map(clean_text)
    out["mail_city"] = out["Mail City"].map(clean_text)
    out["mail_state"] = out["Mail State"].map(clean_text)
    out["mail_zip"] = out["Mail Postal Code"].map(clean_text)
    out["mailing_full"] = out.apply(build_mailing_full, axis=1)

    out["has_mailing_address"] = out["mail_address_only"] != ""
    out["likely_same_property"] = out.apply(
        lambda row: likely_same_property_address(row["mail_address_only"], row["address_clean"]), axis=1
    )

    out["is_separate_mailing"] = out["has_mailing_address"] & ~out["likely_same_property"]

    out["is_po_box"] = out["mail_address_only"].str.lower().str.contains(r"\b(?:p\.?\s*o\.?\s*box|post\s+office\s+box)\b", regex=True)

    out["mail_state_norm"] = out["mail_state"].str.upper()
    out["is_out_of_state"] = (out["mail_state_norm"] != "") & (out["mail_state_norm"] != "NM")

    out["is_tenant_occupied"] = out["status_norm"] == "tenant occupied"
    out["has_tenant_text"] = out["tenants_clean"] != ""
    out["likely_landlord"] = out["is_tenant_occupied"] | out["has_tenant_text"]

    out["is_owner_occupied"] = out["status_norm"] == "owner occupied"

    out["possible_outdated_forwarding"] = (
        out["is_separate_mailing"]
        & out["is_owner_occupied"]
        & ~out["likely_landlord"]
        & ~out["is_out_of_state"]
        & ~out["is_po_box"]
    )

    out["has_users"] = out["users_clean"] != ""

    return out


def trend_lines(output_dir: Path, snapshot_date: str, separate_true_count: int, outdated_count: int) -> list[str]:
    trend_path = output_dir / "separate_mailing_trend_history.csv"
    current = {
        "run_date": snapshot_date,
        "true_separate_mailing": separate_true_count,
        "possible_outdated_forwarding": outdated_count,
    }

    if trend_path.exists():
        trend_df = pd.read_csv(trend_path)
    else:
        trend_df = pd.DataFrame(columns=["run_date", "true_separate_mailing", "possible_outdated_forwarding"])

    for col in ["run_date", "true_separate_mailing", "possible_outdated_forwarding"]:
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

    for col in ["true_separate_mailing", "possible_outdated_forwarding"]:
        trend_df[col] = pd.to_numeric(trend_df[col], errors="coerce").fillna(0).astype(int)

    trend_df.to_csv(trend_path, index=False)

    if previous_snapshot is None:
        return ["- No prior snapshot yet for trend comparison."]

    prev_true = int(float(previous_snapshot.get("true_separate_mailing", 0)))
    prev_outdated = int(float(previous_snapshot.get("possible_outdated_forwarding", 0)))
    prior_date = str(previous_snapshot.get("run_date", "unknown"))

    return [
        f"Prior snapshot date: {prior_date}",
        f"- True separate mailing addresses: {fmt(separate_true_count)} ({signed_int(separate_true_count - prev_true)})",
        f"- Possible outdated forwarding cases: {fmt(outdated_count)} ({signed_int(outdated_count - prev_outdated)})",
    ]


def reason_flags(row: pd.Series) -> str:
    flags: list[str] = []
    if row.get("possible_outdated_forwarding", False):
        flags.append("possible outdated forwarding")
    if row.get("likely_landlord", False):
        flags.append("likely landlord/rental")
    if row.get("is_out_of_state", False):
        flags.append("out-of-state mailing")
    if row.get("is_po_box", False):
        flags.append("PO box")
    if not flags:
        flags.append("other separate mailing")
    return "; ".join(flags)


def write_report(
    output_path: Path,
    snapshot_date: str,
    resolved_data_dir: Path,
    addresses_path: Path,
    analyzed: pd.DataFrame,
    trend_output: list[str],
) -> None:
    total = len(analyzed)
    with_mail = analyzed[analyzed["has_mailing_address"]].copy()
    same_property = with_mail[with_mail["likely_same_property"]].copy()
    true_separate = analyzed[analyzed["is_separate_mailing"]].copy()

    status_counts = true_separate["status_clean"].value_counts().head(10)

    lines: list[str] = []
    lines.append("# Separate Mailing Address Review Report")
    lines.append("")
    lines.append(f"Data pull date: {snapshot_date}")
    lines.append(f"Source folder: {resolved_data_dir}")
    lines.append(f"Source file: {addresses_path}")
    lines.append("")

    lines.append("## Purpose")
    lines.append(
        "Find and categorize addresses with separate mailing information so office admin can focus on likely outdated forwarding and data-cleanup priorities."
    )
    lines.append("")

    lines.append("## Summary")
    lines.append(f"- Total households reviewed: {fmt(total)}")
    lines.append(
        f"- Households with mailing address entered: {fmt(len(with_mail))} ({pct(len(with_mail), total)})"
    )
    lines.append(
        f"- Likely same-property format/unit issues (not truly separate): {fmt(len(same_property))} ({pct(len(same_property), len(with_mail))})"
    )
    lines.append(
        f"- Likely true separate mailing addresses: {fmt(len(true_separate))} ({pct(len(true_separate), total)})"
    )
    lines.append("")

    lines.append("## Key categories within true separate mailing")
    lines.append(
        f"- Possible outdated forwarding (owner occupied + local + not landlord): {fmt(int(true_separate['possible_outdated_forwarding'].sum()))}"
    )
    lines.append(f"- Likely landlord/rental cases: {fmt(int(true_separate['likely_landlord'].sum()))}")
    lines.append(f"- Out-of-state mailing addresses: {fmt(int(true_separate['is_out_of_state'].sum()))}")
    lines.append(f"- PO Box mailing addresses: {fmt(int(true_separate['is_po_box'].sum()))}")
    lines.append("")

    lines.append("## Status breakdown (true separate mailing)")
    if status_counts.empty:
        lines.append("- None")
    else:
        for status, count in status_counts.items():
            label = status if status else "[blank]"
            lines.append(f"- {label}: {fmt(int(count))}")
    lines.append("")

    lines.append("## Change since previous data pull")
    lines.extend(trend_output)
    lines.append("")

    lines.append("## High-priority review sample (first 80)")
    lines.append("This table is limited for readability. Use the CSV for the full review list.")
    lines.append("| Address | Status | Users | Mailing address | Reason flags |")
    lines.append("|---|---|---|---|---|")

    review = true_separate.copy()
    review["reason_flags"] = review.apply(reason_flags, axis=1)

    review["priority_score"] = (
        review["possible_outdated_forwarding"].astype(int) * 8
        + review["is_owner_occupied"].astype(int) * 4
        + review["has_users"].astype(int) * 2
        + review["is_member"].astype(int)
    )

    review = review.sort_values(["priority_score", "Address"], ascending=[False, True]).head(80)

    if review.empty:
        lines.append("| [none] |  |  |  |  |")
    else:
        for _, row in review.iterrows():
            address = (clean_text(row.get("Address", "")) or "[blank]").replace("|", "\\|")
            status = (clean_text(row.get("Status", "")) or "[blank]").replace("|", "\\|")
            users = (clean_text(row.get("Users", "")) or "[none]").replace("|", ", ")
            mailing = (clean_text(row.get("mailing_full", "")) or "[blank]").replace("|", "\\|")
            reason = (clean_text(row.get("reason_flags", "")) or "[blank]").replace("|", "\\|")
            lines.append(f"| {address} | {status} | {users} | {mailing} | {reason} |")
    lines.append("")

    lines.append("## Full review output")
    lines.append("- separate_mailing_address_review.csv")
    lines.append("")

    lines.append("## Definitions")
    lines.append("- True separate mailing: mailing address exists and does not normalize to the same property address.")
    lines.append("- Same-property format/unit issue: appears separate in raw text but normalizes to same base address.")
    lines.append("- Possible outdated forwarding: owner occupied with local separate mailing address and no obvious landlord indicators.")

    output_path.write_text("\n".join(lines), encoding="utf-8")


def run_analysis(data_dir: Path, output_dir: Path, snapshot_date: Optional[str]) -> None:
    resolved_data_dir = resolve_data_dir(data_dir)
    effective_snapshot_date = infer_snapshot_date(snapshot_date, resolved_data_dir)

    addresses_path = resolved_data_dir / ADDRESS_FILE
    addresses_df = pd.read_csv(addresses_path).fillna("")

    analyzed = add_flags(addresses_df)

    true_separate = analyzed[analyzed["is_separate_mailing"]].copy()

    output_dir.mkdir(parents=True, exist_ok=True)

    review_export = true_separate.copy()
    review_export["reason_flags"] = review_export.apply(reason_flags, axis=1)
    export_columns = [
        "ID",
        "Address",
        "Status",
        "Users",
        "Is Member",
        "Mail GRIT",
        "Mail Address",
        "Mail City",
        "Mail State",
        "Mail Postal Code",
        "mailing_full",
        "likely_same_property",
        "is_separate_mailing",
        "possible_outdated_forwarding",
        "likely_landlord",
        "is_out_of_state",
        "is_po_box",
        "reason_flags",
    ]
    review_export[export_columns].to_csv(output_dir / "separate_mailing_address_review.csv", index=False)

    trend_output = trend_lines(
        output_dir=output_dir,
        snapshot_date=effective_snapshot_date,
        separate_true_count=len(true_separate),
        outdated_count=int(true_separate["possible_outdated_forwarding"].sum()),
    )

    report_path = output_dir / "separate_mailing_address_report.md"
    write_report(
        output_path=report_path,
        snapshot_date=effective_snapshot_date,
        resolved_data_dir=resolved_data_dir,
        addresses_path=addresses_path,
        analyzed=analyzed,
        trend_output=trend_output,
    )

    print(f"Report written: {report_path}")
    print(f"Review CSV written: {output_dir / 'separate_mailing_address_review.csv'}")
    print(f"Total households reviewed: {len(analyzed):,}")
    print(f"Likely true separate mailing addresses: {len(true_separate):,}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate separate-mailing-address analysis report for office-admin review."
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
        default=Path("analysis/separate_mailing_address"),
        help="Directory to write markdown report and one review CSV.",
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
