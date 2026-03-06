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


def sanitize_fake_emails(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    cleaned = df.copy()
    for column in columns:
        if column not in cleaned.columns:
            continue
        normalized = cleaned[column].map(normalize_email)
        cleaned[column] = normalized.map(lambda value: value if is_valid_email(value) else "")
    return cleaned


def fmt(value: int) -> str:
    return f"{value:,}"


def pct(part: int, whole: int) -> str:
    if whole == 0:
        return "0.0%"
    return f"{(part / whole) * 100:.1f}%"


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


def norm_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_street(text: object) -> str:
    value = norm_spaces(clean_text(text).lower())
    value = re.sub(r"\balbuquerque\b", "", value)
    value = re.sub(r"\bnew mexico\b", "", value)
    value = re.sub(r"\bnortheast\b", "ne", value)
    value = re.sub(r"\bnorthwest\b", "nw", value)
    value = re.sub(r"\bsoutheast\b", "se", value)
    value = re.sub(r"\bsouthwest\b", "sw", value)
    value = re.sub(r"[^a-z0-9 ]", " ", value)
    return norm_spaces(value)


def normalize_unit(text: object) -> str:
    value = clean_text(text).lower()
    value = re.sub(r"^(unit|apt|#|-)+\s*", "", value)
    value = re.sub(r"[^a-z0-9]", "", value)
    return value


def normalize_number(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""

    numeric = pd.to_numeric(text, errors="coerce")
    if pd.notna(numeric):
        try:
            return str(int(numeric))
        except Exception:
            pass

    match = re.search(r"\d+", text)
    return match.group(0) if match else ""


def split_user_address_field(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []

    for delimiter in [";", "|"]:
        text = text.replace(delimiter, ",")

    return [part.strip() for part in text.split(",") if part.strip()]


def parse_user_address_segment(segment: str) -> tuple[str, str, str]:
    normalized = normalize_street(segment)
    match = re.match(r"^(\d+)\s+(.*)$", normalized)
    if not match:
        return "", "", ""

    number = match.group(1)
    rest = match.group(2).strip()
    unit = ""

    parts = rest.split()
    if parts and re.fullmatch(r"-?[a-z]?\d+[a-z]?|[a-z]", parts[0]):
        unit = normalize_unit(parts[0])
        rest = " ".join(parts[1:]).strip()

    return number, rest, unit


def find_user_export_file(folder: Path) -> Optional[Path]:
    matches = sorted(folder.glob(USER_EXPORT_GLOB), key=lambda path: path.name)
    if not matches:
        return None
    return matches[-1]


def build_household_lookup(addresses_df: pd.DataFrame) -> dict[str, list[dict[str, object]]]:
    working = addresses_df.copy()
    working["street_number_norm"] = working["Number"].map(normalize_number)
    working["street_norm"] = working["Street"].map(normalize_street)
    working["unit_norm"] = working["Unit"].map(normalize_unit)
    working["is_member_num"] = pd.to_numeric(working["Is Member"], errors="coerce").fillna(0).astype(int)

    lookup: dict[str, list[dict[str, object]]] = {}
    for row in working.itertuples(index=False):
        number = clean_text(row.street_number_norm)
        if not number:
            continue

        entry = {
            "household_id": clean_text(row.ID),
            "display_address": clean_text(row.Address),
            "street_norm": clean_text(row.street_norm),
            "unit_norm": clean_text(row.unit_norm),
            "is_member": int(row.is_member_num),
        }
        lookup.setdefault(number, []).append(entry)

    return lookup


def link_user_to_households(users_df: pd.DataFrame, lookup: dict[str, list[dict[str, object]]]) -> pd.DataFrame:
    linked_rows: list[dict[str, object]] = []

    for user in users_df.itertuples(index=False):
        username = clean_text(user.Username)
        segments = split_user_address_field(user.Addresses)

        matched_households: dict[str, dict[str, object]] = {}
        for segment in segments:
            number, street_part, unit_part = parse_user_address_segment(segment)
            if not number:
                continue

            candidates = lookup.get(number, [])
            if not candidates:
                continue

            best_candidate = None
            best_score = -1
            for candidate in candidates:
                street_norm = clean_text(candidate["street_norm"])
                if not street_norm:
                    continue

                if street_norm not in street_part and street_part not in street_norm:
                    continue

                score = len(street_norm)
                unit_norm = clean_text(candidate["unit_norm"])
                if unit_part and unit_norm and unit_part == unit_norm:
                    score += 100
                if unit_part and unit_norm and unit_part != unit_norm:
                    score -= 50

                if score > best_score:
                    best_score = score
                    best_candidate = candidate

            if best_candidate is not None:
                matched_households[clean_text(best_candidate["household_id"])] = best_candidate

        if not matched_households:
            continue

        member_count = sum(1 for h in matched_households.values() if int(h["is_member"]) == 1)
        nonmember_count = sum(1 for h in matched_households.values() if int(h["is_member"]) == 0)

        household_labels = []
        for h in matched_households.values():
            membership_label = "Member" if int(h["is_member"]) == 1 else "Non-member"
            label = f"{clean_text(h['display_address'])} ({membership_label})"
            household_labels.append(label)

        linked_rows.append(
            {
                "username": username,
                "property_count": len(matched_households),
                "member_property_count": member_count,
                "nonmember_property_count": nonmember_count,
                "has_mixed_member_status": member_count > 0 and nonmember_count > 0,
                "household_ids": "|".join(sorted(matched_households.keys())),
                "household_details": " | ".join(sorted(household_labels)),
            }
        )

    if not linked_rows:
        return pd.DataFrame(
            columns=[
                "username",
                "property_count",
                "member_property_count",
                "nonmember_property_count",
                "has_mixed_member_status",
                "household_ids",
                "household_details",
            ]
        )

    linked = pd.DataFrame(linked_rows)
    linked = linked.drop_duplicates(subset=["username", "household_ids"])
    return linked


def write_report(
    output_path: Path,
    snapshot_date: str,
    addresses_df: pd.DataFrame,
    linked_users: pd.DataFrame,
) -> None:
    multi = linked_users[linked_users["property_count"] > 1].copy()
    multi = multi.sort_values(["property_count", "username"], ascending=[False, True])

    total_households = len(addresses_df)
    household_membership = pd.to_numeric(addresses_df["Is Member"], errors="coerce").fillna(0).astype(int)
    total_member_households = int((household_membership == 1).sum())
    total_nonmember_households = int((household_membership == 0).sum())

    multi_household_ids: set[str] = set()
    for ids in multi["household_ids"].tolist():
        for hid in clean_text(ids).split("|"):
            if hid:
                multi_household_ids.add(hid)

    addresses_by_id = addresses_df.copy()
    addresses_by_id["id_text"] = addresses_by_id["ID"].map(clean_text)
    addresses_by_id = addresses_by_id.set_index("id_text", drop=False)

    impacted_member = 0
    impacted_nonmember = 0
    for hid in multi_household_ids:
        if hid not in addresses_by_id.index:
            continue
        is_member_value = pd.to_numeric(clean_text(addresses_by_id.loc[hid, "Is Member"]), errors="coerce")
        if pd.notna(is_member_value) and int(is_member_value) == 1:
            impacted_member += 1
        else:
            impacted_nonmember += 1

    mixed_users = int(multi["has_mixed_member_status"].sum())
    member_only_multi = int(((multi["member_property_count"] > 0) & (multi["nonmember_property_count"] == 0)).sum())
    nonmember_only_multi = int(((multi["member_property_count"] == 0) & (multi["nonmember_property_count"] > 0)).sum())

    lines: list[str] = []
    lines.append("# Users Associated With Multiple Properties")
    lines.append("")
    lines.append(f"Data pull date: {snapshot_date}")
    lines.append("")

    lines.append("## Summary")
    lines.append(f"- Total linked users: {fmt(len(linked_users))}")
    lines.append(
        f"- Users linked to multiple properties: {fmt(len(multi))} ({pct(len(multi), len(linked_users))})"
    )
    lines.append(
        f"- Households touched by multi-property users: {fmt(len(multi_household_ids))} out of {fmt(total_households)} ({pct(len(multi_household_ids), total_households)})"
    )
    lines.append("")

    lines.append("## Membership mix for multi-property users")
    lines.append(f"- Users linked to both member and non-member properties: {fmt(mixed_users)}")
    lines.append(f"- Users linked only to member properties: {fmt(member_only_multi)}")
    lines.append(f"- Users linked only to non-member properties: {fmt(nonmember_only_multi)}")
    lines.append("")

    lines.append("## Household impact by membership status")
    lines.append(
        f"- Member households tied to multi-property users: {fmt(impacted_member)} out of {fmt(total_member_households)} ({pct(impacted_member, total_member_households)})"
    )
    lines.append(
        f"- Non-member households tied to multi-property users: {fmt(impacted_nonmember)} out of {fmt(total_nonmember_households)} ({pct(impacted_nonmember, total_nonmember_households)})"
    )
    lines.append("")

    lines.append("## Full list (all multi-property users)")
    lines.append("| User | Properties | Member properties | Non-member properties | Mixed member/non-member | Linked properties |")
    lines.append("|---|---:|---:|---:|---|---|")

    if multi.empty:
        lines.append("| [none] |  |  |  |  |  |")
    else:
        for _, row in multi.iterrows():
            username = (clean_text(row.get("username", "")) or "[blank]").replace("|", "\\|")
            prop_count = int(row.get("property_count", 0))
            member_count = int(row.get("member_property_count", 0))
            nonmember_count = int(row.get("nonmember_property_count", 0))
            mixed = "Yes" if bool(row.get("has_mixed_member_status", False)) else "No"
            details = (clean_text(row.get("household_details", "")) or "[blank]").replace("|", "; ")
            lines.append(
                f"| {username} | {prop_count} | {member_count} | {nonmember_count} | {mixed} | {details} |"
            )

    lines.append("")
    lines.append("## Interpretation notes")
    lines.append("- This report shows user-to-property associations from website exports, not legal ownership verification.")
    lines.append(
        "- For user-level segment logic in the communication report, a user linked to at least one member property is treated as a member-linked user."
    )

    output_path.write_text("\n".join(lines), encoding="utf-8")


def run_analysis(data_dir: Path, output_dir: Path, snapshot_date: Optional[str]) -> None:
    resolved_data_dir = resolve_data_dir(data_dir)
    effective_snapshot_date = infer_snapshot_date(snapshot_date, resolved_data_dir)

    addresses_path = resolved_data_dir / ADDRESS_FILE
    users_path = find_user_export_file(resolved_data_dir)
    if users_path is None:
        raise FileNotFoundError(f"No users export matching {USER_EXPORT_GLOB} in {resolved_data_dir}")

    addresses_df = pd.read_csv(addresses_path).fillna("")
    users_df = pd.read_csv(users_path).fillna("")
    users_df = sanitize_fake_emails(users_df, ["Email"])

    lookup = build_household_lookup(addresses_df)
    linked_users = link_user_to_households(users_df, lookup)

    output_dir.mkdir(parents=True, exist_ok=True)

    report_path = output_dir / "multi_property_users_report.md"
    write_report(
        output_path=report_path,
        snapshot_date=effective_snapshot_date,
        addresses_df=addresses_df,
        linked_users=linked_users,
    )

    print(f"Report written: {report_path}")
    print(f"Total linked users: {len(linked_users):,}")
    print(f"Users linked to multiple properties: {(linked_users['property_count'] > 1).sum():,}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate report of users associated with multiple properties."
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
        default=Path("analysis/multi_property_users"),
        help="Directory to write markdown report.",
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
