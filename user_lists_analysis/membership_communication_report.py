from __future__ import annotations

import argparse
from datetime import date
import re
from pathlib import Path
from typing import Optional

import pandas as pd

ADDRESS_FILE = "addresses_export.csv"
SUBSCRIBER_FILE = "subscribers.csv"
USER_EXPORT_GLOB = "export-users-*.csv"
TREND_COLUMNS = [
    "run_date",
    "households_total",
    "households_reached_any",
    "households_not_reached",
    "households_only_grit",
    "households_only_email",
    "households_grit_and_email",
    "member_households_total",
    "member_households_not_reached",
    "nonmember_households_total",
    "nonmember_households_not_reached",
    "households_with_users",
    "households_with_users_no_blast",
]


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_email(value: object) -> str:
    return clean_text(value).lower()


def is_valid_email(email: str) -> bool:
    return bool(email) and ("@" in email) and ("fake.fake" not in email)


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


def split_users(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    return [part.strip() for part in text.split(",") if part.strip()]


def split_user_emails(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    for delimiter in [",", ";"]:
        text = text.replace(delimiter, "|")
    emails = [normalize_email(part) for part in text.split("|") if clean_text(part)]
    return [email for email in emails if is_valid_email(email)]


def parse_subscriber_lists(value: object) -> set[str]:
    text = clean_text(value)
    if not text:
        return set()
    return {part.strip().lower() for part in text.split("|") if part.strip()}


def pct(part: int, whole: int) -> str:
    if whole == 0:
        return "0.0%"
    return f"{(part / whole) * 100:.1f}%"


def fmt(value: int) -> str:
    return f"{value:,}"


def signed_int(value: int) -> str:
    return f"{value:+,}"


def format_count_pct(count: int, total: int) -> str:
    return f"{fmt(count)} ({pct(count, total)})"


def find_user_export_file(folder: Path) -> Optional[Path]:
    matches = sorted(folder.glob(USER_EXPORT_GLOB), key=lambda path: path.name)
    if not matches:
        return None
    return matches[-1]


def has_required_exports(folder: Path) -> bool:
    return (
        (folder / ADDRESS_FILE).exists()
        and (folder / SUBSCRIBER_FILE).exists()
        and (find_user_export_file(folder) is not None)
    )


def resolve_data_dir(data_dir: Path) -> Path:
    dated_candidates = [child for child in data_dir.iterdir() if child.is_dir() and has_required_exports(child)]
    if dated_candidates:
        return sorted(dated_candidates, key=lambda path: path.name)[-1]

    if has_required_exports(data_dir):
        return data_dir

    raise FileNotFoundError(
        "No export set found. Expected addresses_export.csv, subscribers.csv, and one export-users-*.csv "
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


def build_subscriber_sets(subscribers_df: pd.DataFrame) -> dict[str, set[str]]:
    working = subscribers_df.copy()
    working["email_norm"] = working["Email"].map(normalize_email)
    working["lists_set"] = working["Lists"].map(parse_subscriber_lists)

    active_mask = working["Unsubscribed At"].map(clean_text).eq("")
    active = working[active_mask].copy()

    frequent = set(
        active.loc[active["lists_set"].map(lambda values: "frequent updates" in values), "email_norm"]
    )
    essential = set(
        active.loc[active["lists_set"].map(lambda values: "essential updates only" in values), "email_norm"]
    )

    frequent = {email for email in frequent if is_valid_email(email)}
    essential = {email for email in essential if is_valid_email(email)}

    return {
        "frequent": frequent,
        "essential": essential,
        "essential_only": essential - frequent,
    }


def build_households(addresses_df: pd.DataFrame, subscriber_sets: dict[str, set[str]]) -> pd.DataFrame:
    households = addresses_df.copy()

    households["household_key"] = households["ID"].map(lambda value: clean_text(value) or "")
    households["display_address"] = households["Address"].map(clean_text)
    households["is_member"] = pd.to_numeric(households["Is Member"], errors="coerce").fillna(0).astype(int)
    households["mail_grit"] = pd.to_numeric(households["Mail GRIT"], errors="coerce").fillna(0).astype(int)

    households["street_number_norm"] = households["Number"].map(normalize_number)
    households["street_norm"] = households["Street"].map(normalize_street)
    households["unit_norm"] = households["Unit"].map(normalize_unit)

    households["linked_user_names"] = households["Users"].map(split_users)
    households["linked_user_count"] = households["linked_user_names"].map(len)
    households["any_user_in_system"] = households["linked_user_count"] > 0

    households["linked_emails"] = households["User Emails"].map(split_user_emails)
    households["linked_valid_email_count"] = households["linked_emails"].map(len)

    frequent = subscriber_sets["frequent"]
    essential = subscriber_sets["essential"]

    households["email_blast_subscriber"] = households["linked_emails"].map(
        lambda emails: any(email in frequent for email in emails)
    )
    households["essential_subscriber"] = households["linked_emails"].map(
        lambda emails: any(email in essential for email in emails)
    )
    households["any_shha_email_subscriber"] = households["email_blast_subscriber"] | households["essential_subscriber"]
    households["no_email_subscribers"] = ~households["any_shha_email_subscriber"]

    households["receives_print_grit"] = households["mail_grit"] == 1
    households["reached_any"] = households["receives_print_grit"] | households["email_blast_subscriber"]
    households["only_grit"] = households["receives_print_grit"] & ~households["email_blast_subscriber"]
    households["only_email"] = households["email_blast_subscriber"] & ~households["receives_print_grit"]
    households["both_grit_email"] = households["receives_print_grit"] & households["email_blast_subscriber"]
    households["neither"] = ~households["reached_any"]

    households["users_no_email_blast_subscriber"] = households["any_user_in_system"] & ~households["email_blast_subscriber"]
    households["users_no_valid_email"] = households["any_user_in_system"] & (households["linked_valid_email_count"] == 0)

    households["blast_email_count"] = households["linked_emails"].map(
        lambda emails: len({email for email in emails if email in frequent})
    )
    households["multi_blast_residents"] = households["blast_email_count"] >= 2

    households["tenant_indicated"] = households["Tenants"].map(lambda value: clean_text(value) != "") | households["Status"].map(
        lambda value: clean_text(value).lower() == "tenant occupied"
    )

    return households


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


def link_users_to_households(users_df: pd.DataFrame, households: pd.DataFrame) -> pd.DataFrame:
    candidates_by_number: dict[str, list[tuple[str, str, str, int]]] = {}
    for row in households.itertuples(index=False):
        number = clean_text(row.street_number_norm)
        if not number:
            continue
        candidates_by_number.setdefault(number, []).append(
            (row.household_key, row.street_norm, row.unit_norm, int(row.is_member))
        )

    link_rows: list[dict[str, object]] = []

    for user in users_df.itertuples(index=False):
        email = normalize_email(user.Email)
        valid_email = is_valid_email(email)

        raw_addresses = clean_text(user.Addresses)
        segments = [part.strip() for part in raw_addresses.split(",") if part.strip()]
        linked_household_keys: set[str] = set()
        linked_membership_flags: set[int] = set()

        for segment in segments:
            number, street_part, unit_part = parse_user_address_segment(segment)
            if not number:
                continue

            candidates = candidates_by_number.get(number, [])
            if not candidates:
                continue

            best_key = None
            best_score = -1
            for household_key, street_norm, unit_norm, is_member in candidates:
                if not street_norm:
                    continue
                if street_norm not in street_part and street_part not in street_norm:
                    continue

                score = len(street_norm)
                if unit_part and unit_norm and unit_part == unit_norm:
                    score += 100
                if unit_part and unit_norm and unit_part != unit_norm:
                    score -= 50

                if score > best_score:
                    best_score = score
                    best_key = (household_key, is_member)

            if best_key is not None:
                linked_household_keys.add(best_key[0])
                linked_membership_flags.add(best_key[1])

        if not linked_household_keys:
            continue

        if 1 in linked_membership_flags:
            segment_type = "member"
        else:
            segment_type = "nonmember"

        link_rows.append(
            {
                "username": clean_text(user.Username),
                "email_norm": email,
                "has_valid_email": valid_email,
                "household_keys": "|".join(sorted(linked_household_keys)),
                "resident_segment": segment_type,
            }
        )

    if not link_rows:
        return pd.DataFrame(columns=["username", "email_norm", "has_valid_email", "household_keys", "resident_segment"])

    linked_users = pd.DataFrame(link_rows)
    linked_users = linked_users.drop_duplicates(subset=["username", "email_norm", "household_keys"])
    return linked_users


def enrich_user_email_status(linked_users: pd.DataFrame, subscriber_sets: dict[str, set[str]]) -> pd.DataFrame:
    users = linked_users.copy()
    frequent = subscriber_sets["frequent"]
    essential = subscriber_sets["essential"]

    users["email_blast_subscriber"] = users["email_norm"].map(lambda email: email in frequent)
    users["essential_subscriber"] = users["email_norm"].map(lambda email: email in essential)
    users["essential_only_subscriber"] = users["essential_subscriber"] & ~users["email_blast_subscriber"]
    users["no_shha_email"] = ~users["email_blast_subscriber"] & ~users["essential_subscriber"]

    return users


def summarize_household_group(df: pd.DataFrame) -> dict[str, int]:
    return {
        "total": int(len(df)),
        "grit": int(df["receives_print_grit"].sum()),
        "blast": int(df["email_blast_subscriber"].sum()),
        "essential": int(df["essential_subscriber"].sum()),
        "both": int(df["both_grit_email"].sum()),
        "neither": int(df["neither"].sum()),
        "reached": int(df["reached_any"].sum()),
        "only_grit": int(df["only_grit"].sum()),
        "only_email": int(df["only_email"].sum()),
        "no_email_subscribers": int(df["no_email_subscribers"].sum()),
        "users": int(df["any_user_in_system"].sum()),
        "users_no_blast": int(df["users_no_email_blast_subscriber"].sum()),
        "users_no_valid_email": int(df["users_no_valid_email"].sum()),
    }


def summarize_individual_group(df: pd.DataFrame) -> dict[str, int]:
    return {
        "total": int(len(df)),
        "blast": int(df["email_blast_subscriber"].sum()),
        "essential_only": int(df["essential_only_subscriber"].sum()),
        "no_shha_email": int(df["no_shha_email"].sum()),
        "valid_email": int(df["has_valid_email"].sum()),
    }


def trend_delta_lines(current: dict[str, int], previous: Optional[dict[str, int]]) -> list[str]:
    if previous is None:
        return ["- No prior snapshot found yet. Run another dated snapshot to see trend deltas."]

    def prev_int(key: str) -> int:
        raw = previous.get(key, 0)
        if pd.isna(raw):
            return 0
        return int(float(raw))

    return [
        f"- Prior snapshot date: {previous['run_date']}",
        f"- Total households: {fmt(current['households_total'])} ({signed_int(current['households_total'] - prev_int('households_total'))})",
        f"- Reached households: {fmt(current['households_reached_any'])} ({signed_int(current['households_reached_any'] - prev_int('households_reached_any'))})",
        f"- Not reached households: {fmt(current['households_not_reached'])} ({signed_int(current['households_not_reached'] - prev_int('households_not_reached'))})",
        f"- Member households not reached: {fmt(current['member_households_not_reached'])} ({signed_int(current['member_households_not_reached'] - prev_int('member_households_not_reached'))})",
        f"- Households with users but no blast subscriber: {fmt(current['households_with_users_no_blast'])} ({signed_int(current['households_with_users_no_blast'] - prev_int('households_with_users_no_blast'))})",
    ]


def write_report(
    output_path: Path,
    resolved_data_dir: Path,
    snapshot_date: str,
    source_paths: dict[str, Path],
    households: pd.DataFrame,
    linked_users: pd.DataFrame,
    household_overall: dict[str, int],
    household_members: dict[str, int],
    household_nonmembers: dict[str, int],
    resident_overall: dict[str, int],
    resident_members: dict[str, int],
    resident_nonmembers: dict[str, int],
    trend_lines: list[str],
) -> None:
    no_linked_users = households[~households["any_user_in_system"]].copy()
    no_linked_user_status = no_linked_users["Status"].map(clean_text).value_counts().head(5)
    no_linked_user_breakdown = ", ".join([f"{status or '[blank]'}: {count}" for status, count in no_linked_user_status.items()])

    tenant_households = int(households["tenant_indicated"].sum())
    members_reached_pct = pct(household_members["reached"], household_members["total"])
    nonmembers_reached_pct = pct(household_nonmembers["reached"], household_nonmembers["total"])

    multi_email_households = int(households["multi_blast_residents"].sum())

    lines: list[str] = []
    lines.append("# SHHA Membership & Communication Reach Report")
    lines.append("")
    lines.append("## Source files")
    lines.append(f"- Data snapshot folder: {resolved_data_dir}")
    lines.append(f"- Snapshot date: {snapshot_date}")
    lines.append(f"- Master addresses: {source_paths['addresses']}")
    lines.append(f"- Users export: {source_paths['users']}")
    lines.append(f"- Subscribers export: {source_paths['subscribers']}")
    lines.append("")
    lines.append("## Definitions")
    lines.append("- **Address / household**: one row in the master addresses export.")
    lines.append("- **User**: one record in the website users export.")
    lines.append("- **Homeowner (analysis definition)**: any user linked to a household; tenants are included.")
    lines.append(f"  - Tenant-indicated households in this snapshot: {fmt(tenant_households)}.")
    lines.append("- **Email blasts**: Frequent Updates list.")
    lines.append("- **Essential-only**: Essential Updates Only without Frequent Updates.")
    lines.append("- **Email subscriber (for email reach in this report)**: active Frequent Updates subscriber (not unsubscribed).")
    lines.append("- **Reached**: household receives at least one of Printed GRIT (`Mail GRIT = 1`) or Email blasts (Frequent Updates).")
    lines.append("- **Realtors list**: excluded from resident communication reach.")
    lines.append("")

    lines.append("## Leadership answers")
    lines.append(f"1. Household reach at all: {format_count_pct(household_overall['reached'], household_overall['total'])}")
    lines.append(f"2. Member households with no communication: {fmt(household_members['neither'])}")
    lines.append(f"3. Households with users but no Frequent Updates subscriber: {fmt(household_overall['users_no_blast'])}")
    lines.append(f"4. Households relying only on GRIT: {fmt(household_overall['only_grit'])}")
    lines.append(f"5. Households relying only on email blasts: {fmt(household_overall['only_email'])}")
    lines.append("")

    lines.append("## Trend since previous snapshot")
    lines.extend(trend_lines)
    lines.append("")

    lines.append("## Data consistency checks")
    lines.append(f"- Total households: {fmt(household_overall['total'])}")
    lines.append(f"- Households with linked users: {fmt(household_overall['users'])}")
    lines.append(f"- Households with no linked users: {fmt(len(no_linked_users))}")
    lines.append(f"- Top statuses among no-linked-user households: {no_linked_user_breakdown}")
    lines.append("")

    total_linked_users = len(linked_users)
    lines.append("## User-level counts")
    lines.append(f"- Linked users (residents linked to addresses): {fmt(total_linked_users)}")
    lines.append(f"- Linked users with valid email: {fmt(resident_overall['valid_email'])}")
    lines.append(f"- Linked users receiving email blasts (Frequent Updates): {fmt(resident_overall['blast'])}")
    denom_users = household_overall["users"] if household_overall["users"] else 0
    users_per_household = total_linked_users / denom_users if denom_users else 0
    lines.append(f"- Users per linked-user household: {users_per_household:.2f}")
    lines.append("")

    lines.append("## Household email coverage")
    lines.append(f"- Households with >=1 email blast subscriber: {format_count_pct(household_overall['blast'], household_overall['total'])}")
    lines.append(f"- Households with >=1 Essential Updates subscriber: {format_count_pct(household_overall['essential'], household_overall['total'])}")
    lines.append(f"- Households with no SHHA email subscribers: {format_count_pct(household_overall['no_email_subscribers'], household_overall['total'])}")
    lines.append(f"- Households with users but no valid email at all: {fmt(household_overall['users_no_valid_email'])}")
    lines.append(f"- Households where multiple residents receive email blasts: {fmt(multi_email_households)}")
    lines.append("")

    lines.append("## GRIT coverage")
    lines.append(f"- Member households receiving printed GRIT: {format_count_pct(household_members['grit'], household_members['total'])}")
    lines.append(f"- Member households not receiving printed GRIT: {fmt(household_members['total'] - household_members['grit'])}")
    lines.append(f"- Member print opt-out / no-print rate: {pct(household_members['total'] - household_members['grit'], household_members['total'])}")
    lines.append("")

    lines.append("## Communication reach by segment")
    lines.append(f"- Member households reached: {fmt(household_members['reached'])} / {fmt(household_members['total'])} ({members_reached_pct})")
    lines.append(f"- Non-member households reached: {fmt(household_nonmembers['reached'])} / {fmt(household_nonmembers['total'])} ({nonmembers_reached_pct})")
    lines.append(f"- Households receiving no SHHA communication: {fmt(household_overall['neither'])} (members: {fmt(household_members['neither'])}, non-members: {fmt(household_nonmembers['neither'])})")
    lines.append("")

    lines.append("## Address-level reach percentages")
    lines.append("| Group | % GRIT | % Email blasts | % Both | % Neither |")
    lines.append("|---|---:|---:|---:|---:|")
    for label, stats in [
        ("Members", household_members),
        ("Non-members", household_nonmembers),
        ("All households", household_overall),
    ]:
        total = stats["total"]
        lines.append(
            f"| {label} | {pct(stats['grit'], total)} | {pct(stats['blast'], total)} | {pct(stats['both'], total)} | {pct(stats['neither'], total)} |"
        )
    lines.append("")

    lines.append("## Individual-level email reach (linked residents)")
    lines.append("| Group | Residents | % Email blasts | % Essential-only | % No SHHA email |")
    lines.append("|---|---:|---:|---:|---:|")
    for label, stats in [
        ("Member-linked residents", resident_members),
        ("Non-member-linked residents", resident_nonmembers),
        ("All linked residents", resident_overall),
    ]:
        total = stats["total"]
        lines.append(
            f"| {label} | {fmt(total)} | {pct(stats['blast'], total)} | {pct(stats['essential_only'], total)} | {pct(stats['no_shha_email'], total)} |"
        )
    lines.append("")

    lines.append("## ASCII Venn (Printed GRIT vs Email blasts/Frequent Updates)")
    lines.append("```")
    lines.append("Printed GRIT vs Email blasts (Frequent Updates)")
    lines.append("=" * 72)
    lines.append(f"  GRIT only   : {fmt(household_overall['only_grit'])}")
    lines.append(f"  Both        : {fmt(household_overall['both'])}")
    lines.append(f"  Email only  : {fmt(household_overall['only_email'])}")
    lines.append(f"  Neither     : {fmt(household_overall['neither'])}")
    lines.append("-" * 72)
    lines.append(
        f"  Reached >=1 : {fmt(household_overall['reached'])} / {fmt(household_overall['total'])} ({pct(household_overall['reached'], household_overall['total'])})"
    )
    lines.append("```")
    lines.append("")

    lines.append("## Interpretation notes")
    lines.append("- 'Households with users but no email subscriber' means no linked email at that address is actively subscribed to Frequent Updates.")
    lines.append("- Such households may still have residents on Essential Updates Only; those are counted separately in household/individual essential metrics.")
    lines.append("- Resident-level metrics are derived from users linked to addresses via address parsing and matching rules.")

    output_path.write_text("\n".join(lines), encoding="utf-8")


def run_report(data_dir: Path, output_dir: Path, snapshot_date: Optional[str]) -> None:
    resolved_data_dir = resolve_data_dir(data_dir)
    effective_snapshot_date = infer_snapshot_date(snapshot_date, resolved_data_dir)

    addresses_path = resolved_data_dir / ADDRESS_FILE
    users_path = find_user_export_file(resolved_data_dir)
    subscribers_path = resolved_data_dir / SUBSCRIBER_FILE

    if users_path is None:
        raise FileNotFoundError(f"No users export matching {USER_EXPORT_GLOB} in {resolved_data_dir}")

    addresses_df = pd.read_csv(addresses_path).fillna("")
    users_df = pd.read_csv(users_path).fillna("")
    subscribers_df = pd.read_csv(subscribers_path).fillna("")

    subscriber_sets = build_subscriber_sets(subscribers_df)
    households = build_households(addresses_df, subscriber_sets)

    linked_users = link_users_to_households(users_df, households)
    linked_users = enrich_user_email_status(linked_users, subscriber_sets)

    households_members = households[households["is_member"] == 1]
    households_nonmembers = households[households["is_member"] == 0]

    residents_members = linked_users[linked_users["resident_segment"] == "member"]
    residents_nonmembers = linked_users[linked_users["resident_segment"] == "nonmember"]

    household_overall = summarize_household_group(households)
    household_members = summarize_household_group(households_members)
    household_nonmembers = summarize_household_group(households_nonmembers)

    resident_overall = summarize_individual_group(linked_users)
    resident_members = summarize_individual_group(residents_members)
    resident_nonmembers = summarize_individual_group(residents_nonmembers)

    output_dir.mkdir(parents=True, exist_ok=True)

    export_households = households.copy()
    export_households["linked_user_names"] = export_households["linked_user_names"].map(lambda values: "|".join(values))
    export_households["linked_emails"] = export_households["linked_emails"].map(lambda values: "|".join(values))
    export_households.to_csv(output_dir / "household_communication_flags.csv", index=False)

    linked_users.to_csv(output_dir / "linked_user_email_reach.csv", index=False)
    households[households["neither"]].to_csv(output_dir / "households_not_reached.csv", index=False)
    households[households["users_no_email_blast_subscriber"]].to_csv(
        output_dir / "households_with_users_but_no_blast_subscriber.csv", index=False
    )

    trend_snapshot = {
        "run_date": effective_snapshot_date,
        "households_total": household_overall["total"],
        "households_reached_any": household_overall["reached"],
        "households_not_reached": household_overall["neither"],
        "households_only_grit": household_overall["only_grit"],
        "households_only_email": household_overall["only_email"],
        "households_grit_and_email": household_overall["both"],
        "member_households_total": household_members["total"],
        "member_households_not_reached": household_members["neither"],
        "nonmember_households_total": household_nonmembers["total"],
        "nonmember_households_not_reached": household_nonmembers["neither"],
        "households_with_users": household_overall["users"],
        "households_with_users_no_blast": household_overall["users_no_blast"],
    }

    trend_path = output_dir / "trend_history.csv"
    if trend_path.exists():
        trend_df = pd.read_csv(trend_path)
    else:
        trend_df = pd.DataFrame(columns=TREND_COLUMNS)

    for column in TREND_COLUMNS:
        if column not in trend_df.columns:
            trend_df[column] = pd.NA

    trend_df = trend_df[TREND_COLUMNS]

    prior_rows = trend_df[trend_df["run_date"].astype(str) < trend_snapshot["run_date"]].copy()
    previous_snapshot = None
    if not prior_rows.empty:
        prior_rows = prior_rows.sort_values("run_date")
        previous_snapshot = prior_rows.iloc[-1].to_dict()

    trend_df = trend_df[trend_df["run_date"].astype(str) != trend_snapshot["run_date"]]
    trend_df = pd.concat([trend_df, pd.DataFrame([trend_snapshot])], ignore_index=True)
    trend_df = trend_df.sort_values("run_date")

    numeric_columns = [column for column in TREND_COLUMNS if column != "run_date"]
    for column in numeric_columns:
        trend_df[column] = pd.to_numeric(trend_df[column], errors="coerce").fillna(0).astype(int)

    trend_df.to_csv(trend_path, index=False)

    trend_lines = trend_delta_lines(trend_snapshot, previous_snapshot)

    report_path = output_dir / "membership_communication_report.md"
    write_report(
        output_path=report_path,
        resolved_data_dir=resolved_data_dir,
        snapshot_date=effective_snapshot_date,
        source_paths={"addresses": addresses_path, "users": users_path, "subscribers": subscribers_path},
        households=households,
        linked_users=linked_users,
        household_overall=household_overall,
        household_members=household_members,
        household_nonmembers=household_nonmembers,
        resident_overall=resident_overall,
        resident_members=resident_members,
        resident_nonmembers=resident_nonmembers,
        trend_lines=trend_lines,
    )

    print(f"Report written: {report_path}")
    print(f"Trend snapshot updated: {trend_path}")
    print(f"Total households: {household_overall['total']:,}")
    print(f"Reached households: {household_overall['reached']:,} ({pct(household_overall['reached'], household_overall['total'])})")
    print(f"Member households not reached: {household_members['neither']:,}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate SHHA membership & communication report from dated export snapshots."
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
        default=Path("analysis/membership_communication_report"),
        help="Directory to write report and output tables",
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
    run_report(args.data_dir, args.output_dir, args.snapshot_date)


if __name__ == "__main__":
    main()
