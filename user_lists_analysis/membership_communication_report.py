from __future__ import annotations

import argparse
from datetime import date
import re
from pathlib import Path
from typing import Optional

import pandas as pd

ADDRESS_FILE = "addresses_export.csv"
SUBSCRIBER_FILE = "subscribers.csv"
ADDRESS_USERS_FILE = "address_users_export.csv"
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


def split_user_address_field(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    for delimiter in [";", "|"]:
        text = text.replace(delimiter, ",")
    return [part.strip() for part in text.split(",") if part.strip()]


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


def inline_breakdown(total_value: int, member_value: int, nonmember_value: int) -> str:
    return f"{fmt(total_value)} (members: {fmt(member_value)}, non-members: {fmt(nonmember_value)})"


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
    valid_mask = working["email_norm"].map(is_valid_email)
    working = working[valid_mask].copy()

    active_mask = working["Unsubscribed At"].map(clean_text).eq("")
    active = working[active_mask].copy()
    unsubscribed = working[~active_mask].copy()

    frequent = set(
        active.loc[active["lists_set"].map(lambda values: "frequent updates" in values), "email_norm"]
    )
    essential = set(
        active.loc[active["lists_set"].map(lambda values: "essential updates only" in values), "email_norm"]
    )
    grit_edelivery = set(
        active.loc[active["lists_set"].map(lambda values: "grit e-delivery" in values), "email_norm"]
    )

    frequent = {email for email in frequent if is_valid_email(email)}
    essential = {email for email in essential if is_valid_email(email)}
    grit_edelivery = {email for email in grit_edelivery if is_valid_email(email)}
    all_known = set(working["email_norm"])
    unsubscribed_all = set(unsubscribed["email_norm"])
    active_nonblast = set(active["email_norm"]) - frequent

    return {
        "frequent": frequent,
        "essential": essential,
        "grit_edelivery": grit_edelivery,
        "essential_only": essential - frequent,
        "all_known": all_known,
        "unsubscribed": unsubscribed_all,
        "active_nonblast": active_nonblast,
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


def build_email_membership_lookup(linked_users: pd.DataFrame) -> dict[str, str]:
    if linked_users.empty:
        return {}

    membership: dict[str, str] = {}
    grouped = linked_users.groupby("email_norm")
    for email, group in grouped:
        if not is_valid_email(email):
            continue
        segments = set(group["resident_segment"].tolist())
        if "member" in segments:
            membership[email] = "known_members"
        elif "nonmember" in segments:
            membership[email] = "known_nonmembers"
        else:
            membership[email] = "other"
    return membership


def compute_list_breakdown(email_set: set[str], membership_lookup: dict[str, str]) -> dict[str, int]:
    known_members = sum(1 for email in email_set if membership_lookup.get(email) == "known_members")
    known_nonmembers = sum(1 for email in email_set if membership_lookup.get(email) == "known_nonmembers")
    other = len(email_set) - known_members - known_nonmembers
    return {
        "total": len(email_set),
        "known_members": known_members,
        "known_nonmembers": known_nonmembers,
        "other": other,
    }


def compute_homeowner_db_not_subscribed_stats(
    linked_users: pd.DataFrame, subscriber_sets: dict[str, set[str]]
) -> dict[str, int]:
    if linked_users.empty:
        return {
            "total": 0,
            "members": 0,
            "nonmembers": 0,
            "never_subscribed": 0,
            "unsubscribed": 0,
            "active_nonblast": 0,
            "other_status": 0,
        }

    valid_users = linked_users[linked_users["has_valid_email"]].copy()
    not_blast = valid_users[~valid_users["email_blast_subscriber"]].copy()

    known = subscriber_sets["all_known"]
    unsubscribed = subscriber_sets["unsubscribed"]
    active_nonblast = subscriber_sets["active_nonblast"]

    def classify(email: str) -> str:
        if email not in known:
            return "never_subscribed"
        if email in unsubscribed:
            return "unsubscribed"
        if email in active_nonblast:
            return "active_nonblast"
        return "other_status"

    not_blast["reason"] = not_blast["email_norm"].map(classify)

    return {
        "total": int(len(not_blast)),
        "members": int((not_blast["resident_segment"] == "member").sum()),
        "nonmembers": int((not_blast["resident_segment"] == "nonmember").sum()),
        "never_subscribed": int((not_blast["reason"] == "never_subscribed").sum()),
        "unsubscribed": int((not_blast["reason"] == "unsubscribed").sum()),
        "active_nonblast": int((not_blast["reason"] == "active_nonblast").sum()),
        "other_status": int((not_blast["reason"] == "other_status").sum()),
    }


def compute_household_composition(households: pd.DataFrame) -> dict[str, float]:
    user_counts = households["linked_user_count"]
    return {
        "zero_users": int((user_counts == 0).sum()),
        "one_user": int((user_counts == 1).sum()),
        "two_users": int((user_counts == 2).sum()),
        "three_plus_users": int((user_counts >= 3).sum()),
        "avg_users_per_address": float(user_counts.mean()) if len(user_counts) else 0.0,
    }


def compute_users_associated_stats(users_df: pd.DataFrame) -> dict[str, int]:
    address_counts = users_df["Addresses"].map(split_user_address_field).map(len)
    return {
        "zero_addresses": int((address_counts == 0).sum()),
        "one_address": int((address_counts == 1).sum()),
        "two_addresses": int((address_counts == 2).sum()),
        "three_addresses": int((address_counts == 3).sum()),
        "four_addresses": int((address_counts == 4).sum()),
        "five_plus_addresses": int((address_counts >= 5).sum()),
    }


def compute_user_roles(address_users_df: Optional[pd.DataFrame]) -> dict[str, int]:
    if address_users_df is None or address_users_df.empty:
        return {"owner": 0, "tenant": 0, "other": 0}

    role_norm = address_users_df["Role"].map(clean_text).str.lower()
    return {
        "owner": int((role_norm == "owner").sum()),
        "tenant": int((role_norm == "tenant").sum()),
        "other": int((role_norm == "other").sum()),
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
        f"Prior snapshot date: {previous['run_date']}",
        f"- Total households: {fmt(current['households_total'])} ({signed_int(current['households_total'] - prev_int('households_total'))})",
        f"- Reached households: {fmt(current['households_reached_any'])} ({signed_int(current['households_reached_any'] - prev_int('households_reached_any'))})",
        f"- Not reached households: {fmt(current['households_not_reached'])} ({signed_int(current['households_not_reached'] - prev_int('households_not_reached'))})",
        f"- Member households not reached: {fmt(current['member_households_not_reached'])} ({signed_int(current['member_households_not_reached'] - prev_int('member_households_not_reached'))})",
        f"- Households with users but no blast subscriber: {fmt(current['households_with_users_no_blast'])} ({signed_int(current['households_with_users_no_blast'] - prev_int('households_with_users_no_blast'))})",
    ]


def previous_int(previous_snapshot: Optional[dict[str, int]], key: str) -> int:
    if previous_snapshot is None:
        return 0

    value = previous_snapshot.get(key, 0)
    if pd.isna(value):
        return 0

    return int(float(value))


def write_executive_summary(
    output_path: Path,
    snapshot_date: str,
    household_overall: dict[str, int],
    household_members: dict[str, int],
    household_nonmembers: dict[str, int],
    resident_overall: dict[str, int],
    homeowner_db_not_subscribed: dict[str, int],
    current_trend: dict[str, int],
    previous_snapshot: Optional[dict[str, int]],
) -> None:
    member_reach_pct = pct(household_members["reached"], household_members["total"])
    nonmember_reach_pct = pct(household_nonmembers["reached"], household_nonmembers["total"])
    overall_reach_pct = pct(household_overall["reached"], household_overall["total"])
    member_share_pct = pct(household_members["total"], household_overall["total"])
    nonmember_share_pct = pct(household_nonmembers["total"], household_overall["total"])

    reached_delta = 0
    member_not_reached_delta = 0
    users_no_blast_delta = 0
    has_previous = previous_snapshot is not None
    if has_previous:
        reached_delta = current_trend["households_reached_any"] - previous_int(previous_snapshot, "households_reached_any")
        member_not_reached_delta = current_trend["member_households_not_reached"] - previous_int(
            previous_snapshot, "member_households_not_reached"
        )
        users_no_blast_delta = current_trend["households_with_users_no_blast"] - previous_int(
            previous_snapshot, "households_with_users_no_blast"
        )

    lines: list[str] = []
    lines.append("# SHHA Board Executive Summary")
    lines.append("")
    lines.append(f"Data pull date: {snapshot_date}")
    lines.append("")

    lines.append("## Membership snapshot")
    lines.append(f"- Total households in membership database: {fmt(household_overall['total'])}")
    lines.append(
        f"- Member households: {fmt(household_members['total'])} ({member_share_pct})"
    )
    lines.append(
        f"- Non-member households: {fmt(household_nonmembers['total'])} ({nonmember_share_pct})"
    )
    lines.append("")

    lines.append("## Communication snapshot")
    lines.append(
        f"- Households reached by printed GRIT and/or email blasts: {fmt(household_overall['reached'])} out of {fmt(household_overall['total'])} ({overall_reach_pct})"
    )
    lines.append(
        f"- Member household reach: {fmt(household_members['reached'])} out of {fmt(household_members['total'])} ({member_reach_pct})"
    )
    lines.append(
        f"- Non-member household reach: {fmt(household_nonmembers['reached'])} out of {fmt(household_nonmembers['total'])} ({nonmember_reach_pct})"
    )
    lines.append("")

    lines.append("## Key numbers")
    lines.append(
        f"- Member households not reached: {fmt(household_members['neither'])} out of {fmt(household_members['total'])} ({pct(household_members['neither'], household_members['total'])})"
    )
    lines.append(
        f"- Non-member households not reached: {fmt(household_nonmembers['neither'])} out of {fmt(household_nonmembers['total'])} ({pct(household_nonmembers['neither'], household_nonmembers['total'])})"
    )
    lines.append(
        f"- Channel mix among all households: both {fmt(household_overall['both'])}, email-only {fmt(household_overall['only_email'])}, GRIT-only {fmt(household_overall['only_grit'])}"
    )
    lines.append(
        f"- Linked residents not receiving SHHA email: {fmt(resident_overall['no_shha_email'])} out of {fmt(resident_overall['total'])} ({pct(resident_overall['no_shha_email'], resident_overall['total'])})"
    )
    lines.append(
        f"- Homeowner DB users with valid email not on email blasts: {fmt(homeowner_db_not_subscribed['total'])} (never subscribed {fmt(homeowner_db_not_subscribed['never_subscribed'])}, unsubscribed {fmt(homeowner_db_not_subscribed['unsubscribed'])})"
    )
    lines.append("")

    lines.append("## Change since previous data pull")
    if has_previous:
        prior_date = str(previous_snapshot.get("run_date", "unknown"))
        lines.append(f"Prior snapshot date: {prior_date}")
        lines.append(f"- Reached households change: {signed_int(reached_delta)}")
        lines.append(f"- Member households not reached change: {signed_int(member_not_reached_delta)}")
        lines.append(f"- Households with users but no blast subscriber change: {signed_int(users_no_blast_delta)}")
    else:
        lines.append("- No prior snapshot exists yet for change comparison.")
    lines.append("")

    lines.append("## Definitions used in this summary")
    lines.append("- Linked residents: users from the website users export that were matched to at least one household address.")
    lines.append("- Email blasts: active Frequent Updates subscribers.")
    lines.append("- Not reached: households receiving neither printed GRIT nor email blasts.")

    output_path.write_text("\n".join(lines), encoding="utf-8")


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
    household_composition: dict[str, float],
    users_associated_stats: dict[str, int],
    user_roles: dict[str, int],
    list_breakdowns: dict[str, dict[str, int]],
    homeowner_db_not_subscribed: dict[str, int],
    trend_lines: list[str],
) -> None:
    no_linked_users = households[~households["any_user_in_system"]].copy()
    no_linked_user_status = no_linked_users["Status"].map(clean_text).value_counts().head(5)
    no_linked_user_breakdown = ", ".join([f"{status or '[blank]'}: {count}" for status, count in no_linked_user_status.items()])

    tenant_households = int(households["tenant_indicated"].sum())
    members_reached_pct = pct(household_members["reached"], household_members["total"])
    nonmembers_reached_pct = pct(household_nonmembers["reached"], household_nonmembers["total"])

    multi_email_households = int(households["multi_blast_residents"].sum())

    member_households_df = households[households["is_member"] == 1]
    nonmember_households_df = households[households["is_member"] == 0]

    members_no_users = int((~member_households_df["any_user_in_system"]).sum())
    nonmembers_no_users = int((~nonmember_households_df["any_user_in_system"]).sum())

    members_users_no_blast = int(member_households_df["users_no_email_blast_subscriber"].sum())
    nonmembers_users_no_blast = int(nonmember_households_df["users_no_email_blast_subscriber"].sum())

    members_users_no_valid_email = int(member_households_df["users_no_valid_email"].sum())
    nonmembers_users_no_valid_email = int(nonmember_households_df["users_no_valid_email"].sum())

    members_multi_blast = int(member_households_df["multi_blast_residents"].sum())
    nonmembers_multi_blast = int(nonmember_households_df["multi_blast_residents"].sum())

    reached_total = household_overall["reached"]
    email_only_share = pct(household_overall["only_email"], reached_total)
    print_only_share = pct(household_overall["only_grit"], reached_total)
    both_share = pct(household_overall["both"], reached_total)

    lines: list[str] = []
    lines.append("# SHHA Membership & Communication Reach Report")
    lines.append("")
    lines.append("## Table of Contents")
    lines.append("- [1. Source files](#1-source-files)")
    lines.append("- [2. Definitions](#2-definitions)")
    lines.append("- [3. Leadership answers](#3-leadership-answers)")
    lines.append("  - [3.1 Core leadership questions](#31-core-leadership-questions)")
    lines.append("  - [3.2 User-level leadership snapshot](#32-user-level-leadership-snapshot)")
    lines.append("- [4. Trend since previous snapshot](#4-trend-since-previous-snapshot)")
    lines.append("- [5. Technical deep dive](#5-technical-deep-dive)")
    lines.append("  - [5.1 Data consistency checks](#51-data-consistency-checks)")
    lines.append("  - [5.2 Household composition](#52-household-composition)")
    lines.append("  - [5.3 Users associated with addresses](#53-users-associated-with-addresses)")
    lines.append("  - [5.4 User roles](#54-user-roles)")
    lines.append("  - [5.5 User-level counts](#55-user-level-counts)")
    lines.append("  - [5.6 Email list membership snapshot](#56-email-list-membership-snapshot)")
    lines.append("  - [5.7 Household email coverage](#57-household-email-coverage)")
    lines.append("  - [5.8 GRIT coverage](#58-grit-coverage)")
    lines.append("  - [5.9 Communication reach by segment](#59-communication-reach-by-segment)")
    lines.append("  - [5.10 Address-level reach percentages](#510-address-level-reach-percentages)")
    lines.append("  - [5.11 Individual-level email reach](#511-individual-level-email-reach-linked-residents)")
    lines.append("  - [5.12 Homeowner DB users not subscribed to email blasts](#512-homeowner-db-users-not-subscribed-to-email-blasts)")
    lines.append("  - [5.13 ASCII Venn](#513-ascii-venn-printed-grit-vs-email-blasts)")
    lines.append("- [6. Interpretation notes](#6-interpretation-notes)")
    lines.append("")

    lines.append("## 1. Source files")
    lines.append(f"- Data snapshot folder: {resolved_data_dir}")
    lines.append(f"- Snapshot date: {snapshot_date}")
    lines.append(f"- Master addresses: {source_paths['addresses']}")
    lines.append(f"- Users export: {source_paths['users']}")
    lines.append(f"- Subscribers export: {source_paths['subscribers']}")
    lines.append("")

    lines.append("## 2. Definitions")
    lines.append("- **Address / household**: one row in the master addresses export.")
    lines.append(
        "- **User**: one record in the website users export; this includes anyone ever entered in the website database for any reason (for example former homeowners, committee members using a separate committee email address, and miscellaneous operational users such as the webmaster)."
    )
    lines.append("- **Homeowner (analysis definition)**: any user linked to a household; tenants are included.")
    lines.append(f"  - Tenant-indicated households in this snapshot: {fmt(tenant_households)}.")
    lines.append("- **Email blasts**: active Frequent Updates list subscribers.")
    lines.append("- **Essential-only**: Essential Updates Only subscribers who are not on email blasts.")
    lines.append("- **Invalid/placeholder email**: values like `fake.fake` or domains such as `@fake` are treated as no email (excluded from valid-email and subscriber-reach metrics).")
    lines.append("- **Reached**: household receives at least one of Printed GRIT (`Mail GRIT = 1`) or email blasts.")
    lines.append("- **Realtors list**: excluded from resident communication reach.")
    lines.append("")

    lines.append("## 3. Leadership answers")
    lines.append("### 3.1 Core leadership questions")
    lines.append(
        "This section summarizes household-level communication reach across SHHA channels (printed GRIT and email blasts) for member and non-member households."
    )
    lines.append("")
    lines.append(
        f"- Households reached at all: {fmt(household_overall['reached'])} out of {fmt(household_overall['total'])} ({pct(household_overall['reached'], household_overall['total'])})"
    )
    lines.append(
        f"- Households receiving no communication (neither GRIT nor email): {fmt(household_overall['neither'])} ({pct(household_overall['neither'], household_overall['total'])})"
    )
    lines.append("")
    lines.append(
        f"- Member households reached: {fmt(household_members['reached'])} out of {fmt(household_members['total'])} ({pct(household_members['reached'], household_members['total'])})"
    )
    lines.append(
        f"- Non-member households reached: {fmt(household_nonmembers['reached'])} out of {fmt(household_nonmembers['total'])} ({pct(household_nonmembers['reached'], household_nonmembers['total'])})"
    )
    lines.append("")
    lines.append("- Communication channel coverage (households):")
    lines.append(f"  - Households reached by both GRIT and email blasts: {fmt(household_overall['both'])}")
    lines.append(f"  - Households relying only on email blasts: {fmt(household_overall['only_email'])}")
    lines.append(f"  - Households relying only on GRIT: {fmt(household_overall['only_grit'])}")
    lines.append("")

    lines.append("### 3.2 User-level leadership snapshot")
    lines.append(
        f"- Member-linked users receiving email blasts: {fmt(resident_members['blast'])} out of {fmt(resident_members['total'])} ({pct(resident_members['blast'], resident_members['total'])})"
    )
    lines.append(
        f"- Non-member-linked users receiving email blasts: {fmt(resident_nonmembers['blast'])} out of {fmt(resident_nonmembers['total'])} ({pct(resident_nonmembers['blast'], resident_nonmembers['total'])})"
    )
    lines.append(
        f"- Member-linked users not receiving SHHA email: {fmt(resident_members['no_shha_email'])} out of {fmt(resident_members['total'])} ({pct(resident_members['no_shha_email'], resident_members['total'])})"
    )
    lines.append(
        f"- Non-member-linked users not receiving SHHA email: {fmt(resident_nonmembers['no_shha_email'])} out of {fmt(resident_nonmembers['total'])} ({pct(resident_nonmembers['no_shha_email'], resident_nonmembers['total'])})"
    )
    lines.append(
        f"- Users in homeowners database but not subscribed to email blasts: {inline_breakdown(homeowner_db_not_subscribed['total'], homeowner_db_not_subscribed['members'], homeowner_db_not_subscribed['nonmembers'])}"
    )
    lines.append(
        f"  - Of those not subscribed: never subscribed {fmt(homeowner_db_not_subscribed['never_subscribed'])}, unsubscribed {fmt(homeowner_db_not_subscribed['unsubscribed'])}, active on other SHHA lists but not email blasts {fmt(homeowner_db_not_subscribed['active_nonblast'])}, other status {fmt(homeowner_db_not_subscribed['other_status'])}"
    )
    lines.append("")

    lines.append("## 4. Trend since previous snapshot")
    lines.extend(trend_lines)
    lines.append("")

    lines.append("## 5. Technical deep dive")
    lines.append("### 5.1 Data consistency checks")
    lines.append(f"- Total households: {fmt(household_overall['total'])}")
    lines.append(f"- Households with linked users: {fmt(household_overall['users'])}")
    lines.append(f"- Households with no linked users: {fmt(len(no_linked_users))}")
    lines.append(f"- Top statuses among no-linked-user households: {no_linked_user_breakdown}")
    lines.append("")

    lines.append("### 5.2 Household Composition")
    lines.append(
        f"- Addresses with 0 users: {inline_breakdown(household_composition['zero_users'], members_no_users, nonmembers_no_users)}"
    )
    lines.append(
        f"- Addresses with 1 user: {inline_breakdown(household_composition['one_user'], int((member_households_df['linked_user_count'] == 1).sum()), int((nonmember_households_df['linked_user_count'] == 1).sum()))}"
    )
    lines.append(
        f"- Addresses with 2 users: {inline_breakdown(household_composition['two_users'], int((member_households_df['linked_user_count'] == 2).sum()), int((nonmember_households_df['linked_user_count'] == 2).sum()))}"
    )
    lines.append(
        f"- Addresses with 3+ users: {inline_breakdown(household_composition['three_plus_users'], int((member_households_df['linked_user_count'] >= 3).sum()), int((nonmember_households_df['linked_user_count'] >= 3).sum()))}"
    )
    lines.append("")
    lines.append(
        f"- Average users per address: {household_composition['avg_users_per_address']:.2f} (members: {member_households_df['linked_user_count'].mean():.2f}, non-members: {nonmember_households_df['linked_user_count'].mean():.2f})"
    )
    lines.append("")

    lines.append("### 5.3 Users Associated with Addresses")
    lines.append(f"- Users with 0 addresses: {fmt(users_associated_stats['zero_addresses'])}")
    lines.append(f"- Users with 1 address: {fmt(users_associated_stats['one_address'])}")
    lines.append(f"- Users with 2 addresses: {fmt(users_associated_stats['two_addresses'])}")
    lines.append(f"- Users with 3 addresses: {fmt(users_associated_stats['three_addresses'])}")
    lines.append(f"- Users with 4 addresses: {fmt(users_associated_stats['four_addresses'])}")
    lines.append(f"- Users with 5+ addresses: {fmt(users_associated_stats['five_plus_addresses'])}")
    lines.append("")

    lines.append("### 5.4 User Roles")
    lines.append(f"- owner: {fmt(user_roles['owner'])}")
    lines.append(f"- tenant: {fmt(user_roles['tenant'])}")
    lines.append(f"- other: {fmt(user_roles['other'])}")
    lines.append("")

    total_linked_users = len(linked_users)
    lines.append("### 5.5 User-level counts")
    lines.append(
        f"- Linked users (residents linked to addresses): {inline_breakdown(total_linked_users, resident_members['total'], resident_nonmembers['total'])}"
    )
    lines.append(
        f"- Linked users with valid email: {inline_breakdown(resident_overall['valid_email'], resident_members['valid_email'], resident_nonmembers['valid_email'])}"
    )
    lines.append(
        f"- Linked users receiving email blasts: {inline_breakdown(resident_overall['blast'], resident_members['blast'], resident_nonmembers['blast'])}"
    )
    denom_users = household_overall["users"] if household_overall["users"] else 0
    users_per_household = total_linked_users / denom_users if denom_users else 0
    lines.append(f"- Users per linked-user household: {users_per_household:.2f}")
    lines.append("")

    lines.append("### 5.6 Email list membership snapshot")
    lines.append("| Email list | Total active | Known members | Known non-members | Other |")
    lines.append("|---|---:|---:|---:|---:|")
    lines.append(
        f"| Email blasts (Frequent Updates) | {fmt(list_breakdowns['frequent']['total'])} | {fmt(list_breakdowns['frequent']['known_members'])} | {fmt(list_breakdowns['frequent']['known_nonmembers'])} | {fmt(list_breakdowns['frequent']['other'])} |"
    )
    lines.append(
        f"| Essential Updates Only | {fmt(list_breakdowns['essential_only']['total'])} | {fmt(list_breakdowns['essential_only']['known_members'])} | {fmt(list_breakdowns['essential_only']['known_nonmembers'])} | {fmt(list_breakdowns['essential_only']['other'])} |"
    )
    lines.append(
        f"| GRIT e-Delivery | {fmt(list_breakdowns['grit_edelivery']['total'])} | {fmt(list_breakdowns['grit_edelivery']['known_members'])} | {fmt(list_breakdowns['grit_edelivery']['known_nonmembers'])} | {fmt(list_breakdowns['grit_edelivery']['other'])} |"
    )
    lines.append("")

    lines.append("### 5.7 Household email coverage")
    lines.append(
        f"- Households with >=1 email blast subscriber: {inline_breakdown(household_overall['blast'], household_members['blast'], household_nonmembers['blast'])}"
    )
    lines.append(
        f"- Households with >=1 Essential Updates subscriber: {inline_breakdown(household_overall['essential'], household_members['essential'], household_nonmembers['essential'])}"
    )
    lines.append(
        f"- Households with no SHHA email subscribers: {inline_breakdown(household_overall['no_email_subscribers'], household_members['no_email_subscribers'], household_nonmembers['no_email_subscribers'])}"
    )
    lines.append(
        f"- Households with users but no valid email at all: {inline_breakdown(household_overall['users_no_valid_email'], members_users_no_valid_email, nonmembers_users_no_valid_email)}"
    )
    lines.append(
        f"- Households where multiple residents receive email blasts: {inline_breakdown(multi_email_households, members_multi_blast, nonmembers_multi_blast)}"
    )
    lines.append("")

    lines.append("### 5.8 GRIT coverage")
    lines.append(f"- Member households receiving printed GRIT: {format_count_pct(household_members['grit'], household_members['total'])}")
    lines.append(f"- Member households not receiving printed GRIT: {fmt(household_members['total'] - household_members['grit'])}")
    lines.append(f"- Non-members marked to receive GRIT (inconsistency): {fmt(household_nonmembers['grit'])}")
    lines.append("")

    lines.append("### 5.9 Communication reach by segment")
    lines.append(f"- Member households reached: {fmt(household_members['reached'])} / {fmt(household_members['total'])} ({members_reached_pct})")
    lines.append(f"- Non-member households reached: {fmt(household_nonmembers['reached'])} / {fmt(household_nonmembers['total'])} ({nonmembers_reached_pct})")
    lines.append(f"- Households receiving no SHHA communication: {fmt(household_overall['neither'])} (members: {fmt(household_members['neither'])}, non-members: {fmt(household_nonmembers['neither'])})")
    lines.append(
        f"- Households reached by at least one method: {inline_breakdown(household_overall['reached'], household_members['reached'], household_nonmembers['reached'])}"
    )
    lines.append("")

    lines.append("### 5.10 Address-level reach percentages")
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

    lines.append("### 5.11 Individual-level email reach (linked residents)")
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

    lines.append("### 5.12 Homeowner DB users not subscribed to email blasts")
    lines.append(
        f"- Users in homeowners database with valid email but not subscribed to email blasts: {inline_breakdown(homeowner_db_not_subscribed['total'], homeowner_db_not_subscribed['members'], homeowner_db_not_subscribed['nonmembers'])}"
    )
    lines.append(
        f"- Reason split: never subscribed {fmt(homeowner_db_not_subscribed['never_subscribed'])}, unsubscribed {fmt(homeowner_db_not_subscribed['unsubscribed'])}, active on other SHHA lists but not email blasts {fmt(homeowner_db_not_subscribed['active_nonblast'])}, other status {fmt(homeowner_db_not_subscribed['other_status'])}"
    )
    lines.append("")

    lines.append("### 5.13 ASCII Venn (Printed GRIT vs Email blasts)")
    lines.append("```")
    lines.append("Printed GRIT vs Email blasts")
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

    lines.append("## 6. Interpretation notes")
    lines.append("- 'Households with users but no email blasts subscriber' means no linked email at that address is actively subscribed to email blasts.")
    lines.append("- Such households may still have residents on Essential Updates Only; those are counted separately in household/individual metrics.")
    lines.append("- Resident-level metrics are derived from users linked to addresses via address parsing and matching rules.")

    output_path.write_text("\n".join(lines), encoding="utf-8")


def run_report(data_dir: Path, output_dir: Path, snapshot_date: Optional[str]) -> None:
    resolved_data_dir = resolve_data_dir(data_dir)
    effective_snapshot_date = infer_snapshot_date(snapshot_date, resolved_data_dir)

    addresses_path = resolved_data_dir / ADDRESS_FILE
    users_path = find_user_export_file(resolved_data_dir)
    subscribers_path = resolved_data_dir / SUBSCRIBER_FILE
    address_users_path = resolved_data_dir / ADDRESS_USERS_FILE

    if users_path is None:
        raise FileNotFoundError(f"No users export matching {USER_EXPORT_GLOB} in {resolved_data_dir}")

    addresses_df = pd.read_csv(addresses_path).fillna("")
    users_df = pd.read_csv(users_path).fillna("")
    subscribers_df = pd.read_csv(subscribers_path).fillna("")
    address_users_df = pd.read_csv(address_users_path).fillna("") if address_users_path.exists() else None

    subscriber_sets = build_subscriber_sets(subscribers_df)
    households = build_households(addresses_df, subscriber_sets)

    linked_users = link_users_to_households(users_df, households)
    linked_users = enrich_user_email_status(linked_users, subscriber_sets)

    email_membership_lookup = build_email_membership_lookup(linked_users)
    list_breakdowns = {
        "frequent": compute_list_breakdown(subscriber_sets["frequent"], email_membership_lookup),
        "essential_only": compute_list_breakdown(subscriber_sets["essential_only"], email_membership_lookup),
        "grit_edelivery": compute_list_breakdown(subscriber_sets["grit_edelivery"], email_membership_lookup),
    }
    homeowner_db_not_subscribed = compute_homeowner_db_not_subscribed_stats(linked_users, subscriber_sets)

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

    household_composition = compute_household_composition(households)
    users_associated_stats = compute_users_associated_stats(users_df)
    user_roles = compute_user_roles(address_users_df)

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
        household_composition=household_composition,
        users_associated_stats=users_associated_stats,
        user_roles=user_roles,
        list_breakdowns=list_breakdowns,
        homeowner_db_not_subscribed=homeowner_db_not_subscribed,
        trend_lines=trend_lines,
    )

    current_trend = trend_snapshot.copy()
    exec_latest_path = output_dir / "board_executive_summary_latest.md"
    write_executive_summary(
        output_path=exec_latest_path,
        snapshot_date=effective_snapshot_date,
        household_overall=household_overall,
        household_members=household_members,
        household_nonmembers=household_nonmembers,
        resident_overall=resident_overall,
        homeowner_db_not_subscribed=homeowner_db_not_subscribed,
        current_trend=current_trend,
        previous_snapshot=previous_snapshot,
    )

    exec_dated_path = output_dir / f"board_executive_summary_{effective_snapshot_date}.md"

    write_executive_summary(
        output_path=exec_dated_path,
        snapshot_date=effective_snapshot_date,
        household_overall=household_overall,
        household_members=household_members,
        household_nonmembers=household_nonmembers,
        resident_overall=resident_overall,
        homeowner_db_not_subscribed=homeowner_db_not_subscribed,
        current_trend=current_trend,
        previous_snapshot=previous_snapshot,
    )

    print(f"Report written: {report_path}")
    print(f"Executive summary written: {exec_latest_path}")
    print(f"Dated summary written: {exec_dated_path}")
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
