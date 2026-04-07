#!/usr/bin/env python3

import argparse
import html
import re
from pathlib import Path

import pandas as pd


def normalize_columns(cols):
    out = []
    for c in cols:
        c = "" if pd.isna(c) else str(c).strip()
        c = re.sub(r"\s+", " ", c)
        out.append(c)
    return out


def find_header_row(raw):
    """
    Find the Tableau header row automatically.
    Looks for a row containing likely field names.
    """
    for i in range(min(15, len(raw))):
        vals = ["" if pd.isna(x) else str(x).strip() for x in raw.iloc[i].tolist()]
        valset = set(vals)
        has_group = "Change Owner Group" in valset or "Assignment Group" in valset
        has_number = "Change#" in valset or "Number" in valset or "Change Number" in valset
        has_desc = "Short Description" in valset or "Short description" in valset
        if has_group and has_number and has_desc:
            return i
    raise ValueError("Could not locate header row in workbook.")


def parse_dt(series):
    """
    Parse mixed datetime strings from Tableau export.
    """
    s = series.astype(str).str.strip()
    return pd.to_datetime(s, errors="coerce")


def load_raw_changes(path):
    raw = pd.read_excel(path, header=None)
    header_row = find_header_row(raw)

    header = normalize_columns(raw.iloc[header_row].tolist())
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = header
    df.columns = normalize_columns(df.columns)

    # Drop fully blank columns
    df = df.loc[:, [c for c in df.columns if c != ""]].copy()

    # Normalize alternate names
    rename_map = {}
    if "Short description" in df.columns and "Short Description" not in df.columns:
        rename_map["Short description"] = "Short Description"
    if "Assignment group" in df.columns and "Assignment Group" not in df.columns:
        rename_map["Assignment group"] = "Assignment Group"
    df = df.rename(columns=rename_map)

    # Determine the change-number column
    change_col = None
    for candidate in ["Change#", "Number", "Change Number"]:
        if candidate in df.columns:
            change_col = candidate
            break
    if change_col is None:
        raise ValueError(
            f"Could not find a change number column. Found columns: {list(df.columns)}"
        )

    df["Change_ID"] = df[change_col]

    # Determine the grouping column
    if "Change Owner Group" not in df.columns:
        if "Assignment Group" in df.columns:
            df["Change Owner Group"] = df["Assignment Group"]
        else:
            df["Change Owner Group"] = ""

    # Forward-fill group names because Tableau exports often repeat them sparsely
    df["Change Owner Group"] = df["Change Owner Group"].ffill()

    # Keep only rows that represent a real change record
    df = df[df["Change_ID"].notna()].copy()

    # Ensure expected columns exist
    for col in [
        "Start Date",
        "End Date",
        "State",
        "Short Description",
        "Description",
        "Requested By",
        "Configuration Item",
        "Downtime",
        "Environments Impacted",
    ]:
        if col not in df.columns:
            df[col] = ""

    # Parse datetimes
    df["start_dt"] = parse_dt(df["Start Date"])
    df["end_dt"] = parse_dt(df["End Date"])

    return df


def classify_environment(env):
    if pd.isna(env) or str(env).strip() == "":
        return "missing"

    env_text = str(env).lower()

    prod_terms = ["production", "prod"]
    non_prod_terms = ["test", "qa", "dev", "uat", "non-prod", "non production", "sandbox"]

    has_prod = any(term in env_text for term in prod_terms)
    has_non_prod = any(term in env_text for term in non_prod_terms)

    if has_prod and has_non_prod:
        return "mixed"
    if has_prod:
        return "production"
    return "non_production"


def is_ucpath(row):
    text = " ".join(
        [
            str(row.get("Change Owner Group", "")),
            str(row.get("Short Description", "")),
            str(row.get("Configuration Item", "")),
            str(row.get("Description", "")),
        ]
    ).lower()

    return "ucpath" in text or "ucphrprd" in text


def fmt_dt(dt):
    if pd.isna(dt):
        return ""
    return dt.strftime("%-m/%-d/%y %-I:%M %p")

def fmt_window(start, end):
    if pd.isna(start) or pd.isna(end):
        return "Time window unavailable."

    def fmt_day_date(dt):
        dow = dt.strftime("%a") + "."
        date = dt.strftime("%-m/%-d")
        return f"{dow}, {date}"

    def fmt_time(dt):
        return dt.strftime("%-I:%M %p")

    def fmt_full(dt):
        return f"{fmt_day_date(dt)} {fmt_time(dt)}"

    if start.date() == end.date():
        return f"{fmt_full(start)} to {fmt_time(end)}"

    return f"{fmt_full(start)} to {fmt_full(end)}"

def fmt_downtime(val):
    s = "" if pd.isna(val) else str(val).strip()

    if not s:
        return "There will be no downtime."

    sl = s.lower()
    if sl in {"no", "none", "n/a", "na", "no downtime expected", "0", "0.0"}:
        return "There will be no downtime."

    try:
        minutes = int(float(s))
    except (ValueError, TypeError):
        return f"There will be {s} minutes of downtime."

    if minutes <= 0:
        return "There will be no downtime."

    if minutes < 60:
        return f"There will be {minutes} minute{'s' if minutes != 1 else ''} of downtime."

    hours = minutes // 60
    remaining = minutes % 60

    if remaining == 0:
        return f"There will be {hours} hour{'s' if hours != 1 else ''} of downtime."

    return (
        f"There will be {hours} hour{'s' if hours != 1 else ''} "
        f"and {remaining} minute{'s' if remaining != 1 else ''} of downtime."
    )


def fmt_requestor(val):
    s = "" if pd.isna(val) else str(val).strip()
    if not s:
        return "Please contact the requestor for any additional information."
    return f"Please contact {s} for additional information."

def pending_approval_suffix(state):
    state_text = "" if pd.isna(state) else str(state).lower()
    if "pending manager" in state_text or "pending approval" in state_text:
        return " {Pending Approval}"
    return ""


def build_item_html(row):
    change_id_raw = str(row["Change_ID"])
    change_id_clean = re.sub(r"\*+", "", change_id_raw)
    change_id = html.escape(change_id_clean)

    desc = html.escape(str(row["Short Description"]).strip())
    suffix = html.escape(pending_approval_suffix(row["State"]))

    window = html.escape(fmt_window(row["start_dt"], row["end_dt"]))
    downtime = html.escape(fmt_downtime(row.get("Downtime", "")))
    requestor = html.escape(fmt_requestor(row.get("Requested By", "")))

    return (
        "<table role='presentation' cellpadding='0' cellspacing='0' border='0' "
        "style='border-collapse:collapse; margin:0; padding:0; width:100%; "
        "font-family:Calibri, Arial, sans-serif; font-size:11pt; line-height:1.15;'>"
        "<tr>"
        "<td style='padding:0; margin:0; vertical-align:top;'>"
        f"&#8226; <b>{change_id}</b>: <b>{desc}</b>{suffix} &mdash; {window}. {downtime} {requestor}"
        "</td>"
        "</tr>"
        "<tr><td style='padding:2px 0 0 0; font-size:2px; line-height:2px;'>&nbsp;</td></tr>"
        "</table>"
    )


def build_html(prod):
    parts = [
        "<html>",
        "<head><meta charset='UTF-8'></head>",
        "<body style='font-family:Calibri, Arial, sans-serif; font-size:11pt; margin:0; padding:0; mso-line-height-rule:exactly;'>",
        "<div style='font-weight:bold; font-size:14pt; margin:0 0 8px 0;'>RFC Highlights</div>",
    ]

    non_uc = prod[~prod["ucpath"]].copy()
    if not non_uc.empty:
        for group_name, grp in non_uc.groupby("Change Owner Group", dropna=False):
            heading = str(group_name).strip() if str(group_name).strip() else "Other"
            parts.append(
                f"<div style='font-weight:bold; margin:8px 0 2px 0; "
                f"font-family:Calibri, Arial, sans-serif; font-size:11pt;'>{html.escape(heading)}</div>"
            )
            grp = grp.sort_values(["start_dt", "Change_ID"])
            for _, row in grp.iterrows():
                parts.append(build_item_html(row))

    uc = prod[prod["ucpath"]].copy()
    if not uc.empty:
        parts.append(
            "<div style='font-weight:bold; margin-top:14px; margin-bottom:6px;'>UCPath</div>"
        )
        uc = uc.sort_values(["start_dt", "Change_ID"])
        for _, row in uc.iterrows():
            parts.append(build_item_html(row))

    parts.append("</body></html>")
    return "\n".join(parts)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to Tableau-exported Excel file")
    parser.add_argument("--cab-date", required=True, help="CAB date, e.g. 2026-04-07")
    parser.add_argument("--output-html", default="weekly_change_email.html")
    parser.add_argument("--output-review", default="manual_review.txt")
    args = parser.parse_args()

    cab_date = pd.to_datetime(args.cab_date)
    df = load_raw_changes(args.input)

    df["env_class"] = df["Environments Impacted"].apply(classify_environment)
    df["ucpath"] = df.apply(is_ucpath, axis=1)

    # Rows to review manually: missing or mixed environment
    review = df[df["env_class"].isin(["missing", "mixed"])].copy()

    # Production-only rows starting within the CAB + 7 day window
    prod = df[df["env_class"] == "production"].copy()
    prod = prod[
        (prod["start_dt"].notna()) &
        (prod["start_dt"] <= cab_date + pd.Timedelta(days=7))
    ].copy()

    # Output files
    html_text = build_html(prod)
    html_content = build_html(prod)
    Path(args.output_html).write_text(html_content, encoding="utf-8")
    Path(args.output_review).write_text(review.to_string(index=False), encoding="utf-8")

    print("Done")


if __name__ == "__main__":
    main()
