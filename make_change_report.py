#!/usr/bin/env python3

import argparse
import html
import re
from pathlib import Path

import pandas as pd

SHAREPOINT_BASE = "https://ucofficeofthepresident.sharepoint.com/sites/its/cm/Shared%20Documents"

def load_footer_rtf(path):
    if not path:
        return ""

    text = Path(path).read_text(encoding="utf-8", errors="replace")

    start = text.find(r"\pard")
    end = text.rfind("}")

    if start == -1 or end == -1 or end <= start:
        return ""

    return text[start:end].strip()

def sharepoint_pdf_url(filename):
    return f"{SHAREPOINT_BASE}/{filename.replace(' ', '%20')}"

def fmt_report_date(dt):
    return dt.strftime("%m-%d-%y")

def report_links(cab_date):
    cab_date = pd.to_datetime(cab_date)

    post_cab_date = fmt_report_date(cab_date)

    range_end = cab_date - pd.Timedelta(days=3)
    range_start = range_end - pd.Timedelta(days=6)

    range_text = f"{fmt_report_date(range_start)} to {fmt_report_date(range_end)}"

    return {
        "post_cab_label": f"{post_cab_date} Post CAB Report",
        "post_cab_url": sharepoint_pdf_url(f"{post_cab_date} Post CAB Report.pdf"),

        "pir_label": "PIR Weekly Report",
        "pir_url": sharepoint_pdf_url(f"Weekly PIR Report {range_text}.pdf"),

        "pir_standard_label": "PIR Weekly Report - Standard Changes",
        "pir_standard_url": sharepoint_pdf_url(f"Weekly PIR Report - Standard Changes {range_text}.pdf"),

        "rca_label": "Weekly Problem RCAs",
        "rca_url": sharepoint_pdf_url(f"Weekly Problem RCAs {range_text}.pdf"),
    }

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
    return pd.to_datetime(
    s,
    format="%m/%d/%y  %I:%M %p",
    errors="coerce"
)


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


def has_ucpath_impact(text):
    text = str(text or "").lower()

    negative_phrases = [
        "no impact to ucpath",
        "no ucpath impact",
        "no impact on ucpath",
        "does not impact ucpath",
        "will not impact ucpath",
        "no uc path impact",
        "no impact to uc path",
        "no impact to ucphr",
        "no impact to ucphrprd",
    ]

    if any(phrase in text for phrase in negative_phrases):
        return False

    return "ucpath" in text or "uc path" in text or "ucphrprd" in text


def is_ucpath(row):
    text = " ".join(
        [
            str(row.get("Change Owner Group", "")),
            str(row.get("Short Description", "")),
            str(row.get("Configuration Item", "")),
            str(row.get("Description", "")),
        ]
    )

    return has_ucpath_impact(text)

def is_ucpath_maintenance_window(row):
    if not row.get("ucpath", False):
        return False

    start = row.get("start_dt")
    end = row.get("end_dt")

    if pd.isna(start) or pd.isna(end):
        return False

    # Sunday is 6 in Python's weekday() where Monday=0
    if start.weekday() != 6:
        return False

    start_minutes = start.hour * 60 + start.minute
    end_minutes = end.hour * 60 + end.minute

    # Must fall entirely within Sunday 12:00 AM to 6:00 AM
    return start_minutes >= 0 and end_minutes <= 360 and end.date() == start.date()

def fmt_ucpath_maintenance_heading(dt):
    return dt.strftime("UCPath Production Maintenance Window (Sunday, %m/%d/%y, 12:00 AM to 6:00 AM)")

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
        return "There is no downtime."

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

def fmt_requestor_more_info(val):
    s = "" if pd.isna(val) else str(val).strip()
    if not s:
        return "Please contact the requestor for more information."
    return f"Please contact {s} for more information."

"""
**********************************************
**************** Build HTML ******************
**********************************************
"""

def build_ucpath_maintenance_item_html(row):
    change_id_raw = str(row["Change_ID"])
    change_id_clean = re.sub(r"\*+", "", change_id_raw)
    change_id = html.escape(change_id_clean)

    desc = html.escape(str(row["Short Description"]).strip())
    suffix = html.escape(pending_approval_suffix(row["State"]))
    requestor = html.escape(fmt_requestor_more_info(row.get("Requested By", "")))

    return (
        "<div style='margin:0 0 8px 28px; font-family:Calibri, Arial, sans-serif; "
        "font-size:11pt; line-height:1.25;'>"
        f"<span style='margin-left:-14px;'>&#8226;</span> "
        f"<b>{change_id}: {desc}</b>{suffix} &mdash; {requestor}"
        "</div>"
    )

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
        "<div style='margin:0 0 8px 28px; font-family:Calibri, Arial, sans-serif; "
        "font-size:11pt; line-height:1.25;'>"
        f"<span style='margin-left:-14px;'>&#8226;</span> "
        f"<b>{change_id}: {desc}</b>{suffix} &mdash; {window}. "
        f"{downtime} {requestor}"
        "</div>"
    )

def build_report_links_html(cab_date):
    links = report_links(cab_date)

    return (
        "<div style='font-family:Calibri, Arial, sans-serif; font-size:11pt; margin-top:12px;'>"
        "The "
        f'<a href="{html.escape(links["post_cab_url"])}">{html.escape(links["post_cab_label"])}</a>, '
        f'<a href="{html.escape(links["pir_url"])}">{html.escape(links["pir_label"])}</a>, '
        f'<a href="{html.escape(links["pir_standard_url"])}">{html.escape(links["pir_standard_label"])}</a>, '
        f'and <a href="{html.escape(links["rca_url"])}">{html.escape(links["rca_label"])}</a> '
        "are available in our Change Management SharePoint site."
        "</div>"
    )

def build_html(prod, cab_date):
    parts = [
        "<html>",
        "<head><meta charset='UTF-8'></head>",
        "<body style='font-family:Calibri, Arial, sans-serif; font-size:11pt; "
        "margin:0; padding:0; mso-line-height-rule:exactly;'>",
        "<div style='font-weight:bold; font-size:14pt; margin:0 0 18px 0;'>RFC Highlights:</div>",
    ]

    non_uc = prod[~prod["ucpath"]].copy()
    if not non_uc.empty:
        for group_name, grp in non_uc.groupby("Change Owner Group", dropna=False):
            heading = str(group_name).strip() if str(group_name).strip() else "Other"
            parts.append(
                "<div style='font-weight:bold; font-size:12pt; "
                "margin:0 0 8px 14px; font-family:Calibri, Arial, sans-serif;'>"
                f"{html.escape(heading)}"
                "</div>"
            )
            grp = grp.sort_values(["start_dt", "Change_ID"])
            for _, row in grp.iterrows():
                parts.append(build_item_html(row))

            parts.append("<div style='height:14px; line-height:14px;'>&nbsp;</div>")

    uc = prod[prod["ucpath"]].copy()
    if not uc.empty:
        uc = uc.copy()
        uc["ucpath_maint"] = uc.apply(is_ucpath_maintenance_window, axis=1)

        uc_maint = uc[uc["ucpath_maint"]].copy()
        uc_other = uc[~uc["ucpath_maint"]].copy()

        if not uc_maint.empty:
            first_dt = uc_maint.sort_values(["start_dt"]).iloc[0]["start_dt"]
            parts.append(
                "<div style='font-weight:bold; font-size:12pt; "
                "margin:0 0 8px 14px; font-family:Calibri, Arial, sans-serif;'>"
                f"{html.escape(fmt_ucpath_maintenance_heading(first_dt))}"
                "</div>"
            )
            uc_maint = uc_maint.sort_values(["start_dt", "Change_ID"])
            for _, row in uc_maint.iterrows():
                parts.append(build_ucpath_maintenance_item_html(row))

            parts.append("<div style='height:14px; line-height:14px;'>&nbsp;</div>")

        if not uc_other.empty:
            parts.append(
                "<div style='font-weight:bold; font-size:12pt; "
                "margin:0 0 8px 14px; font-family:Calibri, Arial, sans-serif;'>UCPath</div>"
            )
            uc_other = uc_other.sort_values(["start_dt", "Change_ID"])
            for _, row in uc_other.iterrows():
                parts.append(build_item_html(row))
    parts.append(build_report_links_html(cab_date))           
    parts.append("</body></html>")
    return "\n".join(parts)

"""
**********************************************
******** Build Plain Text Section ************
**********************************************
"""

def build_item_text(row):
    change_id = re.sub(r"\*+", "", str(row["Change_ID"])).strip()
    desc = str(row["Short Description"]).strip()
    suffix = pending_approval_suffix(row["State"])

    window = fmt_window(row["start_dt"], row["end_dt"])
    downtime = fmt_downtime(row.get("Downtime", ""))
    requestor = fmt_requestor(row.get("Requested By", ""))

    return f"  • {change_id}: {desc}{suffix} — {window}. {downtime} {requestor}"

def build_ucpath_maintenance_item_text(row):
    change_id = re.sub(r"\*+", "", str(row["Change_ID"])).strip()
    desc = str(row["Short Description"]).strip()
    suffix = pending_approval_suffix(row["State"])
    requestor = fmt_requestor_more_info(row.get("Requested By", ""))

    return f"  • {change_id}: {desc}{suffix} — {requestor}"

def build_report_links_text(cab_date):
    links = report_links(cab_date)

    return (
        f'{links["post_cab_label"]}, {links["pir_label"]}, '
        f'{links["pir_standard_label"]}, and {links["rca_label"]} are available '
        f'in our Change Management SharePoint site.\n\n'
        f'{links["post_cab_label"]}: {links["post_cab_url"]}\n'
        f'{links["pir_label"]}: {links["pir_url"]}\n'
        f'{links["pir_standard_label"]}: {links["pir_standard_url"]}\n'
        f'{links["rca_label"]}: {links["rca_url"]}'
    )

def build_text(prod, cab_date):
    lines = []
    lines.append("RFC Highlights:")
    lines.append("")

    non_uc = prod[~prod["ucpath"]].copy()
    if not non_uc.empty:
        for group_name, grp in non_uc.groupby("Change Owner Group", dropna=False):
            heading = str(group_name).strip() if str(group_name).strip() else "Other"
            lines.append(heading)
            grp = grp.sort_values(["start_dt", "Change_ID"])
            for _, row in grp.iterrows():
                lines.append(build_item_text(row))
            lines.append("")

    uc = prod[prod["ucpath"]].copy()
    if not uc.empty:
        uc = uc.copy()
        uc["ucpath_maint"] = uc.apply(is_ucpath_maintenance_window, axis=1)

        uc_maint = uc[uc["ucpath_maint"]].copy()
        uc_other = uc[~uc["ucpath_maint"]].copy()

        if not uc_maint.empty:
            first_dt = uc_maint.sort_values(["start_dt"]).iloc[0]["start_dt"]
            lines.append(fmt_ucpath_maintenance_heading(first_dt))
            uc_maint = uc_maint.sort_values(["start_dt", "Change_ID"])
            for _, row in uc_maint.iterrows():
                lines.append(build_ucpath_maintenance_item_text(row))
            lines.append("")

        if not uc_other.empty:
            lines.append("UCPath")
            uc_other = uc_other.sort_values(["start_dt", "Change_ID"])
            for _, row in uc_other.iterrows():
                lines.append(build_item_text(row))
            lines.append("")
            
    lines.append(build_report_links_text(cab_date)) 
    return "\n".join(lines).rstrip() + "\n"

"""
**********************************************
************ Build RTF Section ***************
**********************************************
"""

def rtf_escape(s):
    s = "" if pd.isna(s) else str(s)
    return (
        s.replace("\\", r"\\")
         .replace("{", r"\{")
         .replace("}", r"\}")
         .replace("—", r"\emdash ")
    )

def rtf_par(text="", bold=False, size=22, indent=False):
    # RTF font sizes are half-points: 22 = 11pt, 24 = 12pt, 28 = 14pt
    left_indent = r"\li360\fi-180 " if indent else ""
    b1 = r"\b " if bold else ""
    b0 = r"\b0 " if bold else ""
    return rf"\pard {left_indent}\fs{size} {b1}{rtf_escape(text)}{b0}\par"


def build_item_rtf(row):
    change_id = re.sub(r"\*+", "", str(row["Change_ID"])).strip()
    desc = str(row["Short Description"]).strip()
    suffix = pending_approval_suffix(row["State"])

    window = fmt_window(row["start_dt"], row["end_dt"])
    downtime = fmt_downtime(row.get("Downtime", ""))
    requestor = fmt_requestor(row.get("Requested By", ""))

    main = f"{change_id}: {desc}"
    rest = f"{suffix} — {window}. {downtime} {requestor}"

    return (
        r"\pard \fs22 "
        r"\b " + rtf_escape(main) + r"\b0 "
        + rtf_escape(rest) +
        r"\par"
    )


def build_ucpath_maintenance_item_rtf(row):
    change_id = re.sub(r"\*+", "", str(row["Change_ID"])).strip()
    desc = str(row["Short Description"]).strip()
    suffix = pending_approval_suffix(row["State"])
    requestor = fmt_requestor_more_info(row.get("Requested By", ""))

    main = f"{change_id}: {desc}"
    rest = f"{suffix} — {requestor}"

    return (
        r"\pard \fs22 "
        r"\b " + rtf_escape(main) + r"\b0 "
        + rtf_escape(rest) +
        r"\par"
    )

def rtf_highlight_par(text="", bold=True, size=24, highlight_color=1):
    b1 = r"\b " if bold else ""
    b0 = r"\b0 " if bold else ""
    return (
        rf"\pard \fs{size} "
        rf"\highlight{highlight_color} "
        f"{b1}{rtf_escape(text)}{b0}"
        r"\highlight0 \par"
    )

def rtf_hyperlink(label, url):
    return (
        r'{\field{\*\fldinst HYPERLINK "' + rtf_escape(url) + r'"}'
        r'{\fldrslt ' + rtf_escape(label) + r'}}'
    )

def build_report_links_rtf(cab_date):
    links = report_links(cab_date)

    return (
        r"\pard \fs22 "
        + rtf_escape("The ")
        + rtf_hyperlink(links["post_cab_label"], links["post_cab_url"])
        + rtf_escape(", ")
        + rtf_hyperlink(links["pir_label"], links["pir_url"])
        + rtf_escape(", ")
        + rtf_hyperlink(links["pir_standard_label"], links["pir_standard_url"])
        + rtf_escape(", and ")
        + rtf_hyperlink(links["rca_label"], links["rca_url"])
        + rtf_escape(" are available in our Change Management SharePoint site.")
        + r"\par"
    )

def build_ucpath_impact_note_rtf():
    return (
        r"\pard \fs22 "
        r"\i "  # italic on
        r"\highlight1 " + rtf_escape("*UCPath Impact") + r"\highlight0 "  # yellow highlight
        + rtf_escape(" — Highlighted items denote changes related to/ or may have potential impact to UCPath")
        + r"\i0 "  # italic off
        r"\par"
    )

def build_rtf(prod, cab_date, footer_rtf_path=None):
    parts = [
        r"{\rtf1\ansi\deff0",
        r"{\fonttbl{\f0 Calibri;}}",
        r"{\colortbl ;\red255\green255\blue0;}",
        r"\f0",
        rtf_par("RFC Highlights:", bold=True, size=28),
        rtf_par("")
    ]

    prod = prod.copy()
    prod["ucpath_maint"] = prod.apply(is_ucpath_maintenance_window, axis=1)

    sections = []

    # Normal owner-group sections.
    # This includes UCPath-related changes that are NOT in the Sunday maintenance window.
    normal = prod[~prod["ucpath_maint"]].copy()

    if not normal.empty:
        for group_name, grp in normal.groupby("Change Owner Group", dropna=False):
            heading = str(group_name).strip() if str(group_name).strip() else "Other"

            highlight = grp["ucpath"].any()

            items = [
                build_item_rtf(row)
                for _, row in grp.sort_values(["start_dt", "Change_ID"]).iterrows()
            ]

            sections.append((heading.lower(), heading, highlight, items))

    # UCPath Sunday maintenance window section.
    uc_maint = prod[prod["ucpath_maint"]].copy()

    if not uc_maint.empty:
        first_dt = uc_maint.sort_values(["start_dt"]).iloc[0]["start_dt"]
        heading = fmt_ucpath_maintenance_heading(first_dt)

        items = [
            build_ucpath_maintenance_item_rtf(row)
            for _, row in uc_maint.sort_values(["start_dt", "Change_ID"]).iterrows()
        ]

        sections.append((heading.lower(), heading, True, items))

    # Render all sections alphabetically by heading.
    for _, heading, highlight, items in sorted(sections, key=lambda x: x[0]):
        if highlight:
            parts.append(rtf_highlight_par(heading, bold=True, size=24))
        else:
            parts.append(rtf_par(heading, bold=True, size=24))

        parts.extend(items)
        parts.append(rtf_par(""))
    parts.append(rtf_par(""))
    parts.append(build_ucpath_impact_note_rtf())
    parts.append(rtf_par(""))
    parts.append(build_report_links_rtf(cab_date))
    parts.append(rtf_par(""))
    footer_rtf = load_footer_rtf(footer_rtf_path)
    #debgging code below
    #print(footer_rtf[:500])

    if footer_rtf:
        parts.append(rtf_par(""))
        parts.append(footer_rtf)
    
    parts.append("}")
    print("footer_rtf_path:", footer_rtf_path)
    print("footer_rtf length:", len(footer_rtf) if footer_rtf else 0)
    return "\n".join(parts)


"""
**********************************************
*************** Build Report *****************
**********************************************
"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to Tableau-exported Excel file")
    parser.add_argument("--cab-date", required=True, help="CAB date, e.g. 2026-04-07")
    parser.add_argument("--output-html", default="weekly_change_email.html")
    parser.add_argument("--output-review", default="manual_review.txt")
    parser.add_argument("--output-text", default=None, help="Optional plain-text output file")
    parser.add_argument("--output-rtf", default=None, help="Optional RTF output file")
    parser.add_argument("--footer-rtf", default=None)
    args = parser.parse_args()

    cab_date = pd.to_datetime(args.cab_date)
    df = load_raw_changes(args.input)
    print(df[["Change_ID", "Start Date", "End Date", "start_dt"]].head(20))
    print(df.columns.tolist())

    df["env_class"] = df["Environments Impacted"].apply(classify_environment)
    df["ucpath"] = df.apply(is_ucpath, axis=1)

    # Rows to review manually: missing or mixed environment
    review = df[df["env_class"].isin(["missing", "mixed"])].copy()

    # Include all production changes through the end of next Tuesday
    next_tuesday = cab_date + pd.Timedelta(days=7)
    report_cutoff = next_tuesday.normalize() + pd.Timedelta(days=1)

    prod = df[df["env_class"] == "production"].copy()
    prod = prod[
        (prod["start_dt"].notna()) &
        (prod["start_dt"] < report_cutoff)
    ].copy()

    print("df rows:", len(df))
    print("prod rows:", len(prod))
    print(df[["Change_ID", "Short Description", "Environments Impacted", "env_class", "start_dt"]].head(20))

    # Output files
    html_content = build_html(prod, cab_date)
    if args.output_html:
        Path(args.output_html).write_text(html_content, encoding="utf-8")

    if args.output_text:
        text_content = build_text(prod, cab_date)
        Path(args.output_text).write_text(text_content, encoding="utf-8")

    if args.output_review:
        Path(args.output_review).write_text(review.to_string(index=False), encoding="utf-8")

    if args.output_rtf:
        rtf_content = build_rtf(prod, cab_date, args.footer_rtf)
        Path(args.output_rtf).write_text(rtf_content, encoding="utf-8")
        
    print("Done")

if __name__ == "__main__":
    main()
