#!/usr/bin/env python3
import pandas as pd
import argparse
from datetime import datetime
import os

def normalize_column(df, target_name):
    """
    Find a column in df that loosely matches target_name.
    """
    target = target_name.lower().replace(" ", "").replace("/", "")
    for col in df.columns:
        normalized = col.lower().replace(" ", "").replace("/", "")
        if target in normalized:
            return col
    raise ValueError(f"Could not find column matching '{target_name}'")

def main():
    parser = argparse.ArgumentParser(description="Find past-due changes not implemented.")
    parser.add_argument("input_file", help="Path to Excel file")
    parser.add_argument("-o", "--output", help="Output text file", default="past_due_changes.txt")
    args = parser.parse_args()

    # Load Excel
    df = pd.read_excel(args.input_file)

    # Normalize column names
    number_col = normalize_column(df, "Number")
    title_col = normalize_column(df, "Title/ Short Description")
    state_col = normalize_column(df, "Change Request State")
    date_col = normalize_column(df, "Planned end date")

    # Convert dates safely
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    today = pd.Timestamp(datetime.now().date())

    # Filter logic
    filtered = df[
        (df[date_col] < today) &  # past due
        (~df[state_col].str.lower().eq("implemented")) &  # not implemented
        (~df[title_col].str.startswith("Run Attached SQL", na=False))  # exclude SQL jobs
    ]

    # Check for Pending CAB
    pending_cab = filtered[
        filtered[state_col].str.startswith("Pending CAB", na=False)
    ]

    if not pending_cab.empty:
        print("This/These changes need to be approved")

    # Write output file
    lines = [
        f"{row[number_col]} - {row[title_col]}"
        for _, row in filtered.iterrows()
    ]

    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n\n".join(lines))

    print(f"Found {len(filtered)} matching changes.")
    print(f"Output written to: {os.path.abspath(args.output)}")

if __name__ == "__main__":
    main()