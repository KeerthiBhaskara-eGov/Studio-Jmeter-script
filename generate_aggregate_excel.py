#!/usr/bin/env python3
"""
Builds a JMeter-style Aggregate Report from a raw .jtl results file
and saves it as an Excel (.xlsx) workbook, one row per API label.

Usage:
    python3 generate_aggregate_excel.py results.jtl aggregate_report.xlsx
"""
import sys
import pandas as pd
import numpy as np
from openpyxl.styles import Font


def percentile(series, p):
    return np.percentile(series, p, method="lower") if len(series) else 0


def summarize_group(label, group):
    elapsed = group["elapsed"].astype(float)
    n = len(group)
    errors = (~group["success"].astype(str).str.lower().eq("true")).sum()
    error_pct = round((errors / n) * 100, 2) if n else 0.0

    duration_sec = (group["timeStamp"].astype(float).max()
                     - group["timeStamp"].astype(float).min()) / 1000.0
    if duration_sec <= 0:
        duration_sec = elapsed.sum() / 1000.0 or 1

    throughput = round(n / duration_sec, 2) if duration_sec else 0.0
    kb_received = round((group["bytes"].astype(float).sum() / 1024) / duration_sec, 2) if "bytes" in group else 0.0
    kb_sent = round((group["sentBytes"].astype(float).sum() / 1024) / duration_sec, 2) if "sentBytes" in group else 0.0

    return {
        "API": label,
        "Samples hits": n,
        "Error %": f"{error_pct:.2f}%",
        "Average response time in milli sec": round(elapsed.mean(), 2),
        "Median": round(percentile(elapsed, 50), 2),
        "90% line": round(percentile(elapsed, 90), 2),
        "95% line": round(percentile(elapsed, 95), 2),
        "99% line": round(percentile(elapsed, 99), 2),
        "Min": round(elapsed.min(), 2),
        "Max": round(elapsed.max(), 2),
        "Throughput (per sec)": throughput,
        "Received KB/sec": kb_received,
        "Sent KB/sec": kb_sent,
    }


def build_aggregate(df: pd.DataFrame) -> pd.DataFrame:
    rows = [summarize_group(label, group) for label, group in df.groupby("label", sort=False)]
    rows.append(summarize_group("TOTAL", df))

    cols = ["API", "Samples hits", "Error %", "Average response time in milli sec",
            "Median", "90% line", "95% line", "99% line", "Min", "Max",
            "Throughput (per sec)", "Received KB/sec", "Sent KB/sec"]
    return pd.DataFrame(rows, columns=cols)


def main():
    if len(sys.argv) < 3:
        print("Usage: python3 generate_aggregate_excel.py <results.jtl> <output.xlsx>")
        sys.exit(1)

    jtl_path, xlsx_path = sys.argv[1], sys.argv[2]

    df = pd.read_csv(jtl_path)
    df.columns = [c.strip() for c in df.columns]

    required = {"timeStamp", "elapsed", "label", "success", "bytes"}
    missing = required - set(df.columns)
    if missing:
        print(f"Warning: missing expected columns in jtl: {missing}. "
              f"Make sure the jtl was saved in CSV format with standard fields.")

    aggregate = build_aggregate(df)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        aggregate.to_excel(writer, sheet_name="Aggregate Report", index=False)
        ws = writer.sheets["Aggregate Report"]
        for i, col in enumerate(aggregate.columns, start=1):
            width = max(12, min(35, aggregate[col].astype(str).map(len).max() + 2, len(col) + 4))
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
        for cell in ws[1]:
            cell.font = Font(bold=True)

    print(f"Aggregate report written to {xlsx_path}")
    print(f"APIs found: {df['label'].nunique()} -> {list(df['label'].unique())}")
    total_errors = (~df["success"].astype(str).str.lower().eq("true")).sum() if "success" in df else 0
    print(f"Total samples: {len(df)} | Errors: {total_errors} "
          f"| Error rate: {round(total_errors/len(df)*100, 2) if len(df) else 0}%")


if __name__ == "__main__":
    main()
