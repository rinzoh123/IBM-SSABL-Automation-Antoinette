"""
SSABL new-signings report automation.

Replaces the manual pivot-table steps:
  raw SSABL export (.xlsx)  ->  filtered, seller-assigned report (.xlsx)  ->  Slack post

Usage:
  python ssabl_automate.py --raw SSABL_export.xlsx --seller-map SellerMap.xlsx [--out report.xlsx]
                           [--slack-webhook https://hooks.slack.com/...]   (or env SLACK_WEBHOOK_URL)

The seller map is the 'SellerMap' sheet (Coverage Type ID | Coverage Type Name |
AI productivity | Data Security and Planning | TSS). It can be its own workbook or
any workbook that contains a sheet named SellerMap.
"""
import argparse, datetime as dt, json, os, sys, urllib.request
import pandas as pd

# ---- Report scope (the "Filters" from the instructions) -------------------
FILTERS = {
    "Country": "United States",
    "CLIENT SUB TYPE": "Select T Activate",
    "Signing Category": "New Signings",
    "UT LEVEL CODE 15 DSCR": "Data Platform",
}
# ---- Product -> category -> which SellerMap column owns it ----------------
CATEGORY_OF_PRODUCT = {
    "Bob": "AI", "watsonx Orchestrate": "AI", "watsonx.ai": "AI",
    "Planning Analytics": "Data", "watsonx.data": "Data", "watsonx.data integration": "Data",
    "Cloud Pak for Business Automation": "TSS", "Cloud Pak for Data": "TSS",
    "Cognos Analytics": "TSS", "FileNet Content Manager": "TSS",
    "Process Automation": "TSS", "SPSS Statistics": "TSS",
}
SELLER_COLUMN = {"AI": "AI productivity", "Data": "Data Security and Planning", "TSS": "TSS"}

COLS = ["UT LEVEL CODE 30 DSCR", "Coverage Id", "SIGNING DATE", "SALES ORDER NUM",
        "Customer Name Full Mrgd", "SAAS ACV USD AT PLAN"]


def report_window(today: dt.date):
    """Start = first day of the fiscal quarter; during the first 5 days of a new
    quarter keep using the previous quarter's start (not enough data yet)."""
    q_start = dt.date(today.year, 3 * ((today.month - 1) // 3) + 1, 1)
    if (today - q_start).days < 5:
        prev = q_start - dt.timedelta(days=1)
        q_start = dt.date(prev.year, 3 * ((prev.month - 1) // 3) + 1, 1)
    return q_start, today


def find_data_sheet(path):
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(path)
    for name in wb.sheet_names:
        hdr = wb.get_sheet_by_name(name).to_python(nrows=1)
        if hdr and "SALES ORDER NUM" in hdr[0] and "UT LEVEL CODE 15 DSCR" in hdr[0]:
            return name
    sys.exit("No sheet with SSABL column headers found in " + path)


def load_raw(path):
    sheet = find_data_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet, engine="calamine",
                       usecols=list(FILTERS) + COLS,
                       dtype={"Coverage Id": str, "SALES ORDER NUM": str})
    df["SIGNING DATE"] = pd.to_datetime(df["SIGNING DATE"]).dt.date
    df["Coverage Id"] = df["Coverage Id"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(7)
    return df


def load_seller_map(path):
    xl = pd.ExcelFile(path, engine="calamine")
    sheet = "SellerMap" if "SellerMap" in xl.sheet_names else xl.sheet_names[0]
    sm = pd.read_excel(xl, sheet_name=sheet, dtype=str).fillna("")
    sm["Coverage Type ID"] = sm["Coverage Type ID"].str.strip()
    return sm.set_index("Coverage Type ID")


def build(df, sm, start, end):
    m = df
    for col, val in FILTERS.items():
        m = m[m[col] == val]
    m = m[(m["SIGNING DATE"] >= start) & (m["SIGNING DATE"] <= end)].copy()

    m["Category"] = m["UT LEVEL CODE 30 DSCR"].map(CATEGORY_OF_PRODUCT).fillna("Unmapped")
    m["Coverage Key"] = "T" + m["Coverage Id"]
    m["Territory"] = m["Coverage Key"].map(sm["Coverage Type Name"]).fillna("")

    def seller(r):
        if r["Coverage Key"] not in sm.index:
            return "Unassigned Coverage ID"
        if r["Category"] not in SELLER_COLUMN:
            return "Unmapped product"
        name = sm.at[r["Coverage Key"], SELLER_COLUMN[r["Category"]]].strip()
        return name or "Vacant seat"
    m["Seller"] = m.apply(seller, axis=1)

    detail = (m.groupby(["Category", "UT LEVEL CODE 30 DSCR", "Seller", "Coverage Id", "Territory",
                         "SIGNING DATE", "SALES ORDER NUM", "Customer Name Full Mrgd"], as_index=False)
                ["SAAS ACV USD AT PLAN"].sum()
                .rename(columns={"UT LEVEL CODE 30 DSCR": "Product", "Customer Name Full Mrgd": "Customer",
                                 "SALES ORDER NUM": "Sales Order", "SIGNING DATE": "Signing Date",
                                 "SAAS ACV USD AT PLAN": "ACV (USD at plan)"})
                .sort_values(["Category", "Product", "Seller", "Signing Date"]))
    return m, detail


def write_excel(detail, out, start, end, raw_total):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    money = '"$"#,##0;[Red]-"$"#,##0'
    by_seller = (detail.groupby(["Seller", "Category"], as_index=False)
                 .agg(Deals=("Sales Order", "nunique"), Customers=("Customer", "nunique"),
                      ACV=("ACV (USD at plan)", "sum"))
                 .sort_values("ACV", ascending=False))
    by_product = (detail.groupby(["Category", "Product"], as_index=False)
                  .agg(Deals=("Sales Order", "nunique"), ACV=("ACV (USD at plan)", "sum"))
                  .sort_values(["Category", "ACV"], ascending=[True, False]))
    issues = detail[detail["Seller"].isin(["Unassigned Coverage ID", "Vacant seat", "Unmapped product"])]
    issues = (issues.groupby(["Seller", "Coverage Id", "Territory", "Product"], as_index=False)
              .agg(Customers=("Customer", lambda s: ", ".join(sorted(set(s)))),
                   Lines=("Sales Order", "count"), ACV=("ACV (USD at plan)", "sum"))
              .rename(columns={"Seller": "Issue"}))

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        meta = pd.DataFrame({"": ["Window", "Scope", "Total ACV", "Check vs raw filter"],
                             "Value": [f"{start:%b %d, %Y} - {end:%b %d, %Y}",
                                       " | ".join(f"{k}: {v}" for k, v in FILTERS.items()),
                                       detail["ACV (USD at plan)"].sum(),
                                       "OK" if abs(detail["ACV (USD at plan)"].sum() - raw_total) < 0.01 else "MISMATCH"]})
        meta.to_excel(xw, "Summary", index=False)
        by_product.to_excel(xw, "Summary", index=False, startrow=6)
        by_seller.to_excel(xw, "By Seller", index=False)
        detail.to_excel(xw, "Detail", index=False)
        issues.to_excel(xw, "Needs Attention", index=False)
        for ws in xw.book.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.row in (1, 7) and ws.title == "Summary" or (c.row == 1 and ws.title != "Summary"):
                        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
            hdr = {c.value: c.column for c in ws[7 if ws.title == "Summary" else 1]}
            for name in ("ACV", "ACV (USD at plan)"):
                if name in hdr:
                    for r in ws.iter_rows(min_row=2, min_col=hdr[name], max_col=hdr[name]):
                        r[0].number_format = money
            if ws.title == "Summary":
                ws["B3"].number_format = money
            if "Signing Date" in hdr:
                for r in ws.iter_rows(min_row=2, min_col=hdr["Signing Date"], max_col=hdr["Signing Date"]):
                    r[0].number_format = "mmm d, yyyy"
            for i, col in enumerate(ws.columns, 1):
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, w + 2), 60)
            if ws.title != "Summary":
                ws.auto_filter.ref = ws.dimensions; ws.freeze_panes = "A2"
    return by_seller, issues


def slack_text(detail, issues, start, end, top_n=3):
    fmt = lambda v: f"${v:,.0f}"
    total = detail["ACV (USD at plan)"].sum()
    lines = [f"*SSABL New Signings — Data Platform, US Select T Activate*",
             f"_{start:%b %d} – {end:%b %d, %Y}_  •  Total ACV *{fmt(total)}*  •  "
             f"{detail['Sales Order'].nunique()} orders  •  {detail['Customer'].nunique()} customers", ""]
    for cat, g in detail.groupby("Category"):
        lines.append(f"*{cat}* — {fmt(g['ACV (USD at plan)'].sum())}")
        for seller, s in g.groupby("Seller"):
            if seller in ("Unassigned Coverage ID", "Vacant seat", "Unmapped product"):
                continue
            acct = (s.groupby("Customer")["ACV (USD at plan)"].sum().sort_values(ascending=False))
            acct = acct[acct > 0]
            if acct.empty:
                continue
            top = ", ".join(f"{c} ({fmt(v)})" for c, v in acct.head(top_n).items())
            more = f" +{len(acct) - top_n} more" if len(acct) > top_n else ""
            lines.append(f"  • {seller}: {top}{more}")
        lines.append("")
    if not issues.empty:
        lines.append(f":warning: {len(issues)} item(s) need an owner — see 'Needs Attention' tab.")
    return "\n".join(lines)


def post_slack(webhook, text):
    req = urllib.request.Request(webhook, data=json.dumps({"text": text}).encode(),
                                 headers={"Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=30).read()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw")
    ap.add_argument("--seller-map")
    ap.add_argument("--dates-only", action="store_true", help="just print the SSABL date range")
    ap.add_argument("--out")
    ap.add_argument("--today", help="YYYY-MM-DD; defaults to today")
    ap.add_argument("--slack-webhook", default=os.environ.get("SLACK_WEBHOOK_URL"))
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    start, end = report_window(today)
    print(f"Pull SSABL with Start Date {start} and End Date {end}")
    if a.dates_only:
        return
    if not (a.raw and a.seller_map):
        ap.error("--raw and --seller-map are required")

    df = load_raw(a.raw)
    sm = load_seller_map(a.seller_map)
    m, detail = build(df, sm, start, end)
    out = a.out or f"SSABL_Report_{end:%m-%d-%y}.xlsx"
    by_seller, issues = write_excel(detail, out, start, end, m["SAAS ACV USD AT PLAN"].sum())
    text = slack_text(detail, issues, start, end)
    with open(os.path.splitext(out)[0] + "_slack.txt", "w") as f:
        f.write(text)
    print(f"Rows in scope: {len(m)}  Total ACV: {m['SAAS ACV USD AT PLAN'].sum():,.2f}")
    print(f"Wrote {out}")
    if a.slack_webhook:
        post_slack(a.slack_webhook, text); print("Posted to Slack")
    else:
        print("No Slack webhook set; message saved next to the report.")


if __name__ == "__main__":
    main()
