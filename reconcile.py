"""
reconcile.py
------------
Pure reconciliation logic.
Call: run(invoice_bytes, portal_bytes) -> bytes (xlsx output)
No file paths, no Flask. Fully testable standalone.
"""

import io
import pandas as pd
import numpy as np


def run(invoice_bytes: bytes, portal_bytes: bytes) -> bytes:
    # ── Load (calamine engine for fast parsing) ───────────────────────────
    df1 = pd.read_excel(
        io.BytesIO(invoice_bytes), sheet_name=0, engine="calamine",
        dtype={"invoiceno": str, "supplier_gstin": str}
    )
    df2 = pd.read_excel(
        io.BytesIO(portal_bytes), sheet_name=0, engine="calamine",
        dtype={"invoiceno": str, "supplier_gstin": str}
    )

    df1.columns = [
        "State Code", "supplier_gstin", "Recipient_gstin", "Recipient_name",
        "invoiceno", "invoice_date", "invoice_date_str", "taxablevalue",
        "cgst", "sgst", "igst", "cess", "Totaltax", "Total Value",
        "DocumentType", "FileName"
    ]
    df2.columns = [
        "EWBNo", "EWBDate", "SupplyType", "invoiceno", "invoice_date",
        "invoice_date_str", "DocumentType", "OtherPartyGSTIN",
        "TransporterDetails", "supplier_gstin", "Recipient_gstin",
        "FromGSTINInfo", "Recipient_name", "status", "NoofItems",
        "MainHSNCode", "MainHSNDesc", "taxablevalue", "sgst", "cgst",
        "igst", "cess", "CESSNonAdvolValue", "OtherValue", "invoice_value",
        "period", "OtherPartyRejectionStatus", "IRN", "GenMode"
    ]

    df1["Datasource"] = "Invoice"
    df2["Datasource"] = "Portal"

    # ── Merge ─────────────────────────────────────────────────────────────
    inv_cols = ["Datasource", "supplier_gstin", "Recipient_gstin",
                "Recipient_name", "DocumentType", "invoiceno",
                "invoice_date", "taxablevalue", "cgst", "sgst", "igst", "cess"]
    df = pd.concat([df1[inv_cols], df2[inv_cols]], ignore_index=True)

    df["DS_val"] = df["Datasource"].map({"Invoice": 1, "Portal": -1})

    for col in ["cgst", "sgst", "igst", "cess", "taxablevalue", "Recipient_gstin",
                "Recipient_name", "invoiceno"]:
        if col in ("Recipient_gstin", "Recipient_name", "invoiceno"):
            df[col] = df[col].fillna("No_Data")
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Totaltax"] = (df["cgst"] + df["sgst"] + df["igst"] + df["cess"]).round()
    df["Totaltax_A"] = df["Totaltax"].where(df["Datasource"] == "Invoice", -df["Totaltax"])
    df["taxablevalue_A"] = df["taxablevalue"].where(df["Datasource"] == "Invoice", -df["taxablevalue"])

    grp_key = ["Datasource", "Recipient_gstin", "invoiceno", "invoice_date"]
    df["sum_Totaltax"] = df.groupby(grp_key)["Totaltax"].transform("sum")
    df["sum_Taxable"]  = df.groupby(grp_key)["taxablevalue"].transform("sum")
    df["sum_Taxable_Roff"]  = (np.floor(df["sum_Taxable"].round() / 10)) * 10
    df["sum_Totaltax_Roff"] = (np.floor(df["sum_Totaltax"] / 10)) * 10

    df["temp_date"] = pd.to_datetime(df["invoice_date"], dayfirst=True, errors="coerce")
    df["YearMonth"] = df["temp_date"].dt.year * 100 + df["temp_date"].dt.month

    def get_fy(d):
        return f"{d.year-1}-{str(d.year)[2:]}" if d.month <= 3 \
               else f"{d.year}-{str(d.year+1)[2:]}"

    df["FY"] = df["temp_date"].apply(get_fy)

    df["header"]      = df["Recipient_gstin"].str[:2]
    df["user_header"] = df["supplier_gstin"].str[:2]
    df["header1"] = "Header_OK"
    df.loc[(df["header"] != df["user_header"]) & (df["cgst"] > 0) & (df["sgst"] > 0), "header1"] = "IGST Not Levied"
    df.loc[(df["header"] == df["user_header"]) & (df["igst"] > 0), "header1"] = "IGST Incorrectly Levied"

    a, b = -10, 10

    # ── Vectorized stage() ────────────────────────────────────────────────
    def stage(df, label, grp, extra=None, need_count=False):
        s = df.groupby(grp)["Totaltax_A"].transform("sum")
        rmk = s.between(a, b) & (df["Totaltax"] != 0)
        if extra is not None:
            rmk = rmk & extra
        fu = df.groupby(["Datasource"] + grp).cumcount() + 1
        dsv = df["DS_val"].where(fu == 1, 0)
        dsv_col = f"__dsv_{label}"
        df = df.assign(**{dsv_col: dsv})
        ds_rmk = df.groupby(grp)[dsv_col].transform("sum")
        df.drop(columns=[dsv_col], inplace=True)
        mask = df["Match_Type"].isnull() & rmk & (ds_rmk == 0)
        if need_count:
            cnt = df.groupby(grp)["Totaltax_A"].transform("count")
            mask = mask & (cnt > 1) & (df["sum_Totaltax_Roff"] > 0)
        df.loc[mask & (s == 0), "Match_Type"] = f"1_Exactly_Matched_{label}"
        df.loc[mask & (s != 0), "Match_Type"] = f"2_Tolerance_Matched_{label}"
        return df

    df["Match_Type"] = pd.NA

    df = stage(df, "A1", ["Recipient_gstin", "invoiceno", "invoice_date", "sum_Taxable_Roff", "sum_Totaltax_Roff"])
    df = stage(df, "A2", ["Recipient_gstin", "invoiceno", "invoice_date"])
    df = stage(df, "E1", ["Recipient_gstin", "invoiceno"], need_count=True)
    df = stage(df, "H1", ["Recipient_gstin", "invoice_date"], need_count=True)
    df = stage(df, "J1", ["Recipient_gstin", "YearMonth", "sum_Taxable_Roff", "sum_Totaltax_Roff"], need_count=True)
    df = stage(df, "K1", ["Recipient_gstin", "YearMonth", "sum_Totaltax_Roff"], need_count=True)
    df = stage(df, "L1", ["Recipient_gstin"], need_count=True)
    df = stage(df, "M1", ["invoiceno", "invoice_date", "sum_Taxable_Roff", "sum_Totaltax_Roff"], need_count=True)
    df = stage(df, "M2", ["invoice_date", "sum_Taxable_Roff", "sum_Totaltax_Roff"], need_count=True)

    # ── N1 — unmatched excess/short invoicewise (vectorized) ─────────────
    fu_n1  = df.groupby(["Datasource", "Recipient_gstin", "invoiceno"]).cumcount() + 1
    dsv_n1 = df["DS_val"].where(fu_n1 == 1, 0)
    df = df.assign(__dsv_n1=dsv_n1)
    ds_rmk_n1 = df.groupby(["Recipient_gstin", "invoiceno"])["__dsv_n1"].transform("sum")
    df.drop(columns=["__dsv_n1"], inplace=True)

    mask_n1 = df["Match_Type"].isnull() & (df["Totaltax"] != 0) & (ds_rmk_n1 == 0)
    df.loc[mask_n1, "Match_Type"] = "9_UnMatched_Excess_or_Short_Invoicewise"

    # N2 — datewise tax match only
    fu2  = df.groupby(["Datasource", "Recipient_gstin", "invoice_date"]).cumcount() + 1
    dsv2 = df["DS_val"].where(fu2 == 1, 0)
    df.loc[df["Match_Type"].isnull() & (df["Totaltax"] != 0) & (dsv2 == 0), "Match_Type"] = "10_Matched_GST_DT_Tax_Only"

    # O1 — yearmonth match only
    fu3  = df.groupby(["Datasource", "Recipient_gstin", "YearMonth"]).cumcount() + 1
    dsv3 = df["DS_val"].where(fu3 == 1, 0)
    df.loc[df["Match_Type"].isnull() & (df["Totaltax"] != 0) & (dsv3 == 0), "Match_Type"] = "11_Matched_GST_YM_Tax_Only"

    # ── URP / Export logic ────────────────────────────────────────────────
    urp_mask = df["Match_Type"].isnull() & (df["Recipient_gstin"] == "URP")
    pr_sum  = df[df["Datasource"] == "Invoice"][urp_mask].groupby("invoiceno")["taxablevalue"].sum()
    por_sum = df[df["Datasource"] == "Portal"][urp_mask].groupby("invoiceno")["taxablevalue"].sum()
    df["_sum_pr"]  = df["invoiceno"].map(pr_sum).fillna(0)
    df["_sum_por"] = df["invoiceno"].map(por_sum).fillna(0)
    in_both = df["invoiceno"].isin(pr_sum.index) & df["invoiceno"].isin(por_sum.index)
    diff = (df["_sum_pr"] - df["_sum_por"]).round(2)
    df.loc[urp_mask & in_both & (diff == 0),                                    "Match_Type"] = "Export_Matched"
    df.loc[urp_mask & in_both & df["Match_Type"].isnull() & diff.between(-10, 10), "Match_Type"] = "Export_Matched_Tolerance"
    df.loc[urp_mask & in_both & df["Match_Type"].isnull(),                      "Match_Type"] = "Export_UnMatched"

    df.loc[df["Match_Type"].isnull() & (df["Datasource"] == "Invoice"), "Match_Type"] = "Available_In_PR_Not_In_Portal"
    df.loc[df["Match_Type"].isnull() & (df["Datasource"] == "Portal"),  "Match_Type"] = "Available_In_Portal_Not_In_PR"

    # ── Categories ────────────────────────────────────────────────────────
    completely = {"1_Exactly_Matched_A1", "2_Tolerance_Matched_A1",
                  "1_Exactly_Matched_A2", "2_Tolerance_Matched_A2",
                  "Export_Matched", "Export_Matched_Tolerance"}
    partially  = {"1_Exactly_Matched_E1", "2_Tolerance_Matched_E1",
                  "1_Exactly_Matched_H1", "2_Tolerance_Matched_H1"}
    probable   = {"1_Exactly_Matched_J1", "2_Tolerance_Matched_J1",
                  "1_Exactly_Matched_K1", "2_Tolerance_Matched_K1",
                  "1_Exactly_Matched_L1", "2_Tolerance_Matched_L1",
                  "1_Exactly_Matched_M1", "2_Tolerance_Matched_M1",
                  "1_Exactly_Matched_M2", "2_Tolerance_Matched_M2"}

    def cat(mt):
        if mt in completely: return "Completely_Matched"
        if mt in partially:  return "Partially_Matched"
        if mt in probable:   return "Probable_Matched"
        return "UnMatched"

    df["Categories"] = df["Match_Type"].apply(cat)
    df["Matching_Results"] = df["Categories"].apply(
        lambda c: "Matched" if c in ("Completely_Matched", "Partially_Matched", "Probable_Matched") else "Mismatch"
    )

    # ── EWB logic ─────────────────────────────────────────────────────────
    thresholds = {
        "03": 100000, "05": 50000, "06": 50000, "07": 100000, "08": 100000,
        "09": 50000,  "10": 100000, "11": 50000, "12": 50000, "18": 50000,
        "19": 50000,  "22": 50000, "23": 100000, "27": 100000, "29": 50000,
        "33": 100000, "36": 50000, "37": 50000,  "21": 50000
    }

    def ewb(row):
        mt  = row["Match_Type"]
        inv = str(row["invoiceno"]).upper()
        val = row["sum_Taxable"]
        rec = str(row["Recipient_gstin"]).upper().strip()
        sup = str(row["supplier_gstin"]).upper().strip()

        if mt == "Available_In_PR_Not_In_Portal":
            if rec in ("URP", "EXPORT", "") or len(rec) < 15:
                return mt
            if any(k in inv for k in ("SI", "JS")):
                return "Service Invoice - EWB not required"
            if any(k in inv for k in ("FSB-MK", "RNT")):
                return "Rental Invoice - EWB not required"
            if "BOS" in inv:
                return "Bill of Supply - Exempt supply"
            if rec[:2] != sup[:2]:
                if val < 50000:
                    return "Inter-state - EWB not required (below 50,000)"
            else:
                lim = thresholds.get(rec[:2], 50000)
                if val < lim:
                    return f"Intra-state - EWB not required (below {lim})"
        if mt == "Available_In_Portal_Not_In_PR":
            return "EWB generated but invoice not in books"
        return mt

    output = df[[
        "supplier_gstin", "header1", "YearMonth", "FY", "DocumentType",
        "Datasource", "Match_Type", "Matching_Results", "Categories",
        "Recipient_gstin", "Recipient_name", "invoiceno", "invoice_date",
        "taxablevalue", "sum_Taxable", "Totaltax_A", "taxablevalue_A",
        "Totaltax", "cgst", "sgst", "igst", "cess"
    ]].copy()

    output["Match_Type"] = output.apply(ewb, axis=1)

    df50 = output[output["Match_Type"] == "Available_In_PR_Not_In_Portal"].copy()
    df51 = output[output["Match_Type"] == "EWB generated but invoice not in books"].copy()

    rename_map = {
        "Available_In_PR_Not_In_Portal":               "Invoice in Books - EWB Missing",
        "Available_In_Portal_Not_In_PR":               "EWB Present - Missing in Books",
        "9_UnMatched_Excess_or_Short_Invoicewise":     "Date or Value Mismatch",
        "1_Exactly_Matched_A1":                        "All Matched",
        "2_Tolerance_Matched_A1":                      "All Matched Within Tolerance",
        "1_Exactly_Matched_A2":                        "All Matched",
        "2_Tolerance_Matched_A2":                      "All Matched Within Tolerance",
        "Export_Matched":                              "Export - Matched",
        "Export_Matched_Tolerance":                    "Export - Matched Within Tolerance",
        "Export_UnMatched":                            "Export - UnMatched",
    }
    output["Match_Type"] = output["Match_Type"].replace(rename_map)

    # ── Pivots ────────────────────────────────────────────────────────────
    df52 = pd.pivot_table(
        output, index=["Categories", "Match_Type"], columns=["Datasource"],
        values=["Totaltax_A", "taxablevalue_A"],
        aggfunc=["sum", "count"], fill_value=0,
        margins=True, margins_name="GrandTotal"
    )

    piv_cols = [c for c in ["taxablevalue", "Totaltax"] if c in df50.columns]
    df56 = pd.pivot_table(df50, index=["Recipient_gstin", "Recipient_name", "invoiceno", "invoice_date"],
                          values=piv_cols, aggfunc="sum", fill_value=0,
                          margins=True, margins_name="GrandTotal") if not df50.empty else pd.DataFrame()

    df57 = pd.pivot_table(df51, index=["Recipient_gstin", "Recipient_name", "invoiceno", "invoice_date"],
                          values=piv_cols, aggfunc="sum", fill_value=0,
                          margins=True, margins_name="GrandTotal") if not df51.empty else pd.DataFrame()

    # ── Write Excel ───────────────────────────────────────────────────────
    out_buf = io.BytesIO()
    writer  = pd.ExcelWriter(out_buf, engine="xlsxwriter")
    wb      = writer.book

    df52.to_excel(writer, sheet_name="Summary")
    output.to_excel(writer, sheet_name="Main_Output", index=False)
    if not df56.empty: df56.to_excel(writer, sheet_name="Avail_Sales")
    if not df57.empty: df57.to_excel(writer, sheet_name="Avail_EWBill")
    df.to_excel(writer, sheet_name="All", index=False)

    # Formatting
    fmt_red    = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    fmt_green  = wb.add_format({"bg_color": "#00FF00", "font_color": "#006100"})
    fmt_yellow = wb.add_format({"bg_color": "#FFFF99", "font_color": "#9C0006"})
    fmt_border = wb.add_format({"border": 1, "border_color": "green", "font_color": "black"})
    fmt_big    = wb.add_format({"font_size": 13})

    ws_sum  = writer.sheets["Summary"]
    ws_main = writer.sheets["Main_Output"]

    ws_sum.conditional_format("C5:N30", {"type": "top",  "value": 1, "format": fmt_red})
    ws_sum.conditional_format("A1:N20", {"type": "cell", "criteria": ">=", "value": 0, "format": fmt_border})
    ws_sum.set_column("A:N", 14, fmt_big)

    ds_col     = output.columns.get_loc("Datasource") + 1
    col_letter = chr(ord("A") + ds_col)
    ws_main.conditional_format(f"{col_letter}2:{col_letter}50000",
                               {"type": "cell", "criteria": "equal to", "value": '"Portal"',  "format": fmt_yellow})
    ws_main.conditional_format(f"{col_letter}2:{col_letter}50000",
                               {"type": "cell", "criteria": "equal to", "value": '"Invoice"', "format": fmt_green})

    writer.close()
    return out_buf.getvalue()
