# newappcheck.py – Final Build
# Fully robust skip handling (Not/And/Or, parentheses, dots, brackets)
# Two-way skip validation

import streamlit as st
import pandas as pd
import numpy as np
import re, io
from datetime import datetime

st.set_page_config(page_title="KnowledgeExcel — DV Automation (newappcheck)", layout="wide")
st.title("KnowledgeExcel — Data Validation Automation (newappcheck)")

DEFAULT_DK_CODES = [88, 99]
DEFAULT_DK_TOKENS = ["DK","Refused","Don't know","Dont know","Refuse","REFUSED"]

# ---------- Read uploaded files ----------
def read_any_df(uploaded):
    if uploaded is None:
        return None
    data = io.BytesIO(uploaded.read())
    name = uploaded.name.lower()
    try:
        if name.endswith(("xlsx","xls")):
            return pd.read_excel(data, engine="openpyxl")
        else:
            return pd.read_csv(data, encoding="utf-8-sig")
    except Exception:
        data.seek(0)
        return pd.read_csv(data, encoding="ISO-8859-1")

# ---------- Robust Skip Parser ----------
def parse_skip_expression_to_mask(expr, df):
    """
    Robust Sawtooth skip parser.
    Handles:
      - <>, =, AND/OR/NOT (any case)
      - Parentheses, space/no-space syntax
      - Variable names with dots, brackets, underscores
      - Numeric coercion
    """
    expr_orig = str(expr)
    try:
        e = expr_orig
        # --- normalize ---
        e = e.replace("<>", "!=")
        e = re.sub(r'(?<![!<>=])=(?!=)', '==', e)
        e = re.sub(r'(?i)\bAND\b', '&', e)
        e = re.sub(r'(?i)\bOR\b', '|', e)
        e = re.sub(r'(?i)\bNOT\s*\(', '~(', e)
        e = re.sub(r'\s+', ' ', e)

        # --- replace variables with df references ---
        cols = sorted(df.columns, key=len, reverse=True)
        for col in cols:
            safe = re.escape(col)
            e = re.sub(rf'(?<!\w){safe}(?!\w)',
                       f"pd.to_numeric(df[{repr(col)}], errors='coerce')", e)

        mask = eval(e, {"df": df, "pd": pd, "np": np})
        return pd.Series(mask, index=df.index).fillna(False).astype(bool)
    except Exception as err:
        st.warning(f"Skip Parsing Error for '{expr_orig}': {err}")
        return pd.Series(False, index=df.index)

# ---------- Helper ----------
def format_ids(series, n=200):
    return ";".join(map(str, series.astype(str).unique()[:n]))

# END OF PART 1
# ---------- Streamlit inputs ----------
raw_file = st.sidebar.file_uploader("Raw Data (xlsx/csv)", type=["xlsx","xls","csv"])
skips_file = st.sidebar.file_uploader("Sawtooth Skips (csv/xlsx)", type=["csv","xlsx"])
run_btn = st.sidebar.button("Run Validation")

def detect_junk_oe(v, repeat_min=4, min_len=2):
    if pd.isna(v): return False
    s = str(v).strip()
    if s == "" or (s.isdigit() and len(s) <= 3): return True
    if re.match(rf'^(.)\1{{{repeat_min-1},}}$', s): return True
    if len(s) <= min_len: return True
    return False

def group_variables(cols):
    groups = {}
    for c in cols:
        m = re.match(r"^(.*?)(_?\d+|R\d+)?$", c, flags=re.I)
        if m:
            prefix = re.sub(r"[_Rr\d]+$", "", m.group(1))
        else:
            prefix = c
        groups.setdefault(prefix, []).append(c)
    return groups

# ---------- Two-way skip validator ----------
def check_skip_violations(var, expr, df, id_col):
    out = []
    try:
        mask = parse_skip_expression_to_mask(expr, df)
        ans = df[var].astype(str).fillna("").str.strip()
        blank = ans.eq("") | ans.str.lower().isin(["na","n/a","nan","none"])
        v1 = df[mask & ~blank]      # answered when should skip
        v2 = df[~mask & blank]      # skipped when should answer
        if len(v1)>0:
            out.append(("Skip Violation (Answered when should Skip)",
                        len(v1), format_ids(v1[id_col])))
        if len(v2)>0:
            out.append(("Skip Violation (Skipped when should Answer)",
                        len(v2), format_ids(v2[id_col])))
    except Exception as e:
        out.append(("Skip Parsing Error", 0, f"{expr} | {e}"))
    return out

# ---------- Run ----------
if run_btn:
    if not raw_file or not skips_file:
        st.error("Please upload both Raw Data and Skips files.")
        st.stop()

    df = read_any_df(raw_file)
    skips = read_any_df(skips_file)
    id_col = next((c for c in df.columns if str(c).lower() in
                   ["respondentid","resp_id","id","sys_respnum"]), df.columns[0])

    st.info(f"Respondent ID column detected → **{id_col}**")
    data_vars = [c for c in df.columns if not str(c).lower().startswith("sys_")]
    rules, findings = [], []

    lc = {c.lower(): c for c in skips.columns}
    from_col = lc.get("skip from") or list(skips.columns)[0]
    logic_col = lc.get("logic") or lc.get("condition") or None
    to_col = lc.get("skip to") or lc.get("target") or None

    if logic_col:
        for _, r in skips.iterrows():
            src = str(r.get(from_col,"")).strip()
            expr = str(r.get(logic_col,"")).strip()
            tgt = str(r.get(to_col,"")).strip()
            if not src or not expr: continue
            desc = f"Skip {src} when {expr} (Target {tgt})"
            rules.append({"Variable":src,"Type":"Skip","Rule Applied":expr,"Description":desc})

    st.write("Running skip validations…")
    for rule in rules:
        var, expr = rule["Variable"], rule["Rule Applied"]
        if var not in df.columns: continue
        for ctype,count,ids in check_skip_violations(var, expr, df, id_col):
            findings.append({"Variable":var,"Check_Type":ctype,
                             "Description":rule["Description"],
                             "Affected_Count":count,"Respondent_IDs":ids})

    if findings:
        rep = pd.DataFrame(findings)
        st.subheader("Validation Results – Preview")
        st.dataframe(rep.head(200), use_container_width=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            rep.to_excel(w, "Detailed Checks", index=False)
            rep.groupby("Check_Type", as_index=False)["Affected_Count"].sum()\
               .to_excel(w, "Summary", index=False)
            pd.DataFrame({"Generated":[datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")],
                          "Respondent ID":[id_col],
                          "Rows":[df.shape[0]],
                          "Cols":[df.shape[1]]}).to_excel(w, "Project Info", index=False)
        buf.seek(0)
        st.download_button("📥 Download Validation Report.xlsx", data=buf,
            file_name="Validation Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("Validation Report ready for download.")
    else:
        st.warning("No skip violations found.")

# EOF
