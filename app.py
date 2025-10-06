import re
import math
from typing import List, Dict, Any
import pandas as pd
import streamlit as st
from io import BytesIO

# ---------------------------------------------------------
# Sabit kolon adları (ekranda gösterilmiyor)
LOWER_COL = "LW_TOL_LMT"
UPPER_COL = "UP_TOL_LMT"
# ---------------------------------------------------------

# ==========================
# Hardcoded formül sözlüğü
# ==========================
formuller: Dict[str, Dict[str, List[str]]] = {
    "3101": {
        "YKM G/G":  ["KM G/G", "YAG G/G"],
        "YKM2 G/G": ["KM2 G/G", "YAG2 G/G"],
        "YKM3 G/G": ["KM3 G/G", "YAG3 G/G"],
        "LOS2":     ["KM2 G/G", "YAG2 G/G", "PRT2 G/G"],
        "LOS3":     ["KM3 G/G", "YAG3 G/G", "PRT3 G/G"],
        "KMT":      ["TUZ", "KM G/G"],
        "KMY":      ["YAG G/G", "KM G/G"],
    },
    "3102": {
        "YKM G/G":  ["KM G/G", "YAG G/G"],
        "YKM2 G/G": ["KM2 G/G", "YAG2 G/G"],
        "YKM3 G/G": ["KM3 G/G", "YAG3 G/G"],
        "LOS2":     ["KM2 G/G", "YAG2 G/G", "PRT2 G/G"],
        "LOS3":     ["KM3 G/G", "YAG3 G/G", "PRT3 G/G"],
        "KMT":      ["TUZ", "KM G/G"],
        "KMY":      ["YAG G/G", "KM G/G"],
    },
    "3103": {
        "YKM G/G":  ["KM G/G", "YAG G/G"],
        "YKM2 G/G": ["KM2 G/G", "YAG2 G/G"],
        "YKM3 G/G": ["KM3 G/G", "YAG3 G/G"],
        "LOS2":     ["KM2 G/G", "YAG2 G/G", "PRT2 G/G"],
        "LOS3":     ["KM3 G/G", "YAG3 G/G", "PRT3 G/G"],
        "KMY":      ["YAG3 G/G", "KM3 G/G"],
        "KMT":      ["TUZ", "KM3 G/G"],
        "KMY3":     ["YAG3 G/G", "KM3 G/G"],
        "KMT3":     ["TUZ", "KM3 G/G"],
    },
    "2901": {
        "TOPLAMBD": ["NEM", "YAG", "PROTEIN", "KUL"],
        "Kx100/P":  ["KOLAJEN", "PROTEIN"],
        "SKx100/P": ["SKOLAJEN", "SPROTEIN"],
        "SY/SP":    ["SYAG", "SPROTEIN"],
        "Y/P":      ["YAG", "PROTEIN"],
        "SN/SP":    ["SNEM", "SPROTEIN"],
        "N/P":      ["NEM", "PROTEIN"],
    },
}

# ==========================
# Yardımcı fonksiyonlar
# ==========================
def extract_valid_refs(formula: str) -> List[str]:
    return re.findall(r"C0\d{3}", str(formula or ""))

def has_invalid_tokens(formula: str) -> bool:
    tokens = re.findall(r"C\w{4}", str(formula or ""))
    return any(re.fullmatch(r"C0\d{3}", t) is None for t in tokens)

def kural4_flags_for_group(insp_list: List[int]) -> List[str]:
    expected = list(range(10, 10 * (len(insp_list) + 1), 10))[:len(insp_list)]
    return ["Doğru" if a == b else "Hatalı sıra artışı" for a, b in zip(insp_list, expected)]

def safe_eval(expr: str) -> float:
    if not re.fullmatch(r"[0-9\.\+\-\*\/\(\)\s]+", expr):
        raise ValueError("İzinli olmayan karakter")
    return float(eval(expr, {"__builtins__": {}}, {"math": math}))

def in_range_with_missing(value, low, up, eps: float = 1e-9):
    if low is None and up is None:
        return None
    if (low is not None) and (up is not None):
        return (value >= (low - eps)) and (value <= (up + eps))
    if low is not None:
        return value >= (low - eps)
    return value <= (up + eps)

# ==========================
# KURAL 5 (2 referans)
# ==========================
def kural5_for_row(row: pd.Series, group_df: pd.DataFrame,
                   lower_col: str = LOWER_COL, upper_col: str = UPPER_COL) -> Dict[str, str]:
    out = {"KURAL5_CASE_1": "", "KURAL5_CASE_2": "", "KURAL5_CASE_3": "", "KURAL5_CASE_4": "", "KURAL5_NOT": ""}
    formula = str(row.get("FORMULA_FIELD_1", "") or "").strip()
    if not formula:
        return out

    refs = extract_valid_refs(formula)
    if len(refs) != 2:
        note = f"KURAL5: Sadece 2 referans destekleniyor; mevcut: {len(refs)}"
        for k in ["KURAL5_CASE_1","KURAL5_CASE_2","KURAL5_CASE_3","KURAL5_CASE_4"]:
            out[k] = note
        out["KURAL5_NOT"] = note
        return out

    def _to_num_or_none(x):
        try:
            return None if pd.isna(x) else float(x)
        except Exception:
            return None

    c_low = _to_num_or_none(row.get(lower_col))
    c_up  = _to_num_or_none(row.get(upper_col))
    if c_low is None and c_up is None:
        for k in ["KURAL5_CASE_1","KURAL5_CASE_2","KURAL5_CASE_3","KURAL5_CASE_4"]:
            out[k] = "FormülChar limiti yok"
        return out

    ins_to_char = {int(r.INSPCHAR): r.MSTR_CHAR for r in group_df.itertuples()}
    lim_low = {r.MSTR_CHAR: _to_num_or_none(getattr(r, lower_col, None)) for r in group_df.itertuples()}
    lim_up  = {r.MSTR_CHAR: _to_num_or_none(getattr(r, upper_col, None))  for r in group_df.itertuples()}

    A_ref, B_ref = refs
    A_name = ins_to_char.get(int(A_ref[1:]))
    B_name = ins_to_char.get(int(B_ref[1:]))

    cases = [
        (lim_low.get(A_name), lim_low.get(B_name), "KURAL5_CASE_1"),
        (lim_low.get(A_name), lim_up.get(B_name),  "KURAL5_CASE_2"),
        (lim_up.get(A_name),  lim_low.get(B_name), "KURAL5_CASE_3"),
        (lim_up.get(A_name),  lim_up.get(B_name),  "KURAL5_CASE_4"),
    ]

    for Aval, Bval, key in cases:
        if Aval is None or Bval is None:
            out[key] = "Veri eksik"; continue
        expr = re.sub(rf"\b{A_ref}\b", str(Aval), formula)
        expr = re.sub(rf"\b{B_ref}\b", str(Bval), expr)
        if extract_valid_refs(expr):
            out[key] = "Veri eksik"; continue
        try:
            val = safe_eval(expr)
        except Exception:
            out[key] = "Veri eksik"; continue
        ok = in_range_with_missing(val, c_low, c_up, eps=1e-9)
        out[key] = "Olumlu" if ok else "Olumsuz" if ok is not None else "FormülChar limiti yok"
    return out

# ==========================
# ANA KONTROL (KURAL1–KURAL5)
# ==========================
def kontrol_et(df: pd.DataFrame, uretim_yeri: str,
               lower_col=LOWER_COL, upper_col=UPPER_COL) -> pd.DataFrame:
    results: List[Dict[str, Any]] = []
    formuller_uy = formuller.get(uretim_yeri, {})

    for (pg, op), group in df.groupby(["PLAN_GROUP", "OPER_NUM"]):
        g = group.sort_values("INSPCHAR").reset_index(drop=True)
        char_to_insp = {r.MSTR_CHAR: int(r.INSPCHAR) for r in g.itertuples()}
        insp_to_char = {int(r.INSPCHAR): r.MSTR_CHAR for r in g.itertuples()}
        k4_list = kural4_flags_for_group(g["INSPCHAR"].astype(int).tolist())

        for idx, row in g.iterrows():
            mstr = row["MSTR_CHAR"]
            inspchar = int(row["INSPCHAR"])
            formula = str(row.get("FORMULA_FIELD_1", "")).strip()
            kural4 = k4_list[idx]

            kural1 = kural2 = kural3 = ""
            k5_cols = {"KURAL5_CASE_1": "", "KURAL5_CASE_2": "", "KURAL5_CASE_3": "", "KURAL5_CASE_4": "", "KURAL5_NOT": ""}

            if mstr in formuller_uy:
                expected_chars = formuller_uy[mstr]
                expected_refs = [f"C{char_to_insp[ch]:04d}" for ch in expected_chars if ch in char_to_insp]

                # KURAL1
                if has_invalid_tokens(formula):
                    kural1 = "Hatalı: Geçersiz referans formatı"
                else:
                    refs_in_formula = extract_valid_refs(formula)
                    kural1 = "Doğru" if (set(refs_in_formula) == set(expected_refs) and
                                         len(refs_in_formula) == len(expected_refs)) \
                                    else f"Hatalı: Beklenen {'-'.join(expected_refs)}"

                # KURAL2
                ref_nums = [int(r[1:]) for r in extract_valid_refs(formula)]
                kural2 = "Uygun" if (not ref_nums or max(ref_nums) <= inspchar) else "Uygun Değil"

                # KURAL3
                actual_chars = [insp_to_char.get(int(r[1:]), "") for r in extract_valid_refs(formula)]
                eksik = [c for c in expected_chars if c not in actual_chars]
                fazla = [c for c in actual_chars if c not in expected_chars and c != ""]
                msgs = []
                if eksik: msgs.append("Eksik: " + ", ".join(eksik))
                if fazla: msgs.append("Fazla: " + ", ".join(fazla))
                if actual_chars and actual_chars != expected_chars:
                    msgs.append("Sıra Hatası")
                kural3 = "; ".join(msgs) if msgs else "Doğru"

                # KURAL5
                k5_cols = kural5_for_row(row, g, lower_col=lower_col, upper_col=upper_col)

            results.append({
                **row.to_dict(),
                "KURAL1_DURUM": kural1,
                "KURAL2_DURUM": kural2,
                "KURAL3_KONTROL": kural3,
                "KURAL4_DURUM": kural4,
                **k5_cols
            })

    return pd.DataFrame(results)

# ==========================
# STREAMLIT ARAYÜZÜ
# ==========================
st.set_page_config(page_title="Formül Kontrol (KURAL1..KURAL5)", page_icon="📊", layout="wide")
st.title("📊 Formül Kontrol — KURAL1..KURAL5")

with st.sidebar:
    uretim_yeri = st.selectbox("Üretim Yeri:", list(formuller.keys()), index=0)
    st.caption(f"Alt/Üst limit kolonları dahili olarak '{LOWER_COL}' / '{UPPER_COL}' varsayılanlarıyla okunur.")

uploaded_file = st.file_uploader("Excel dosyasını yükle (.xlsx veya .xls)", type=["xlsx", "xls"])

def _download_template():
    cols = ["PLAN_GROUP","OPER_NUM","OPER_DESC","INSPCHAR","MSTR_CHAR","FORMULA_FIELD_1", LOWER_COL, UPPER_COL]
    template = pd.DataFrame(columns=cols)
    buf = BytesIO()
    template.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()

st.download_button("📥 Şablon Excel indir", data=_download_template(),
                   file_name="sablon.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Excel okunamadı: {e}")
        st.stop()

    expected_cols = {"PLAN_GROUP","OPER_NUM","OPER_DESC","INSPCHAR","MSTR_CHAR","FORMULA_FIELD_1"}
    missing = [c for c in expected_cols if c not in df.columns]
    if missing:
        st.error(f"Zorunlu kolonlar eksik: {', '.join(missing)}")
        st.dataframe(df.head(), use_container_width=True)
        st.stop()

    st.subheader("🔍 İlk satırlar")
    st.dataframe(df.head(20), use_container_width=True)

    if st.button("Kontrolü Başlat", type="primary"):
        try:
            out = kontrol_et(df, uretim_yeri, lower_col=LOWER_COL, upper_col=UPPER_COL)
        except Exception as e:
            st.error(f"Hata: {e}")
            st.stop()

        st.success("Kontrol tamamlandı ✅")
        st.subheader("📄 Sonuç")
        st.dataframe(out, use_container_width=True)

        buf = BytesIO()
        out.to_excel(buf, index=False, engine="openpyxl")
        st.download_button("📥 Sonucu Excel olarak indir",
                           data=buf.getvalue(),
                           file_name="kontrol_sonuclari.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Başlamak için bir Excel dosyası yükleyin.")
