import re
import math
import pandas as pd
import streamlit as st

# ==========================
# Hardcoded formül sözlüğü
# ==========================
formuller = {
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
def extract_valid_refs(formula: str):
    return re.findall(r"C0\d{3}", str(formula or ""))

def has_invalid_tokens(formula: str) -> bool:
    tokens = re.findall(r"C\w{4}", str(formula or ""))
    return any(re.fullmatch(r"C0\d{3}", t) is None for t in tokens)

def kural4_flags_for_group(insp_list):
    expected = list(range(10, 10 * (len(insp_list) + 1), 10))[:len(insp_list)]
    return ["Doğru" if a == b else "Hatalı sıra artışı" for a, b in zip(insp_list, expected)]

def safe_eval(expr: str) -> float:
    if not re.fullmatch(r"[0-9\.\+\-\*\/\(\)\s]+", expr):
        raise ValueError("İzinli olmayan karakter")
    return float(eval(expr, {"_builtins_": None}, {"math": math}))

def in_range_with_missing(value, low, up):
    if low is None and up is None:
        return None
    if (low is not None) and (up is not None):
        return (value >= low) and (value <= up)
    if low is not None:
        return value >= low
    return value <= up

# ==========================
# KURAL 5
# ==========================
def kural5_for_row(row: pd.Series, group_df: pd.DataFrame, lower_col: str, upper_col: str):
    out = {f"KURAL5_CASE_{i}": "" for i in range(1,5)}
    out["KURAL5_NOT"] = ""

    formula = str(row.get("FORMULA_FIELD_1", "") or "").strip()
    if not formula:
        return out

    refs = extract_valid_refs(formula)
    if len(refs) != 2:
        note = f"KURAL5: Sadece 2 referans destekleniyor; mevcut: {len(refs)}"
        for k in out.keys(): out[k] = note
        return out

    def to_num(x):
        try: return float(x)
        except: return None

    c_low = to_num(row.get(lower_col))
    c_up  = to_num(row.get(upper_col))
    if c_low is None and c_up is None:
        for k in out.keys(): out[k] = "C limiti yok"
        return out

    ins_to_char = {int(r.INSPCHAR): r.MSTR_CHAR for r in group_df.itertuples()}
    lim_low = {r.MSTR_CHAR: to_num(getattr(r, lower_col, None)) for r in group_df.itertuples()}
    lim_up  = {r.MSTR_CHAR: to_num(getattr(r, upper_col, None)) for r in group_df.itertuples()}

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
            out[key] = "Veri eksik"
            continue
        expr = formula
        expr = re.sub(rf"\b{A_ref}\b", str(Aval), expr)
        expr = re.sub(rf"\b{B_ref}\b", str(Bval), expr)
        if extract_valid_refs(expr):
            out[key] = "Veri eksik"; continue
        try: val = safe_eval(expr)
        except: out[key] = "Veri eksik"; continue
        ok = in_range_with_missing(val, c_low, c_up)
        out[key] = "Olumlu" if ok else "Olumsuz" if ok is not None else "C limiti yok"

    return out

# ==========================
# ANA KONTROL (KURAL1–5)
# ==========================
def kontrol_et(df: pd.DataFrame, uretim_yeri: str, lower_col="LW_TOL_LMT", upper_col="UP_TOL_LMT") -> pd.DataFrame:
    formuller_uy = formuller.get(uretim_yeri, {})
    results = []

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
            k5_cols = {f"KURAL5_CASE_{i}": "" for i in range(1,5)}
            k5_cols["KURAL5_NOT"] = ""

            if mstr in formuller_uy:
                expected_chars = formuller_uy[mstr]
                expected_refs = [f"C{char_to_insp[ch]:04d}" for ch in expected_chars if ch in char_to_insp]

                # KURAL1
                if has_invalid_tokens(formula):
                    kural1 = "Hatalı: Geçersiz referans formatı"
                else:
                    refs_in_formula = extract_valid_refs(formula)
                    if set(refs_in_formula) == set(expected_refs) and len(refs_in_formula) == len(expected_refs):
                        kural1 = "Doğru"
                    else:
                        kural1 = f"Hatalı: Beklenen {'-'.join(expected_refs)}"

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
st.title("📊 Formül Kontrol - KURAL1..KURAL5")

uploaded_file = st.file_uploader("Excel dosyasını yükle (.xlsx veya .xls)", type=["xlsx", "xls"])
uretim_yeri = st.selectbox("Üretim Yeri:", list(formuller.keys()))
lower_col = st.text_input("Alt limit kolonu:", value="LW_TOL_LMT")
upper_col = st.text_input("Üst limit kolonu:", value="UP_TOL_LMT")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("🔍 İlk birkaç satır:")
    st.dataframe(df.head())

    if st.button("Kontrolü Başlat"):
        try:
            out = kontrol_et(df, uretim_yeri, lower_col=lower_col, upper_col=upper_col)
            st.success("Kontrol tamamlandı ✅")
            st.dataframe(out.head())

            from io import BytesIO
            output = BytesIO()
            out.to_excel(output, index=False, engine="openpyxl")
            st.download_button(
                label="📥 Sonucu Excel olarak indir",
                data=output.getvalue(),
                file_name="kontrol_sonuclari.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Hata: {e}")
