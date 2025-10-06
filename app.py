import re
import math
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any

# ------------------------------------------------------------------------------
# FORMÜL TANIMLARI (BİLGİ AMAÇLI - KODDA İŞLEME GİRMEZ)
#
# 3101 Üretim Yeri Formülleri:
#   YKM G/G   = KM G/G - YAG G/G
#   YKM2 G/G  = KM2 G/G - YAG2 G/G
#   YKM3 G/G  = KM3 G/G - YAG3 G/G
#   LOS2      = KM2 G/G - YAG2 G/G - PRT2 G/G
#   LOS3      = KM3 G/G - YAG3 G/G - PRT3 G/G
#   KMT       = (TUZ / KM G/G) * 100
#   KMY       = (YAG G/G / KM G/G) * 100
#
# 3102 Üretim Yeri Formülleri:
#   YKM G/G   = KM G/G - YAG G/G
#   YKM2 G/G  = KM2 G/G - YAG2 G/G
#   YKM3 G/G  = KM3 G/G - YAG3 G/G
#   LOS2      = KM2 G/G - YAG2 G/G - PRT2 G/G
#   LOS3      = KM3 G/G - YAG3 G/G - PRT3 G/G
#   KMT       = (TUZ / KM G/G) * 100
#   KMY       = (YAG G/G / KM G/G) * 100
#
# 3103 Üretim Yeri Formülleri:
#   YKM G/G   = KM G/G - YAG G/G
#   YKM2 G/G  = KM2 G/G - YAG2 G/G
#   YKM3 G/G  = KM3 G/G - YAG3 G/G
#   LOS2      = KM2 G/G - YAG2 G/G - PRT2 G/G
#   LOS3      = KM3 G/G - YAG3 G/G - PRT3 G/G
#   KMY       = (YAG3 G/G / KM3 G/G) * 100
#   KMT       = (TUZ / KM3 G/G) * 100
#   KMY3      = (YAG3 G/G / KM3 G/G) * 100
#   KMT3      = (TUZ / KM3 G/G) * 100
#
# 2901 Üretim Yeri Formülleri:
#   TOPLAMBD  = NEM + YAG + PROTEIN + KUL
#   Kx100/P   = (KOLAJEN * 100) / PROTEIN
#   SKx100/P  = (SKOLAJEN * 100) / SPROTEIN
#   SY/SP     = SYAG / SPROTEIN
#   Y/P       = YAG / PROTEIN
#   SN/SP     = SNEM / SPROTEIN
#   N/P       = NEM / PROTEIN
#
# NOT: Bu blok yalnızca bilgi amaçlıdır. Asıl kontrol, aşağıdaki "formuller" sözlüğünde
# hangi karakteristiklerin formülde bulunması gerektiği üzerinden yapılır. İşleçler
# (toplama/çıkarma/bölme/*100) FORMULA_FIELD_1 içeriğinden okunur.
# ------------------------------------------------------------------------------

# ==========================
# Hardcoded formül sözlüğü
# (HANGİ DEĞİŞKENLERİN bulunması gerektiğini belirtir)
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
    # 3102 isim güncellemeleri uygulanmış
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
    """C0XXX formatındaki referansları sırayla döndürür."""
    return re.findall(r"C0\d{3}", str(formula or ""))

def has_invalid_tokens(formula: str) -> bool:
    """
    C ile başlayan 5 karakterli tüm tokenları tarar; C0XXX dışındakileri hatalı sayar.
    Ör: CC0050, C00100, CABC12 -> hatalı
    """
    tokens = re.findall(r"C\w{4}", str(formula or ""))
    return any(re.fullmatch(r"C0\d{3}", t) is None for t in tokens)

def kural4_flags_for_group(insp_list: List[int]) -> List[str]:
    """INSPCHAR değerleri 10'ar artıyor mu? (10,20,30,...)"""
    expected = list(range(10, 10 * (len(insp_list) + 1), 10))[:len(insp_list)]
    return ["Doğru" if a == b else "Hatalı sıra artışı" for a, b in zip(insp_list, expected)]

def safe_eval(expr: str) -> float:
    """Yalnız +-*/ ve parantez içeren sayısal ifadeleri güvenli değerlendirir."""
    if not re.fullmatch(r"[0-9\.\+\-\*\/\(\)\s]+", expr):
        raise ValueError("İzinli olmayan karakter")
    return float(eval(expr, {"__builtins__": None}, {"math": math}))

def in_range_with_missing(value, low, up, eps: float = 1e-9):
    """
    C için tek/çift limit durumlarına göre kapsayıcı uygunluk:
    low <= value <= up (eps toleransı ile)
    """
    if low is None and up is None:
        return None
    if (low is not None) and (up is not None):
        return (value >= (low - eps)) and (value <= (up + eps))
    if low is not None:
        return value >= (low - eps)
    return value <= (up + eps)

# ==========================
# KURAL 5 (Test, 2 referans)
# ==========================
def kural5_for_row(
    row: pd.Series, group_df: pd.DataFrame,
    lower_col: str = "LW_TOL_LMT", upper_col: str = "UP_TOL_LMT"
) -> Dict[str, str]:
    """
    2 referans içeren formül satırları (C) için 4 case üretir:
      Case1: A_low vs B_low
      Case2: A_low vs B_up
      Case3: A_up  vs B_low
      Case4: A_up  vs B_up
    İşleçler FORMULA_FIELD_1 içeriğinden okunur (toplama/çıkarma/bölme/*100/parantez).
    """
    out = {"KURAL5_CASE_1": "", "KURAL5_CASE_2": "", "KURAL5_CASE_3": "", "KURAL5_CASE_4": "", "KURAL5_NOT": ""}

    formula = str(row.get("FORMULA_FIELD_1", "") or "").strip()
    if not formula:
        return out

    refs = extract_valid_refs(formula)
    if len(refs) != 2:
        note = f"KURAL5: Sadece 2 referans destekleniyor; mevcut: {len(refs)}"
        out.update({k: note for k in ["KURAL5_CASE_1","KURAL5_CASE_2","KURAL5_CASE_3","KURAL5_CASE_4"]})
        out["KURAL5_NOT"] = note
        return out

    def _to_num_or_none(x):
        try:
            return None if pd.isna(x) else float(x)
        except Exception:
            return None

    # C limitleri (formül sonuç karakteristiği)
    c_low = _to_num_or_none(row.get(lower_col))
    c_up  = _to_num_or_none(row.get(upper_col))
    if c_low is None and c_up is None:
        out.update({k: "FormülChar limiti yok" for k in ["KURAL5_CASE_1","KURAL5_CASE_2","KURAL5_CASE_3","KURAL5_CASE_4"]})
        return out

    # INSPCHAR -> MSTR_CHAR
    ins_to_char = {int(r.INSPCHAR): r.MSTR_CHAR for r in group_df.itertuples()}

    # MSTR_CHAR -> (low, up)
    lim_low, lim_up = {}, {}
    for r in group_df.itertuples():
        l = _to_num_or_none(getattr(r, lower_col, None)) if hasattr(r, lower_col) else None
        u = _to_num_or_none(getattr(r, upper_col, None)) if hasattr(r, upper_col) else None
        lim_low[r.MSTR_CHAR] = l
        lim_up[r.MSTR_CHAR]  = u

    A_ref, B_ref = refs[0], refs[1]
    A_name = ins_to_char.get(int(A_ref[1:]))
    B_name = ins_to_char.get(int(B_ref[1:]))

    if not A_name or not B_name:
        out.update({k: "Veri eksik" for k in ["KURAL5_CASE_1","KURAL5_CASE_2","KURAL5_CASE_3","KURAL5_CASE_4"]})
        out["KURAL5_NOT"] = "A/B karakteristiği bulunamadı"
        return out

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

        # Yerine koyma sonrası hâlâ C0XXX kalırsa (3+ referans vb.) test kapsamı dışında
        if extract_valid_refs(expr):
            out[key] = "Veri eksik"
            continue

        try:
            val = safe_eval(expr)
        except Exception:
            out[key] = "Veri eksik"
            continue

        ok = in_range_with_missing(val, c_low, c_up, eps=1e-9)
        if ok is None:
            out[key] = "FormülChar limiti yok"
        else:
            out[key] = "Olumlu" if ok else "Olumsuz"

    return out

# ==========================
# ANA KONTROL (KURAL1–KURAL5)
# ==========================
def kontrol_et(df: pd.DataFrame, uretim_yeri: str, lower_col="LW_TOL_LMT", upper_col="UP_TOL_LMT") -> pd.DataFrame:
    """
    Beklenen minimum kolonlar:
      PLAN_GROUP, OPER_NUM, OPER_DESC, INSPCHAR, MSTR_CHAR, FORMULA_FIELD_1
    (KURAL5 için opsiyonel: lower_col, upper_col — varsayılan: LW_TOL_LMT / UP_TOL_LMT)
    """
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

            # KURAL4 – INSPCHAR 10’ar artış kontrolü
            kural4 = k4_list[idx]

            # Varsayılanlar
            kural1 = kural2 = kural3 = ""
            k5_cols = {"KURAL5_CASE_1": "", "KURAL5_CASE_2": "", "KURAL5_CASE_3": "", "KURAL5_CASE_4": "", "KURAL5_NOT": ""}

            # Yalnız formül karakteristiği olanlara KURAL1–3–5
            if mstr in formuller_uy:
                expected_chars = formuller_uy[mstr]
                expected_refs = [f"C{char_to_insp[ch]:04d}" for ch in expected_chars if ch in char_to_insp]

                # KURAL1 – Referans formatı + referans seti (matematiksel ifade serbest)
                if has_invalid_tokens(formula):
                    kural1 = "Hatalı: Geçersiz referans formatı"
                else:
                    refs_in_formula = extract_valid_refs(formula)
                    if set(refs_in_formula) == set(expected_refs) and len(refs_in_formula) == len(expected_refs):
                        kural1 = "Doğru"
                    else:
                        kural1 = f"Hatalı: Beklenen {'-'.join(expected_refs)}"

                # KURAL2 – Referans INSPCHAR > satır INSPCHAR olamaz
                ref_nums = [int(r[1:]) for r in extract_valid_refs(formula)]
                kural2 = "Uygun" if (not ref_nums or max(ref_nums) <= inspchar) else "Uygun Değil"

                # KURAL3 – Eksik/Fazla + Sıra kontrolü
                actual_chars = [insp_to_char.get(int(r[1:]), "") for r in extract_valid_refs(formula)]
                eksik = [c for c in expected_chars if c not in actual_chars]
                fazla = [c for c in actual_chars if c not in expected_chars and c != ""]
                msgs = []
                if eksik: msgs.append("Eksik: " + ", ".join(eksik))
                if fazla: msgs.append("Fazla: " + ", ".join(fazla))
                if actual_chars and actual_chars != expected_chars:
                    msgs.append("Sıra Hatası")
                kural3 = "; ".join(msgs) if msgs else "Doğru"

                # KURAL5 – Test (2 referanslı formül limit kontrolü)
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
# Basit GUI (Tkinter)
# ==========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Formül Kontrol - KURAL1..KURAL5")
        self.geometry("680x360")

        # Değişkenler
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.uretim_yeri = tk.StringVar(value="3101")
        self.lower_col = tk.StringVar(value="LW_TOL_LMT")
        self.upper_col = tk.StringVar(value="UP_TOL_LMT")

        # UI
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        # Input
        ttk.Label(frm, text="Girdi (Excel):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.input_path, width=60).grid(row=0, column=1, padx=6)
        ttk.Button(frm, text="Seç...", command=self.select_input).grid(row=0, column=2)

        # Output
        ttk.Label(frm, text="Çıktı (Excel):").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.output_path, width=60).grid(row=1, column=1, padx=6)
        ttk.Button(frm, text="Kaydet...", command=self.select_output).grid(row=1, column=2)

        # Üretim yeri
        ttk.Label(frm, text="Üretim Yeri:").grid(row=2, column=0, sticky="w", pady=(8,0))
        uy_cb = ttk.Combobox(frm, textvariable=self.uretim_yeri, values=list(formuller.keys()), state="readonly", width=15)
        uy_cb.grid(row=2, column=1, sticky="w", pady=(8,0))

        # Limit kolon adları (KURAL5)
        ttk.Label(frm, text="Alt limit kolonu:").grid(row=3, column=0, sticky="w", pady=(8,0))
        ttk.Entry(frm, textvariable=self.lower_col, width=20).grid(row=3, column=1, sticky="w", pady=(8,0))

        ttk.Label(frm, text="Üst limit kolonu:").grid(row=4, column=0, sticky="w", pady=(4,0))
        ttk.Entry(frm, textvariable=self.upper_col, width=20).grid(row=4, column=1, sticky="w", pady=(4,0))

        # Çalıştır
        ttk.Button(frm, text="Kontrolü Başlat", command=self.run).grid(row=5, column=1, sticky="w", pady=16)

        # Not/Log
        self.log = tk.Text(frm, height=8, width=74)
        self.log.grid(row=6, column=0, columnspan=3, pady=(8,0))
        self.log.insert("end", "Not: Excel sütunları = PLAN_GROUP, OPER_NUM, OPER_DESC, INSPCHAR, MSTR_CHAR, FORMULA_FIELD_1 (+opsiyonel limit sütunları)\n")

        for i in range(3):
            frm.grid_columnconfigure(i, weight=1)

    def select_input(self):
        fp = filedialog.askopenfilename(
            title="Girdi Excel dosyasını seç",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if fp:
            self.input_path.set(fp)

    def select_output(self):
        fp = filedialog.asksaveasfilename(
            title="Çıktı Excel dosyası",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if fp:
            self.output_path.set(fp)

    def run(self):
        in_path = self.input_path.get().strip()
        out_path = self.output_path.get().strip()
        uy = self.uretim_yeri.get().strip()
        low_col = self.lower_col.get().strip() or "LW_TOL_LMT"
        up_col  = self.upper_col.get().strip() or "UP_TOL_LMT"

        if not in_path:
            messagebox.showerror("Hata", "Girdi dosyası seçilmedi.")
            return
        if not out_path:
            messagebox.showerror("Hata", "Çıktı dosya yolu seçilmedi.")
            return

        try:
            self.log.insert("end", f"Girdi okunuyor: {in_path}\n")
            df = pd.read_excel(in_path)

            # Minimum kolon kontrolü
            expected_cols = {"PLAN_GROUP", "OPER_NUM", "OPER_DESC", "INSPCHAR", "MSTR_CHAR", "FORMULA_FIELD_1"}
            missing = [c for c in expected_cols if c not in df.columns]
            if missing:
                messagebox.showerror("Hata", f"Excel kolonları eksik: {', '.join(missing)}")
                return

            self.log.insert("end", f"Üretim yeri: {uy} | Alt limit: {low_col} | Üst limit: {up_col}\n")
            out = kontrol_et(df, uy, lower_col=low_col, upper_col=up_col)
            out.to_excel(out_path, index=False)
            self.log.insert("end", f"Tamamlandı. Çıktı: {out_path}\n")
            messagebox.showinfo("Bitti", f"Kontrol tamamlandı.\nÇıktı: {out_path}")

        except Exception as e:
            messagebox.showerror("Hata", f"İşlem sırasında hata oluştu:\n{e}")
            self.log.insert("end", f"Hata: {e}\n")


if __name__ == "__main__":
    app = App()
    app.mainloop()
