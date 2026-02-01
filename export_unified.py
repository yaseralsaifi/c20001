import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

from .utils import to_numeric, clean_number

def compute_delta_table(base_df: pd.DataFrame, cols: dict, config: dict) -> pd.DataFrame:
    col_avgq = cols["avgq"]
    col_high = cols["high_avgq"]
    if not col_high:
        return pd.DataFrame()
    if col_avgq not in base_df.columns or col_high not in base_df.columns:
        return pd.DataFrame()

    avg_series  = to_numeric(base_df[col_avgq].map(clean_number))
    high_series = to_numeric(base_df[col_high].map(clean_number))
    decimals_pct = config["delta"]["decimals_pct"]

    d = pd.DataFrame(index=base_df.index)

    delta_ratio = np.where(
        (avg_series.notna()) & (avg_series != 0),
        (avg_series - high_series) / avg_series,
        np.nan
    )
    d["[فارق] فارق التغير (نسبي)"] = delta_ratio
    d["[فارق] فئة نسبة الفارق %"] = (pd.Series(delta_ratio, index=base_df.index).abs() * 100.0).round(decimals_pct)

    def _dir(x):
        if pd.isna(x): return "—"
        if x < 0: return "ارتفاع"
        if x > 0: return "انخفاض"
        return "مستقر"
    d["[فارق] اتجاه مبسط"] = pd.Series(delta_ratio, index=base_df.index).apply(_dir)

    def _mag(pct):
        if pd.isna(pct): return "—"
        if pct < 10: return "خفيف"
        elif pct < 30: return "متوسط"
        else: return "قوي"
    d["[فارق] شدة الفارق"] = d["[فارق] فئة نسبة الفارق %"].apply(_mag)

    return d

def compute_returns_table(base_df: pd.DataFrame, cols: dict, config: dict) -> pd.DataFrame:
    col_avgpay = cols["avgq"]
    col_base   = "نسبة المرتجع من المباع"
    col_new    = "نسبة نوع جديد من مرتجعات العميل"
    col_comp   = "نسبة نوع تعويض من مرتجعات العميل"

    m_ok = config["returns"]["m_ok"]
    m_watch = config["returns"]["m_watch"]
    m_high = config["returns"]["m_high"]

    def _score_pp(p):
        if pd.isna(p): return np.nan
        if p >= 50: return 10
        elif p >= 25: return 8
        elif p >= 15: return 7
        elif p >= 10: return 6
        elif p >= 5:  return 5
        elif p >= 4:  return 4
        elif p >= 3:  return 3
        elif p >= 2:  return 2
        elif p >= 1:  return 1
        else:        return 0

    def _classify(rate, ref):
        if pd.isna(rate) or pd.isna(ref) or ref == 0: return "بيانات غير كافية"
        ratio = rate / ref
        if ratio <= m_ok: return "ضمن المعيار"
        elif ratio <= m_watch: return "يحتاج متابعة"
        elif ratio <= m_high: return "مرتفع"
        else: return "مرتفع جدًا"

    if col_avgpay not in base_df.columns:
        return pd.DataFrame()

    avg_series = to_numeric(base_df[col_avgpay].map(clean_number))
    max_avg = avg_series.max()
    pct_avg = np.where(max_avg > 0, (avg_series / max_avg * 100).round(2), np.nan)
    pp = pd.Series(pct_avg).apply(_score_pp)

    def _one(col_name: str, label: str) -> pd.DataFrame:
        if col_name not in base_df.columns:
            return pd.DataFrame()
        tmp = to_numeric(base_df[col_name].map(clean_number)).replace([np.inf, -np.inf], np.nan)
        ref_vals = tmp[pp.between(5, 10, inclusive="both")]
        ref_avg = ref_vals.mean()

        out = pd.DataFrame(index=base_df.index)
        out[f"[مرتجع] معيار ({label} 10–5)"] = round(ref_avg, 4) if pd.notna(ref_avg) else np.nan
        ratio = tmp / ref_avg if (ref_avg and ref_avg != 0) else np.nan
        out[f"[مرتجع] مضاعف ({label}) مقابل المعيار"] = ratio
        out[f"[مرتجع] تصنيف ({label})"] = tmp.apply(lambda x: _classify(x, ref_avg))
        out[f"[مرتجع] قيمة ({label})"] = tmp
        return out

    parts = [
        _one(col_base, "المرتجع من المباع"),
        _one(col_new, "النوع الجديد"),
        _one(col_comp, "نوع تعويض"),
    ]
    parts = [p for p in parts if not p.empty]
    if not parts:
        return pd.DataFrame()
    return pd.concat(parts, axis=1)

def render_unified_export(df, df_original, cols, config):
    with st.expander("📦 تصدير ملف Excel موحّد (جميع الأعمدة والنتائج)", expanded=False):
        st.write("ينشئ ملفًا واحدًا يجمع: أعمدة الملف الأصلي + نتائج التبويب الأساسي + أعمدة الفارق + تصنيفات المرتجع.")

        df_delta_all = compute_delta_table(df, cols, config)
        df_returns_all = compute_returns_table(df, cols, config)

        unified = df_original.copy()
        main_extra_cols = [c for c in df.columns if c not in df_original.columns]
        unified = unified.join(df[main_extra_cols].add_prefix("[أساسي] "))

        if df_delta_all is not None and not df_delta_all.empty:
            unified = unified.join(df_delta_all)
        if df_returns_all is not None and not df_returns_all.empty:
            unified = unified.join(df_returns_all)

        st.dataframe(unified, use_container_width=True)

        buf = BytesIO()
        unified.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button("⬇️ تحميل الملف الموحّد (Excel)", buf, file_name="نتائج_موحّدة_كل_التبويبات_v5_6_5.xlsx")
