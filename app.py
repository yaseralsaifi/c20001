
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="المساعد الذكي لتصنيف العملاء - v5.6.3.2", layout="wide")
st.title("المساعد الذكي لتصنيف العملاء وتحليل المديونية — v5.6.3.2")

uploaded_file = st.file_uploader("📂 ارفع ملف Excel أو CSV", type=["xlsx", "csv"])

st.sidebar.header("⚙️ الإعدادات أعمدة الفارق المبسطة")
snap_to_int = st.sidebar.checkbox("تقريب فئة الفارق (نقاط) إلى أقرب عدد صحيح", value=True)
decimals_pct = st.sidebar.number_input("عدد المنازل العشرية لفئة نسبة الفارق %", value=0, step=1, min_value=0, max_value=4)

def normalize(name: str) -> str:
    return str(name).strip().replace("\\u200f","").replace("\\u200e","")

if uploaded_file:
    # قراءة الملف
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        st.stop()

    # تطبيع أسماء الأعمدة
    df.columns = [normalize(c) for c in df.columns]

    # التحقق من الأعمدة المطلوبة
    need_a = "متوسط السداد الربعي"
    need_b = "أعلى متوسط السداد الربعي"
    if need_a not in df.columns or need_b not in df.columns:
        st.error(f"الملف يجب أن يحتوي الأعمدة: '{need_a}', '{need_b}'. الأعمدة الحالية: {list(df.columns)}")
        st.stop()

    # تحويل القيم إلى رقمية وتجاهل غير الرقمية
    avg_series  = pd.to_numeric(df[need_a], errors="coerce")
    high_series = pd.to_numeric(df[need_b], errors="coerce")

    max_avg  = avg_series.max()
    max_high = high_series.max()
    if pd.isna(max_avg) or max_avg == 0 or pd.isna(max_high) or max_high == 0:
        st.error("لا يمكن الحساب لأن أحد الأعمدة المرجعية يساوي صفر/فارغ.")
        st.stop()

    # حساب نسب من القائد
    df["نسبة من القائد (متوسط)"] = (avg_series / max_avg * 100).round(2)
    df["نسبة من القائد (أعلى)"]   = (high_series / max_high * 100).round(2)

    # الفارق بالنقاط
    delta_points = (df["نسبة من القائد (أعلى)"] - df["نسبة من القائد (متوسط)"]).astype(float)
    abs_points = delta_points.abs()

    if snap_to_int:
        df["فئة الفارق (نقاط)"] = abs_points.fillna(0).replace([np.inf, -np.inf], 0).round(0).astype(int)
    else:
        df["فئة الفارق (نقاط)"] = abs_points.replace([np.inf, -np.inf], np.nan).round(2)

    # فئة نسبة الفارق % نسبة إلى أساس العميل (المتوسط)
    def pct_change(base, other):
        if pd.isna(base) or base == 0:
            return np.nan
        return abs(other - base) / abs(base) * 100.0

    df["فئة نسبة الفارق %"] = [
        pct_change(b, o) for b, o in zip(df["نسبة من القائد (متوسط)"], df["نسبة من القائد (أعلى)"])
    ]
    df["فئة نسبة الفارق %"] = df["فئة نسبة الفارق %"].round(int(decimals_pct))

    # اتجاه مبسط
    def simple_dir(x):
        if pd.isna(x): return "—"
        if x > 0: return "ارتفاع"
        if x < 0: return "انخفاض"
        return "مستقر"

    df["اتجاه مبسط"] = delta_points.apply(simple_dir)

    st.success("✅ تم إنشاء الأعمدة الثلاثة المبسطة بدون أخطاء تبويب.")
    st.dataframe(df[["نسبة من القائد (متوسط)","نسبة من القائد (أعلى)","فئة الفارق (نقاط)","فئة نسبة الفارق %","اتجاه مبسط"]], use_container_width=True)

    # حفظ الملف الناتج
    out_path = "نتائج_التصنيف_مع_الأعمدة_المبسطة_v5_6_3_2.xlsx"
    df.to_excel(out_path, index=False)
    with open(out_path, "rb") as f_out:
        st.download_button("⬇️ تحميل الملف الناتج (Excel)", f_out, file_name=out_path)
else:
    st.info("⬆️ ارفع ملف العملاء ثم سيتم إنشاء الأعمدة الثلاثة المطلوبة.")
