import streamlit as st
import pandas as pd

def build_config_from_sidebar() -> dict:
    with st.sidebar.expander("نقاط القوة الشرائية (النسبة من القائد)", expanded=False):
        st.info("جدول النقاط يعتمد على حدود دنيا لكل مستوى.")
        pp_10 = st.number_input("حد أدنى % للحصول على 10 نقاط", value=50.0, step=0.1)
        pp_8  = st.number_input("حد أدنى % للحصول على 8 نقاط",  value=25.0, step=0.1)
        pp_7  = st.number_input("حد أدنى % للحصول على 7 نقاط",  value=15.0, step=0.1)
        pp_6  = st.number_input("حد أدنى % للحصول على 6 نقاط",  value=10.0, step=0.1)
        pp_5  = st.number_input("حد أدنى % للحصول على 5 نقاط",  value=5.0,  step=0.1)
        pp_4  = st.number_input("حد أدنى % للحصول على 4 نقاط",  value=4.0,  step=0.1)
        pp_3  = st.number_input("حد أدنى % للحصول على 3 نقاط",  value=3.0,  step=0.1)
        pp_2  = st.number_input("حد أدنى % للحصول على 2 نقاط",  value=2.0,  step=0.1)
        pp_1  = st.number_input("حد أدنى % للحصول على 1 نقطة",   value=1.0,  step=0.1)

    with st.sidebar.expander("نقاط الالتزام (عمر المديونية)", expanded=False):
        st.info("خصم تلقائي بالسالب لما بعد 60 يوم على شكل شرائح كل 30 يوم.")
        age_5 = st.number_input("≤ هذا العدد من الأيام = 5 نقاط", value=30, step=1)
        age_4 = st.number_input("≤ هذا العدد من الأيام = 4 نقاط", value=40, step=1)
        age_3 = st.number_input("≤ هذا العدد من الأيام = 3 نقاط", value=51, step=1)
        age_2 = st.number_input("≤ هذا العدد من الأيام = 2 نقاط", value=60, step=1)

    with st.sidebar.expander("نقاط المخاطرة (المديونية ÷ متوسط السداد الربعي)", expanded=False):
        st.info("يشمل نقاطًا سالبة إذا ارتفع المؤشر.")
        r_5 = st.number_input("≤ هذا المؤشر = 5 نقاط", value=1.00, step=0.1, format="%.2f")
        r_4 = st.number_input("≤ هذا المؤشر = 4 نقاط", value=1.50, step=0.1, format="%.2f")
        r_3 = st.number_input("≤ هذا المؤشر = 3 نقاط", value=2.00, step=0.1, format="%.2f")
        r_2 = st.number_input("≤ هذا المؤشر = 2 نقاط", value=2.50, step=0.1, format="%.2f")
        r_1 = st.number_input("≤ هذا المؤشر = 1 نقطة", value=3.00, step=0.1, format="%.2f")

    with st.sidebar.expander("إعدادات أعمدة الفارق المبسطة", expanded=False):
        snap_to_int = st.checkbox("تقريب فئة الفارق (نقاط) إلى أقرب عدد صحيح", value=True, key="snap_to_int")
        decimals_pct = st.number_input("عدد المنازل العشرية لفئة نسبة الفارق %", value=0, step=1, min_value=0, max_value=4, key="decimals_pct")

    with st.sidebar.expander("حدود تصنيف المرتجع (المضاعف مقابل المعيار)", expanded=False):
        m_ok     = st.number_input("≤ هذا المضاعف = ضمن المعيار", value=1.00, step=0.1, format="%.2f")
        m_watch  = st.number_input("≤ هذا المضاعف = يحتاج متابعة", value=1.50, step=0.1, format="%.2f")
        m_high   = st.number_input("≤ هذا المضاعف = مرتفع", value=2.00, step=0.1, format="%.2f")

    with st.sidebar.expander("التصنيف النهائي (حسب مجموع النقاط)", expanded=False):
        st.info("يشمل مستوى جديد: 8–9.9 = قبل النهاية.")
        final_motazem_min = st.number_input("≥ هذا المجموع = ملتزم", value=17.0, step=0.5)
        final_jayed_min   = st.number_input("≥ هذا المجموع = جيد", value=14.0, step=0.5)
        final_fix_cap_min = st.number_input("≥ هذا المجموع = جدولة + تثبيت السقف", value=12.0, step=0.1)
        final_reduce_min  = st.number_input("≥ هذا المجموع = جدولة + تخفيف", value=10.0, step=0.1)

    return {
        "pp": {"pp_10": pp_10, "pp_8": pp_8, "pp_7": pp_7, "pp_6": pp_6, "pp_5": pp_5, "pp_4": pp_4, "pp_3": pp_3, "pp_2": pp_2, "pp_1": pp_1},
        "age": {"age_5": age_5, "age_4": age_4, "age_3": age_3, "age_2": age_2},
        "risk": {"r_5": r_5, "r_4": r_4, "r_3": r_3, "r_2": r_2, "r_1": r_1},
        "delta": {"snap_to_int": snap_to_int, "decimals_pct": int(decimals_pct)},
        "returns": {"m_ok": m_ok, "m_watch": m_watch, "m_high": m_high},
        "final": {"motazem": final_motazem_min, "jayed": final_jayed_min, "fix_cap": final_fix_cap_min, "reduce": final_reduce_min},
    }

def score_purchase_power(pct, cfg_pp: dict) -> int:
    try:
        if pd.isna(pct): return 0
        x = float(pct)
    except Exception:
        return 0
    if x >= cfg_pp["pp_10"]: return 10
    elif x >= cfg_pp["pp_8"]: return 8
    elif x >= cfg_pp["pp_7"]: return 7
    elif x >= cfg_pp["pp_6"]: return 6
    elif x >= cfg_pp["pp_5"]: return 5
    elif x >= cfg_pp["pp_4"]: return 4
    elif x >= cfg_pp["pp_3"]: return 3
    elif x >= cfg_pp["pp_2"]: return 2
    elif x >= cfg_pp["pp_1"]: return 1
    else: return 0

def score_debt_age(days, cfg_age: dict):
    try:
        if pd.isna(days): return 5
        d = float(days)
    except Exception:
        return 5
    if d <= cfg_age["age_5"]: return 5
    elif d <= cfg_age["age_4"]: return 4
    elif d <= cfg_age["age_3"]: return 3
    elif d <= cfg_age["age_2"]: return 2
    else:
        extra_days = d - cfg_age["age_2"]
        penalty = extra_days / 30
        return -int(penalty) if float(penalty).is_integer() else -round(float(penalty), 2)

def score_risk(amount, avg_payment, cfg_risk: dict) -> int:
    try:
        a = float(amount) if not pd.isna(amount) else 0.0
        b = float(avg_payment) if not pd.isna(avg_payment) else 0.0
    except Exception:
        return 0
    if b == 0:
        return 0 if a == 0 else 0
    ratio = a / b
    if ratio <= cfg_risk["r_5"]: return 5
    elif ratio <= cfg_risk["r_4"]: return 4
    elif ratio <= cfg_risk["r_3"]: return 2
    elif ratio <= cfg_risk["r_2"]: return 1
    elif ratio <= cfg_risk["r_1"]: return 0
    elif ratio <= 4.0: return 0
    elif ratio <= 6.0: return -5
    elif ratio <= 12.0: return -10
    else: return 0

def final_classification(score, cfg_final: dict) -> str:
    if score >= cfg_final["motazem"]: return "ملتزم"
    elif score >= cfg_final["jayed"]: return "جيد"
    elif score >= cfg_final["fix_cap"]: return "جدوله مديونية وتثبيت السقف (حد أعلى المبيعات الآجل)"
    elif score >= cfg_final["reduce"]: return "جدوله مديونية وتخفيف المبيعات الآجل"
    elif score >= 8: return "قبل النهاية"
    else: return "عميل غير مجدي"
