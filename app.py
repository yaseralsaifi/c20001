# -*- coding: utf-8 -*-
# Streamlit unified app (Arabic) — v5.6.5 (vectorized returns, smart thresholds, robust numbers)
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ======================== إعدادات الصفحة ========================
st.set_page_config(page_title="المساعد الذكي لتصنيف العملاء - v5.6.5", layout="wide")
st.title("المساعد الذكي لتصنيف العملاء وتحليل المديونية — v5.6.5")

# ======================== وظائف مساعدة عامة ========================
def normalize(name: str) -> str:
    # إزالة محارف الاتجاه الخفية
    return str(name).strip().replace("\u200f", "").replace("\u200e", "")

def to_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def clean_number(x):
    """تحويل القيم النصية إلى رقم: أرقام عربية، فواصل عربية/إنجليزية، وإزالة %."""
    if pd.isna(x):
        return x
    try:
        s = str(x).strip()
        trans = {
            ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3',
            ord('٤'): '4', ord('٥'): '5', ord('٦'): '6', ord('٧'): '7',
            ord('٨'): '8', ord('٩'): '9',
            ord('٬'): ',', ord('،'): ',', ord('٫'): '.',  # فاصل أعشار عربي → نقطة
        }
        s = s.translate(trans)
        s = s.replace('%', '')
        # إزالة فواصل الآلاف
        s = s.replace(',', '')
        return s
    except Exception:
        return x

def norm_key(s: str) -> str:
    """تطبيع قوي للأسماء العربية لتسهيل المطابقة."""
    s = normalize(s)
    trans = {
        ord('آ'): 'ا', ord('أ'): 'ا', ord('إ'): 'ا',
        ord('ى'): 'ي', ord('ة'): 'ه', ord('ؤ'): 'و', ord('ئ'): 'ي',
        ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3',
        ord('٤'): '4', ord('٥'): '5', ord('٦'): '6', ord('٧'): '7',
        ord('٨'): '8', ord('٩'): '9',
        ord('٬'): ',', ord('،'): ',', ord('٫'): ',',
        ord('ـ'): '',  # حذف التطويل
    }
    s = s.translate(trans).lower()
    return ''.join(ch for ch in s if ch.isalnum())

def match_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """يرجع اسم العمود الأصلي إذا طابق أي اسم في القائمة بعد التطبيع."""
    keys = {norm_key(c): c for c in df.columns}
    for a in aliases:
        k = norm_key(a)
        if k in keys:
            return keys[k]
    return None

# ======================== الشريط الجانبي ========================
st.sidebar.header("📂 ملف البيانات")
uploaded_file = st.sidebar.file_uploader("ارفع ملف Excel أو CSV", type=["xlsx", "csv"]) 

st.sidebar.header("⚙️ الإعدادات الأساسية")

# ---- نقاط القوة الشرائية ----
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

    def score_purchase_power(pct):
        try:
            if pd.isna(pct): return 0
            x = float(pct)
        except Exception:
            return 0
        if x >= pp_10: return 10
        elif x >= pp_8: return 8
        elif x >= pp_7: return 7
        elif x >= pp_6: return 6
        elif x >= pp_5: return 5
        elif x >= pp_4: return 4
        elif x >= pp_3: return 3
        elif x >= pp_2: return 2
        elif x >= pp_1: return 1
        else: return 0

# ---- نقاط الالتزام (عمر المديونية) ----
with st.sidebar.expander("نقاط الالتزام (عمر المديونية)", expanded=False):
    st.info("خصم تلقائي بالسالب لما بعد 60 يوم على شكل شرائح كل 30 يوم.")
    age_5 = st.number_input("≤ هذا العدد من الأيام = 5 نقاط", value=30, step=1)
    age_4 = st.number_input("≤ هذا العدد من الأيام = 4 نقاط", value=40, step=1)
    age_3 = st.number_input("≤ هذا العدد من الأيام = 3 نقاط", value=51, step=1)
    age_2 = st.number_input("≤ هذا العدد من الأيام = 2 نقاط", value=60, step=1)

    def score_debt_age(days):
        try:
            if pd.isna(days): return 0
            d = float(days)
        except Exception:
            return 0
        if d <= age_5: return 5
        elif d <= age_4: return 4
        elif d <= age_3: return 3
        elif d <= age_2: return 2
        else:
            extra_days = d - age_2
            penalty = extra_days / 30
            return -int(penalty) if float(penalty).is_integer() else -round(float(penalty), 2)

# ---- نقاط المخاطرة (المديونية ÷ متوسط السداد الربعي) ----
with st.sidebar.expander("نقاط المخاطرة (المديونية ÷ متوسط السداد الربعي)", expanded=False):
    st.info("يشمل نقاطًا سالبة إذا ارتفع المؤشر.")
    r_5 = st.number_input("≤ هذا المؤشر = 5 نقاط", value=1.00, step=0.1, format="%.2f")
    r_4 = st.number_input("≤ هذا المؤشر = 4 نقاط", value=1.50, step=0.1, format="%.2f")
    r_3 = st.number_input("≤ هذا المؤشر = 3 نقاط", value=2.00, step=0.1, format="%.2f")
    r_2 = st.number_input("≤ هذا المؤشر = 2 نقاط", value=2.50, step=0.1, format="%.2f")
    r_1 = st.number_input("≤ هذا المؤشر = 1 نقطة", value=3.00, step=0.1, format="%.2f")

    def score_risk(amount, avg_payment):
        # تطبيق الشرائح + معالجة متوسط = 0
        try:
            a = float(amount) if not pd.isna(amount) else 0.0
            b = float(avg_payment) if not pd.isna(avg_payment) else 0.0
        except Exception:
            return 0
        if b == 0:
            return 0 if a == 0 else 5
        ratio = a / b
        if ratio <= r_5: return 5
        elif ratio <= r_4: return 4
        elif ratio <= r_3: return 2
        elif ratio <= r_2: return 1
        elif ratio <= r_1: return 0
        elif ratio <= 4.0: return 0
        elif ratio <= 6.0: return 0
        elif ratio <= 12.0: return 0
        else: return 5

# ---- إعدادات الفارق ----
with st.sidebar.expander("إعدادات أعمدة الفارق المبسطة", expanded=False):
    snap_to_int = st.checkbox("تقريب فئة الفارق (نقاط) إلى أقرب عدد صحيح", value=True, key="snap_to_int")
    decimals_pct = st.number_input("عدد المنازل العشرية لفئة نسبة الفارق %", value=0, step=1, min_value=0, max_value=4, key="decimals_pct")

# ---- حدود تصنيف المرتجع ----
with st.sidebar.expander("حدود تصنيف المرتجع (المضاعف مقابل المعيار)", expanded=False):
    m_ok     = st.number_input("≤ هذا المضاعف = ضمن المعيار", value=1.00, step=0.1, format="%.2f")
    m_watch  = st.number_input("≤ هذا المضاعف = يحتاج متابعة", value=1.50, step=0.1, format="%.2f")
    m_high   = st.number_input("≤ هذا المضاعف = مرتفع", value=2.00, step=0.1, format="%.2f")

# ======================== قراءة الملف ========================
if not uploaded_file:
    st.info("⬆️ ارفع ملف العملاء للبدء (Excel/CSV).")
    st.stop()

try:
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"تعذر قراءة الملف: {e}")
    st.stop()

# توحيد أسماء الأعمدة
df.columns = [normalize(c) for c in df.columns]

# أسماء بديلة شائعة
debt_aliases = ["المديونية","رصيد المديونية","إجمالي المديونية","اجمالي المديونية","رصيد","مستحقات","رصيد مستحق"]
avgq_aliases = ["متوسط السداد الربعي","متوسط السداد","متوسط السداد 3 اشهر","متوسط السداد ٣ اشهر","متوسط السداد الشهري","المتوسط الشهري للسداد"]
age_aliases  = ["عمر المديونية (يوم)","عمر المديونية","أيام المديونية","ايام المديونية","عمر الدين","عدد الايام"]
high_avg_aliases = ["أعلى متوسط السداد الربعي","اعلى متوسط السداد الربعي","أقصى متوسط السداد الربعي","اقصى متوسط السداد"]

# اكتشاف تلقائي
detected_debt = match_col(df, debt_aliases)
detected_avgq = match_col(df, avgq_aliases)
detected_age  = match_col(df, age_aliases)
detected_high = match_col(df, high_avg_aliases)

# اختيار يدوي
st.sidebar.subheader("🧭 تعيين الأعمدة يدويًا (إن لزم)")
col_debt = st.sidebar.selectbox("عمود المديونية", df.columns, index=(df.columns.get_loc(detected_debt) if detected_debt else 0))
col_avgq = st.sidebar.selectbox("عمود متوسط السداد الربعي", df.columns, index=(df.columns.get_loc(detected_avgq) if detected_avgq else 0))
col_age_opt = ["— (بدون) —"] + list(df.columns)
col_age = st.sidebar.selectbox("عمود عمر المديونية (اختياري)", col_age_opt, index=(1 + df.columns.get_loc(detected_age)) if detected_age else 0)
col_high_avgq = st.sidebar.selectbox("عمود أعلى متوسط السداد الربعي (للفارق)", ["— (غير مستخدم) —"] + list(df.columns), index=(1 + df.columns.get_loc(detected_high)) if detected_high else 0)
if col_age == "— (بدون) —": col_age = None
if col_high_avgq == "— (غير مستخدم) —": col_high_avgq = None

st.sidebar.caption(f"Detected ➜ المديونية: {detected_debt or '—'} | المتوسط: {detected_avgq or '—'} | العمر: {detected_age or '—'} | أعلى متوسط: {detected_high or '—'}")
st.info("سيتم استخدام الأعمدة: المديونية = **{}** ، متوسط السداد = **{}**{}".format(
    col_debt, col_avgq, f" ، العمر = **{col_age}**" if col_age else ""))

# نسخة أصلية
df_original = df.copy()

# ======================== تبويبات الواجهة ========================
tab_main, tab_delta, tab_returns, tab_diag = st.tabs([
    "🔎 التصنيف والتحليل الأساسي",
    "🔁 أعمدة الفارق المبسطة",
    "📊 تصنيفات المرتجع (مستقل)",
    "🛠️ تشخيص سريع"
])

# ======================== التبويب 1: التصنيف والتحليل ========================
with tab_main:
    st.subheader("🔎 التصنيف والتحليل الأساسي")

    missing = [c for c in [col_debt, col_avgq] if c not in df.columns]
    if missing:
        st.error(f"يجب توافر أعمدة: {missing}")
        st.stop()

    # تنظيف القيم المختارة ثم تحويلها لرقمية
    for c in [col_debt, col_avgq] + ([col_age] if col_age else []):
        df[c] = to_numeric(df[c].map(clean_number))

    # نسبة من القائد (متوسط السداد)
    max_avg = df[col_avgq].max()
    df["نسبة من القائد (متوسط)"] = np.where(max_avg > 0, (df[col_avgq] / max_avg * 100).round(2), np.nan)

    # نقاط القوة الشرائية / الالتزام / المخاطرة
    df["نقاط القوة الشرائية"] = df["نسبة من القائد (متوسط)"].apply(score_purchase_power)
    df["نقاط الالتزام"] = df[col_age].apply(score_debt_age) if col_age and col_age in df.columns else 0

    # مؤشر المخاطرة + النقاط
    def safe_ratio(a, b):
        try:
            a = float(a) if not pd.isna(a) else 0.0
            b = float(b) if not pd.isna(b) else 0.0
            return a / b if b != 0 else np.nan
        except Exception:
            return np.nan
    df["مؤشر المخاطرة (مديونية/متوسط)"] = df.apply(lambda r: safe_ratio(r[col_debt], r[col_avgq]), axis=1).round(3)
    df["نقاط المخاطرة"] = df.apply(lambda r: score_risk(r[col_debt], r[col_avgq]), axis=1)

    # ---- التصنيف النهائي ----
    with st.sidebar.expander("التصنيف النهائي (حسب مجموع النقاط)", expanded=False):
        st.info("يشمل مستوى جديد: 8–9.9 = قبل النهاية.")
        final_motazem_min = st.number_input("≥ هذا المجموع = ملتزم", value=17.0, step=0.5)
        final_jayed_min   = st.number_input("≥ هذا المجموع = جيد", value=14.0, step=0.5)
        final_fix_cap_min = st.number_input("≥ هذا المجموع = جدولة + تثبيت السقف", value=12.0, step=0.1)
        final_reduce_min  = st.number_input("≥ هذا المجموع = جدولة + تخفيف", value=10.0, step=0.1)
        def final_classification(score):
            if score >= final_motazem_min: return "ملتزم"
            elif score >= final_jayed_min: return "جيد"
            elif score >= final_fix_cap_min: return "جدوله مديونية وتثبيت السقف (حد أعلى المبيعات الآجل)"
            elif score >= final_reduce_min: return "جدوله مديونية وتخفيف المبيعات الآجل"
            elif score >= 8: return "قبل النهاية"
            else: return "عميل غير مجدي"

    df["إجمالي النقاط"] = df[["نقاط القوة الشرائية", "نقاط الالتزام", "نقاط المخاطرة"]].sum(axis=1)
    df["التصنيف النهائي"] = df["إجمالي النقاط"].apply(final_classification)

    # ===== نسب المندوب =====
    rep_col_candidates = ["اسم المندوب", "المندوب", "مندوب", "اسم مندوب"]
    rep_col = next((c for c in rep_col_candidates if c in df.columns), None)
    if rep_col is not None and col_debt in df.columns:
        cnt = df.groupby([rep_col, "التصنيف النهائي"])["إجمالي النقاط"].size().rename("عدد")
        cnt_by_class = cnt.groupby(level=1).transform("sum")
        share_count = (cnt / cnt_by_class * 100).round(2)

        debt_grp = df.groupby([rep_col, "التصنيف النهائي"])[col_debt].sum().rename("مديونية")
        debt_by_class = debt_grp.groupby(level=1).transform("sum")
        share_debt = (debt_grp / debt_by_class * 100).round(2)

        share_count_map = share_count.to_dict()
        share_debt_map = share_debt.to_dict()
        df["نسبة المندوب من فئة العميل (بالعدد %)"] = df.apply(
            lambda r: share_count_map.get((r[rep_col], r["التصنيف النهائي"]), np.nan), axis=1
        )
        df["نسبة المندوب من فئة العميل (بالمديونية %)"] = df.apply(
            lambda r: share_debt_map.get((r[rep_col], r["التصنيف النهائي"]), np.nan), axis=1
        )

        df["إجمالي مديونية المندوب"] = df.groupby(rep_col)[col_debt].transform("sum")
        df["مديونية المندوب ضمن هذه الفئة"] = df.groupby([rep_col, "التصنيف النهائي"])[col_debt].transform("sum")
        df["نسبة الفئة داخل مديونية المندوب (%)"] = np.where(
            df["إجمالي مديونية المندوب"] > 0,
            (df["مديونية المندوب ضمن هذه الفئة"] / df["إجمالي مديونية المندوب"]) * 100,
            np.nan
        ).round(2)
    else:
        df["نسبة المندوب من فئة العميل (بالعدد %)"] = np.nan
        df["نسبة المندوب من فئة العميل (بالمديونية %)"] = np.nan
        df["نسبة الفئة داخل مديونية المندوب (%)"] = np.nan

    # نسبة كل تصنيف من إجمالي المديونية
    total_debt = df[col_debt].sum(skipna=True)
    if total_debt and total_debt != 0:
        class_debt = df.groupby("التصنيف النهائي")[col_debt].sum()
        share_map = (class_debt / total_debt * 100).to_dict()
        df["نسبة التصنيف من إجمالي المديونية (%)"] = df["التصنيف النهائي"].map(share_map).round(2)
    else:
        df["نسبة التصنيف من إجمالي المديونية (%)"] = 0.0

    # ===== خطة المعالجة الذكية =====
    st.markdown("### 🧠 خطة المعالجة الذكية")
    df["مبلغ الانحراف (للـ3 أشهر)"] = np.maximum(0.0, df[col_debt].fillna(0).astype(float) - 3.0 * df[col_avgq].fillna(0).astype(float)).round(2)
    df["قسط الانحراف الشهري"] = (df["مبلغ الانحراف (للـ3 أشهر)"] / 3.0).round(2)
    df["فقد_نقاط_التزام/مخاطرة؟"] = (df["نقاط الالتزام"] < 5) | (df["نقاط المخاطرة"] < 5)

    def base_targets(class_name, pwr_bucket, avg_q, debt):
        avg_q = float(avg_q or 0); debt = float(debt or 0); pwr_bucket = float(pwr_bucket or 0)
        if pwr_bucket < 5: return debt, avg_q
        if class_name == "ملتزم":                pay, sales = avg_q, avg_q
        elif class_name == "جيد":                 pay, sales = avg_q * 1.10, avg_q
        elif str(class_name).startswith("جدوله مديونية وتثبيت"): pay, sales = avg_q * 1.15, avg_q
        elif str(class_name).startswith("جدوله مديونية وتخفيف"):  pay, sales = avg_q * 1.15, avg_q * 0.90
        elif class_name == "قبل النهاية":        pay, sales = avg_q * 1.15, avg_q * 0.85
        else:                                     pay, sales = 0.0, 0.0
        return float(round(pay, 2)), float(round(sales, 2))

    pay_base_list, sales_base_list, pay_final_list = [], [], []
    for _, r in df.iterrows():
        pay_b, sales_b = base_targets(r["التصنيف النهائي"], r["نقاط القوة الشرائية"], r[col_avgq], r[col_debt])
        pay_base_list.append(pay_b); sales_base_list.append(sales_b)
        pay_final_list.append(round(pay_b + (float(r["قسط الانحراف الشهري"]) if bool(r["فقد_نقاط_التزام/مخاطرة؟"]) else 0.0), 2))

    df["هدف السداد الشهري (أساس)"] = pay_base_list
    df["هدف المبيعات الشهري"] = sales_base_list
    df["هدف السداد الشهري (بعد المعالجة)"] = pay_final_list
    df["ملاحظة خطة السداد"] = np.where(df["فقد_نقاط_التزام/مخاطرة؟"],
                                        "تفعيل الخطة: تمت إضافة قسط الانحراف الشهري",
                                        "لا توجد خسارة نقاط في الالتزام/المخاطرة — الاكتفاء بالهدف الأساسي")

    st.success("✅ تم إعداد التصنيف والخطة بنجاح.")
    st.dataframe(df, use_container_width=True)

    out_main = BytesIO()
    df.to_excel(out_main, index=False); out_main.seek(0)
    st.download_button("⬇️ تحميل الملف الناتج (Excel)", out_main, file_name="نتائج_التصنيف_v5_6_5.xlsx")

# ======================== التبويب 2: أعمدة الفارق المبسطة ========================
with tab_delta:
    st.subheader("🔁 أعمدة الفارق المبسطة")
    need_a = col_avgq; need_b = col_high_avgq
    if not need_b:
        st.info("للحساب هنا يلزم اختيار عمود 'أعلى متوسط السداد الربعي' من الشريط الجانبي.")
    elif need_a in df.columns and need_b in df.columns:
        avg_series  = to_numeric(df[need_a].map(clean_number))
        high_series = to_numeric(df[need_b].map(clean_number))
        max_avg  = avg_series.max(); max_high = high_series.max()
        if pd.isna(max_avg) or max_avg == 0 or pd.isna(max_high) or max_high == 0:
            st.error("لا يمكن الحساب لأن أحد الأعمدة المرجعية يساوي صفر/فارغ.")
        else:
            df_delta = df.copy()
            df_delta["نسبة من القائد (متوسط)"] = (avg_series / max_avg * 100).round(2)
            df_delta["نسبة من القائد (أعلى)"]   = (high_series / max_high * 100).round(2)
            delta_points = (df_delta["نسبة من القائد (أعلى)"] - df_delta["نسبة من القائد (متوسط)"]).astype(float)
            abs_points = delta_points.abs()
            if snap_to_int:
                df_delta["فئة الفارق (نقاط)"] = abs_points.fillna(0).replace([np.inf, -np.inf], 0).round(0).astype(int)
            else:
                df_delta["فئة الفارق (نقاط)"] = abs_points.replace([np.inf, -np.inf], np.nan).round(int(decimals_pct))
            def pct_change(base, other):
                if pd.isna(base) or base == 0: return np.nan
                return abs(other - base) / abs(base) * 100.0
            df_delta["فئة نسبة الفارق %"] = [
                pct_change(b, o) for b, o in zip(df_delta["نسبة من القائد (متوسط)"], df_delta["نسبة من القائد (أعلى)"])
            ]
            df_delta["فئة نسبة الفارق %"] = df_delta["فئة نسبة الفارق %"].round(int(decimals_pct))
            def simple_dir(x):
                if pd.isna(x): return "—"
                if x > 0: return "ارتفاع"
                if x < 0: return "انخفاض"
                return "مستقر"
            df_delta["اتجاه مبسط"] = delta_points.apply(simple_dir)
            st.dataframe(df_delta[["نسبة من القائد (متوسط)","نسبة من القائد (أعلى)","فئة الفارق (نقاط)","فئة نسبة الفارق %","اتجاه مبسط"]], use_container_width=True)
            out_delta = BytesIO(); df_delta.to_excel(out_delta, index=False); out_delta.seek(0)
            st.download_button("⬇️ تحميل ملف الفارق (Excel)", out_delta, file_name="نتائج_أعمدة_الفارق_v5_6_5.xlsx")
    else:
        st.info("للحساب هنا يلزم وجود الأعمدة المختارة في الملف.")

# ======================== التبويب 3: تصنيفات المرتجع (مستقل) ========================
with tab_returns:
    st.subheader("📊 تصنيفات المرتجع — مستقل (v5.6.5)")
    col_avgpay = st.text_input("اسم عمود (متوسط السداد الربعي)", value=col_avgq or "متوسط السداد الربعي")
    col_base   = st.text_input("اسم عمود (نسبة المرتجع من المباع)", value="نسبة المرتجع من المباع")
    col_new    = st.text_input("اسم عمود (نسبة نوع جديد من مرتجعات العميل)", value="نسبة نوع جديد من مرتجعات العميل")
    col_comp   = st.text_input("اسم عمود (نسبة نوع تعويض من مرتجعات العميل)", value="نسبة نوع تعويض من مرتجعات العميل")

    def score_purchase_power_for_returns(p):
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

    def process_return_column(df_in: pd.DataFrame, col_name: str, label: str):
        # تحقق من الأعمدة
        if col_name not in df_in.columns:
            st.warning(f"⚠️ لم يتم العثور على العمود: {col_name}")
            return None
        if col_avgpay not in df_in.columns:
            st.warning(f"لا يمكن احتساب مجموعة المرجع لعدم توفر '{col_avgpay}'.")
            return None

        # 1) تحويل الأرقام بشكل متّجه
        avg_series = pd.to_numeric(df_in[col_avgpay].map(clean_number), errors="coerce")
        vals       = pd.to_numeric(df_in[col_name].map(clean_number),   errors="coerce")

        # 2) مجموعة المرجع (القوة الشرائية 10–5)
        max_avg = avg_series.max(skipna=True)
        pct = np.where(max_avg > 0, (avg_series / max_avg * 100), np.nan)
        pp  = pd.Series(pct, index=df_in.index).apply(score_purchase_power_for_returns)
        mask_ref = pp.between(5, 10, inclusive="both")

        # 3) معيار المرتجع على مجموعة المرجع
        ref_avg = vals[mask_ref].mean(skipna=True)

        # 4) بناء الناتج
        out = pd.DataFrame(index=df_in.index)
        out[col_name] = vals
        out[f"معيار المرتجع ({label} 10–5)"] = ref_avg

        # 5) إن لم يوجد معيار صالح
        if pd.isna(ref_avg) or ref_avg == 0:
            out[f"مضاعف المرتجع ({label}) مقابل المعيار"] = np.nan
            out[f"تصنيف المرتجع ({label})"] = "بيانات غير كافية"
            return out

        # 6) حساب المضاعف + التصنيف بشكل متّجه
        ratio = vals / ref_avg
        out[f"مضاعف المرتجع ({label}) مقابل المعيار"] = ratio

        def label_ratio(x):
            if pd.isna(x): return "بيانات غير كافية"
            if x <= m_ok:      return "ضمن المعيار"
            elif x <= m_watch: return "يحتاج متابعة"
            elif x <= m_high:  return "مرتفع"
            else:              return "مرتفع جدًا"

        out[f"تصنيف المرتجع ({label})"] = ratio.apply(label_ratio)
        return out

    sections = []
    for cname, lbl in [(col_base, "المرتجع من المباع"), (col_new, "النوع الجديد"), (col_comp, "نوع تعويض")]:
        res = process_return_column(df, cname, lbl)
        if res is not None:
            st.subheader(f"🔎 {lbl}")
            st.dataframe(res, use_container_width=True)
            sections.append(res)
    if sections:
        out_ret = BytesIO(); out_df = pd.concat(sections, axis=1)
        out_df.to_excel(out_ret, index=False); out_ret.seek(0)
        st.download_button("⬇️ تحميل الملف (تصنيفات المرتجع)", out_ret, file_name="نتائج_تصنيفات_المرتجع_v5_6_5.xlsx")

# ======================== التبويب 4: تشخيص سريع ========================
with tab_diag:
    st.subheader("🛠️ تشخيص أسباب (0) في نقاط المخاطرة")
    zero_avg = df[col_avgq] == 0
    sample = df.loc[zero_avg, [col_debt, col_avgq]].head(10)
    st.write(f"عدد العملاء الذين متوسط سدادهم = 0: **{zero_avg.sum()}**")
    if not sample.empty:
        st.dataframe(sample, use_container_width=True)
    st.caption("إذا كان متوسط السداد = 0 والمديونية > 0 فسنعطي -10 نقاط تلقائيًا.")

# ======================== تصدير ملف موحّد ========================
with st.expander("📦 تصدير ملف Excel موحّد (جميع الأعمدة والنتائج)", expanded=False):
    st.write("ينشئ ملفًا واحدًا يجمع: أعمدة الملف الأصلي + نتائج التبويب الأساسي + أعمدة الفارق + تصنيفات المرتجع.")

    def compute_delta_table(base_df: pd.DataFrame) -> pd.DataFrame:
        if not col_high_avgq: return pd.DataFrame()
        need_a, need_b = col_avgq, col_high_avgq
        if need_a not in base_df.columns or need_b not in base_df.columns: return pd.DataFrame()
        avg_series  = to_numeric(base_df[need_a].map(clean_number))
        high_series = to_numeric(base_df[need_b].map(clean_number))
        max_avg  = avg_series.max(); max_high = high_series.max()
        if pd.isna(max_avg) or max_avg == 0 or pd.isna(max_high) or max_high == 0: return pd.DataFrame()
        d = pd.DataFrame(index=base_df.index)
        d["[فارق] نسبة من القائد (متوسط)"] = (avg_series / max_avg * 100).round(2)
        d["[فارق] نسبة من القائد (أعلى)"]   = (high_series / max_high * 100).round(2)
        delta_points = (d["[فارق] نسبة من القائد (أعلى)"] - d["[فارق] نسبة من القائد (متوسط)"]).astype(float)
        abs_points = delta_points.abs()
        d["[فارق] فئة الفارق (نقاط)"] = (abs_points.fillna(0).replace([np.inf, -np.inf], 0).round(0).astype(int)
                                          if snap_to_int else abs_points.replace([np.inf, -np.inf], np.nan).round(int(decimals_pct)))
        def pct_change(base, other):
            if pd.isna(base) or base == 0: return np.nan
            return abs(other - base) / abs(base) * 100.0
        d["[فارق] فئة نسبة الفارق %"] = [pct_change(b, o) for b, o in zip(d["[فارق] نسبة من القائد (متوسط)"], d["[فارق] نسبة من القائد (أعلى)"])]
        d["[فارق] فئة نسبة الفارق %"] = d["[فارق] فئة نسبة الفارق %"].round(int(decimals_pct))
        def simple_dir(x):
            if pd.isna(x): return "—"
            if x > 0: return "ارتفاع"
            if x < 0: return "انخفاض"
            return "مستقر"
        d["[فارق] اتجاه مبسط"] = delta_points.apply(simple_dir)
        return d

    def compute_returns_table(base_df: pd.DataFrame) -> pd.DataFrame:
        col_avgpay = col_avgq
        col_base   = "نسبة المرتجع من المباع"
        col_new    = "نسبة نوع جديد من مرتجعات العميل"
        col_comp   = "نسبة نوع تعويض من مرتجعات العميل"
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
        if col_avgpay not in base_df.columns: return pd.DataFrame()
        avg_series = to_numeric(base_df[col_avgpay].map(clean_number)); max_avg = avg_series.max()
        pct_avg = np.where(max_avg>0, (avg_series / max_avg * 100).round(2), np.nan)
        pp = pd.Series(pct_avg).apply(_score_pp)
        def _one(col_name: str, label: str) -> pd.DataFrame:
            if col_name not in base_df.columns: return pd.DataFrame()
            tmp = to_numeric(base_df[col_name].map(clean_number)).replace([np.inf, -np.inf], np.nan)
            ref_vals = tmp[pp.between(5,10, inclusive="both")]
            ref_avg = ref_vals.mean()
            out = pd.DataFrame(index=base_df.index)
            out[f"[مرتجع] معيار ({label} 10–5)"] = round(ref_avg, 4) if pd.notna(ref_avg) else np.nan
            ratio = tmp / ref_avg if (ref_avg and ref_avg!=0) else np.nan
            out[f"[مرتجع] مضاعف ({label}) مقابل المعيار"] = ratio
            out[f"[مرتجع] تصنيف ({label})"] = tmp.apply(lambda x: _classify(x, ref_avg))
            out[f"[مرتجع] قيمة ({label})"] = tmp
            return out
        parts = [_one(col_base, "المرتجع من المباع"), _one(col_new, "النوع الجديد"), _one(col_comp, "نوع تعويض")]
        parts = [p for p in parts if not p.empty]
        if not parts: return pd.DataFrame()
        return pd.concat(parts, axis=1)

    df_main_only = df.copy()
    df_delta_all = compute_delta_table(df)
    df_returns_all = compute_returns_table(df)

    unified = df_original.copy()
    main_extra_cols = [c for c in df_main_only.columns if c not in df_original.columns]
    unified = unified.join(df_main_only[main_extra_cols].add_prefix("[أساسي] "))
    if df_delta_all is not None and not df_delta_all.empty: unified = unified.join(df_delta_all)
    if df_returns_all is not None and not df_returns_all.empty: unified = unified.join(df_returns_all)

    st.dataframe(unified, use_container_width=True)
    buf = BytesIO(); unified.to_excel(buf, index=False); buf.seek(0)
    st.download_button("⬇️ تحميل الملف الموحّد (Excel)", buf, file_name="نتائج_موحّدة_كل_التبويبات_v5_6_5.xlsx")

# ======================== نهاية الملف ========================
