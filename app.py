# -*- coding: utf-8 -*-
# Streamlit unified app (Arabic) — v5.6.9
# الجديد في v5.6.9:
# - تبويب "📈 تقارير المندوبين" (نظرة عامة + تفصيل + تنزيل Excel)
# - يحتفظ بكل تحسينات v5.6.8 (تصنيف بلا فجوات، نمطي نقاط المخاطرة، تنظيف قوي، أداء متّجه)

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
    _tz = ZoneInfo("Asia/Riyadh")
except Exception:
    _tz = None

# ======================== إعدادات الصفحة ========================
st.set_page_config(page_title="المساعد الذكي لتصنيف العملاء - v5.6.9", layout="wide")
st.title("المساعد الذكي لتصنيف العملاء وتحليل المديونية — v5.6.9")

# ======================== وظائف مساعدة عامة ========================
def now_stamp(fmt: str = "%Y%m%d_%H%M") -> str:
    try:
        return datetime.now(tz=_tz).strftime(fmt) if _tz else datetime.now().strftime(fmt)
    except Exception:
        return datetime.now().strftime(fmt)

@st.cache_data(show_spinner=False)
def load_df(f):
    try:
        name = f.name.lower()
        if name.endswith(".csv"):
            return pd.read_csv(f, encoding="utf-8-sig")
        return pd.read_excel(f, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"تعذر قراءة الملف: {e}")

def normalize(name: str) -> str:
    return str(name).strip().replace("\u200f", "").replace("\u200e", "")

def to_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def clean_number(x):
    """تحويل القيم النصية إلى رقم: أرقام عربية، فواصل عربية/إنجليزية، وإزالة % مع تنظيف محارف الاتجاه والمسافات الخاصة."""
    if pd.isna(x):
        return x
    try:
        s = str(x).strip()
        s = (s.replace('\u200f','').replace('\u200e','')
               .replace('\u202a','').replace('\u202c','')
               .replace('\xa0',' '))
        trans = {
            ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3',
            ord('٤'): '4', ord('٥'): '5', ord('٦'): '6', ord('٧'): '7',
            ord('٨'): '8', ord('٩'): '9',
            ord('٬'): ',', ord('،'): ',', ord('٫'): '.',
        }
        s = s.translate(trans).replace('%','').replace(',','')
        return s
    except Exception:
        return x

def norm_key(s: str) -> str:
    s = normalize(s)
    trans = {
        ord('آ'): 'ا', ord('أ'): 'ا', ord('إ'): 'ا',
        ord('ى'): 'ي', ord('ة'): 'ه', ord('ؤ'): 'و', ord('ئ'): 'ي',
        ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3',
        ord('٤'): '4', ord('٥'): '5', ord('٦'): '6', ord('٧'): '7',
        ord('٨'): '8', ord('٩'): '9',
        ord('٬'): ',', ord('،'): ',', ord('٫'): ',',
        ord('ـ'): '',
    }
    s = s.translate(trans).lower()
    return ''.join(ch for ch in s if ch.isalnum())

def match_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
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
default_pct_decimals = st.sidebar.number_input("عدد المنازل العشرية الافتراضي للحقول المئوية", value=2, step=1, min_value=0, max_value=6)
show_debug = st.sidebar.toggle("إظهار الأعمدة التشخيصية (Debug)", value=False)

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
    st.info("خصم تلقائي بالسالب لما بعد 60 يوم على شكل شرائح كل 30 يوم (قيم صحيحة فقط).")
    age_5 = st.number_input("≤ هذا العدد من الأيام = 5 نقاط", value=30, step=1)
    age_4 = st.number_input("≤ هذا العدد من الأيام = 4 نقاط", value=40, step=1)
    age_3 = st.number_input("≤ هذا العدد من الأيام = 3 نقاط", value=51, step=1)
    age_2 = st.number_input("≤ هذا العدد من الأيام = 2 نقاط", value=60, step=1)

    def score_debt_age_vec(days: pd.Series) -> pd.Series:
        d = pd.to_numeric(days, errors="coerce")
        res = np.full(len(d), np.nan)
        res = np.where(d <= age_5, 5, res)
        res = np.where((d > age_5) & (d <= age_4), 4, res)
        res = np.where((d > age_4) & (d <= age_3), 3, res)
        res = np.where((d > age_3) & (d <= age_2), 2, res)
        over = d - age_2
        penalties = -np.floor(np.maximum(over, 0) / 30.0)
        res = np.where(np.isnan(res), penalties, res)
        return pd.Series(res, index=days.index).fillna(0)

# ---- نقاط المخاطرة (المديونية ÷ متوسط السداد الربعي) ----
with st.sidebar.expander("نقاط المخاطرة (المديونية ÷ متوسط السداد الربعي)", expanded=False):
    st.info("يشمل مفتاح تبديل بين النمط الجديد (بدون سالب) والقديم (مع سالب عند المؤشر العالي).")
    r_5 = st.number_input("≤ هذا المؤشر = 5 نقاط", value=1.00, step=0.1, format="%.2f")
    r_4 = st.number_input("≤ هذا المؤشر = 4 نقاط", value=1.50, step=0.1, format="%.2f")
    r_3 = st.number_input("≤ هذا المؤشر = 3 نقاط", value=2.00, step=0.1, format="%.2f")
    r_2 = st.number_input("≤ هذا المؤشر = 2 نقاط", value=2.50, step=0.1, format="%.2f")
    r_1 = st.number_input("≤ هذا المؤشر = 1 نقطة", value=3.00, step=0.1, format="%.2f")

with st.sidebar.expander("⚖️ إعدادات نمط نقاط المخاطرة", expanded=False):
    risk_allow_negative = st.checkbox("تفعيل النمط القديم (إرجاع نقاط سالبة عند المؤشر العالي)", value=False)

def risk_ratio_and_score(debt: pd.Series, avg: pd.Series) -> tuple[pd.Series, pd.Series]:
    """
    قواعد النمط الجديد (الافتراضي):
      - إذا كانت المديونية <= 0 → 5 نقاط مباشرةً (بغض النظر عن المتوسط)
      - ≤ 1.0 → 5 نقاط
      - 1.1–1.5 → 4 نقاط
      - 1.6–2.0 → 3 نقاط
      - 2.1–2.5 → 2 نقاط
      - 2.6–<3.0 → 1 نقطة
      - = 3.0 أو > 3.0 → 0 نقاط
    النمط القديم (إذا فُعِّل risk_allow_negative=True):
      - نفس نطاقات البداية حتى 3.0، ثم 0، ثم -5 حتى 6.0، ثم -10 حتى 12.0 وما فوق
      - حالة متوسط=0 ومديونية>0 → -10
    """
    debt_f = pd.to_numeric(debt, errors="coerce").fillna(0.0)
    avg_f  = pd.to_numeric(avg,  errors="coerce")

    ratio_raw = debt_f / avg_f.replace(0, np.nan)
    ratio = np.where(ratio_raw < 0, 0.0, ratio_raw)
    ratio_s = pd.Series(ratio, index=debt.index).round(3)

    special_five_mask = debt_f <= 0

    if not risk_allow_negative:
        score = np.where(ratio <= 1.0, 5,
                 np.where(ratio <= 1.5, 4,
                 np.where(ratio <= 2.0, 3,
                 np.where(ratio <= 2.5, 2,
                 np.where(ratio < 3.0, 1,
                          0)))))
        score = np.where(special_five_mask, 5, score)
        return ratio_s, pd.Series(score, index=debt.index).astype(float)

    conditions = [
        ratio <= r_5,
        ratio <= r_4,
        ratio <= r_3,
        ratio <= r_2,
        ratio <= r_1,
        ratio <= 4.0,
        ratio <= 6.0,
        ratio <= 12.0
    ]
    choices = [5, 4, 3, 2, 1, 0, -5, -10]
    score_old = np.select(conditions, choices, default=-10).astype(float)

    avg_zero = avg_f.fillna(0).eq(0)
    debt_pos = debt_f.gt(0)
    score_old = np.where(avg_zero & debt_pos, -10.0, score_old)

    score_old = np.where(special_five_mask, 5.0, score_old)

    return ratio_s, pd.Series(score_old, index=debt.index).astype(float)

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
    df = load_df(uploaded_file)
except Exception as e:
    st.error(str(e))
    st.stop()

# توحيد أسماء الأعمدة
df.columns = [normalize(c) for c in df.columns]

# أسماء بديلة
_debt_aliases = ["المديونية","رصيد المديونية","إجمالي المديونية","اجمالي المديونية","رصيد","مستحقات","رصيد مستحق"]
_avgq_aliases = ["متوسط السداد الربعي","متوسط السداد","متوسط السداد 3 اشهر","متوسط السداد ٣ اشهر","متوسط السداد الشهري","المتوسط الشهري للسداد"]
_age_aliases  = ["عمر المديونية (يوم)","عمر المديونية","أيام المديونية","ايام المديونية","عمر الدين","عدد الايام"]
_high_avg_aliases = ["أعلى متوسط السداد الربعي","اعلى متوسط السداد الربعي","أقصى متوسط السداد الربعي","اقصى متوسط السداد"]

# اكتشاف تلقائي
detected_debt = match_col(df, _debt_aliases)
detected_avgq = match_col(df, _avgq_aliases)
detected_age  = match_col(df, _age_aliases)
detected_high = match_col(df, _high_avg_aliases)

# اختيار يدوي
st.sidebar.subheader("🧭 تعيين الأعمدة يدويًا (إن لزم)")
col_debt = st.sidebar.selectbox("عمود المديونية", df.columns, index=(df.columns.get_loc(detected_debt) if detected_debt else 0), key="col_debt")
col_avgq = st.sidebar.selectbox("عمود متوسط السداد الربعي", df.columns, index=(df.columns.get_loc(detected_avgq) if detected_avgq else 0), key="col_avgq")
col_age_opt = ["— (بدون) —"] + list(df.columns)
col_age_sel = st.sidebar.selectbox("عمود عمر المديونية (اختياري)", col_age_opt, index=(1 + df.columns.get_loc(detected_age)) if detected_age else 0, key="col_age")
col_high_avgq_sel = st.sidebar.selectbox("عمود أعلى متوسط السداد الربعي (للفارق)", ["— (غير مستخدم) —"] + list(df.columns), index=(1 + df.columns.get_loc(detected_high)) if detected_high else 0, key="col_high_avgq")
col_age = None if col_age_sel == "— (بدون) —" else col_age_sel
col_high_avgq = None if col_high_avgq_sel == "— (غير مستخدم) —" else col_high_avgq_sel

st.sidebar.caption(f"Detected ➜ المديونية: {detected_debt or '—'} | المتوسط: {detected_avgq or '—'} | العمر: {detected_age or '—'} | أعلى متوسط: {detected_high or '—'}")
st.info("سيتم استخدام الأعمدة: المديونية = **{}** ، متوسط السداد = **{}**{}".format(
    col_debt, col_avgq, f" ، العمر = **{col_age}**" if col_age else ""))

# نسخة أصلية
df_original = df.copy()

# ======================== تبويبات الواجهة ========================
tab_main, tab_delta, tab_returns, tab_diag, tab_reps = st.tabs([
    "🔎 التصنيف والتحليل الأساسي",
    "🔁 أعمدة الفارق المبسطة",
    "📊 تصنيفات المرتجع (مستقل)",
    "🛠️ تشخيص سريع",
    "📈 تقارير المندوبين"
])

# ======================== التبويب 1: التصنيف والتحليل ========================
with tab_main:
    st.subheader("🔎 التصنيف والتحليل الأساسي")

    missing = [c for c in [col_debt, col_avgq] if c not in df.columns]
    if missing:
        st.error(f"يجب توافر أعمدة: {missing}")
        st.stop()

    for c in [col_debt, col_avgq] + ([col_age] if col_age else []):
        df[c] = to_numeric(df[c].map(clean_number))

    max_avg = df[col_avgq].max(skipna=True)
    pct_avg = np.where(max_avg > 0, (df[col_avgq] / max_avg * 100), np.nan)
    df["نسبة من القائد (متوسط)"] = np.round(pct_avg, default_pct_decimals)

    df["نقاط القوة الشرائية"] = pd.Series(df["نسبة من القائد (متوسط)"].apply(score_purchase_power), index=df.index)
    df["نقاط الالتزام"] = score_debt_age_vec(df[col_age]) if col_age and col_age in df.columns else 0

    ratio, risk_points = risk_ratio_and_score(df[col_debt], df[col_avgq])
    df["مؤشر المخاطرة (مديونية/متوسط)"] = ratio
    df["نقاط المخاطرة"] = risk_points

    with st.sidebar.expander("التصنيف النهائي (حسب مجموع النقاط)", expanded=False):
        st.info("سلال مغلقة-مفتوحة بلا فجوات: [الحد الأدنى للفئة، الحد الأدنى للفئة الأعلى).")
        final_motazem_min = st.number_input("≥ هذا المجموع = ملتزم", value=16.0, step=0.5)
        final_jayed_min   = st.number_input("≥ هذا المجموع = جيد", value=12.0, step=0.5)
        final_fix_cap_min = st.number_input("≥ هذا المجموع = جدولة + تثبيت السقف", value=10.0, step=0.1)
        final_reduce_min  = st.number_input("≥ هذا المجموع = جدولة + تخفيف", value=7.0, step=0.1)
        def final_classification(score):
            if score >= final_motazem_min:
                return "ملتزم"
            elif (score >= final_jayed_min) and (score < final_motazem_min):
                return "جيد"
            elif (score >= final_fix_cap_min) and (score < final_jayed_min):
                return "جدوله مديونية وتثبيت السقف (حد أعلى المبيعات الآجل)"
            elif (score >= final_reduce_min) and (score < final_fix_cap_min):
                return "جدوله مديونية وتخفيف المبيعات الآجل"
            elif score >= 8:
                return "قبل النهاية"
            else:
                return "عميل غير مجدي"

    df["إجمالي النقاط"] = df[["نقاط القوة الشرائية", "نقاط الالتزام", "نقاط المخاطرة"]].sum(axis=1)
    df["التصنيف النهائي"] = df["إجمالي النقاط"].apply(final_classification)

    rep_col_candidates = ["اسم المندوب", "المندوب", "مندوب", "اسم مندوب"]
    rep_col = next((c for c in rep_col_candidates if c in df.columns), None)
    if rep_col is not None and col_debt in df.columns:
        cnt = df.groupby([rep_col, "التصنيف النهائي"])['إجمالي النقاط'].size().rename("عدد")
        cnt_by_class = cnt.groupby(level=1).transform("sum")
        share_count = (cnt / cnt_by_class.replace(0, np.nan) * 100).round(default_pct_decimals)

        debt_grp = df.groupby([rep_col, "التصنيف النهائي"])[col_debt].sum(min_count=1).rename("مديونية")
        debt_by_class = debt_grp.groupby(level=1).transform("sum")
        share_debt = (debt_grp / debt_by_class.replace(0, np.nan) * 100).round(default_pct_decimals)

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
        )
        df["نسبة الفئة داخل مديونية المندوب (%)"] = df["نسبة الفئة داخل مديونية المندوب (%)"].round(default_pct_decimals)
    else:
        df["نسبة المندوب من فئة العميل (بالعدد %)"] = np.nan
        df["نسبة المندوب من فئة العميل (بالمديونية %)"] = np.nan
        df["نسبة الفئة داخل مديونية المندوب (%)"] = np.nan

    total_debt = pd.to_numeric(df[col_debt], errors="coerce").sum(min_count=1)
    if pd.notna(total_debt) and total_debt != 0:
        class_debt = df.groupby("التصنيف النهائي")[col_debt].sum(min_count=1)
        share_map = (class_debt / total_debt * 100).round(default_pct_decimals).to_dict()
        df["نسبة التصنيف من إجمالي المديونية (%)"] = df["التصنيف النهائي"].map(share_map)
    else:
        df["نسبة التصنيف من إجمالي المديونية (%)"] = np.nan

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

    if not show_debug:
        hide_cols = ["مبلغ الانحراف (للـ3 أشهر)", "قسط الانحراف الشهري", "فقد_نقاط_التزام/مخاطرة؟"]
        preview_cols = [c for c in df.columns if c not in hide_cols]
    else:
        preview_cols = list(df.columns)

    st.success("✅ تم إعداد التصنيف والخطة بنجاح.")
    st.dataframe(df[preview_cols], use_container_width=True)

    out_main = BytesIO()
    df.to_excel(out_main, index=False); out_main.seek(0)
    st.download_button("⬇️ تحميل الملف الناتج (Excel)", out_main, file_name=f"نتائج_التصنيف_v5_6_9_{now_stamp()}.xlsx")

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
            df_delta["نسبة من القائد (متوسط)"] = (avg_series / max_avg * 100).round(default_pct_decimals)
            df_delta["نسبة من القائد (أعلى)"]   = (high_series / max_high * 100).round(default_pct_decimals)
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
            st.download_button("⬇️ تحميل ملف الفارق (Excel)", out_delta, file_name=f"نتائج_أعمدة_الفارق_v5_6_9_{now_stamp()}.xlsx")
    else:
        st.info("للحساب هنا يلزم وجود الأعمدة المختارة في الملف.")

# ======================== التبويب 3: تصنيفات المرتجع (مستقل) ========================
with tab_returns:
    st.subheader("📊 تصنيفات المرتجع — مستقل (v5.6.9)")
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
        if col_name not in df_in.columns:
            st.warning(f"⚠️ لم يتم العثور على العمود: {col_name}")
            return None
        if col_avgpay not in df_in.columns:
            st.warning(f"لا يمكن احتساب مجموعة المرجع لعدم توفر '{col_avgpay}'.")
            return None

        avg_series = pd.to_numeric(df_in[col_avgpay].map(clean_number), errors="coerce")
        vals       = pd.to_numeric(df_in[col_name].map(clean_number),   errors="coerce")

        max_avg = avg_series.max(skipna=True)
        pct = np.where(max_avg > 0, (avg_series / max_avg * 100), np.nan)
        pp  = pd.Series(pct, index=df_in.index).apply(score_purchase_power_for_returns)
        mask_ref = pp.between(5, 10, inclusive="both")

        ref_avg = vals[mask_ref].mean(skipna=True)

        out = pd.DataFrame(index=df_in.index)
        out[col_name] = vals
        out[f"معيار المرتجع ({label} 10–5)"] = ref_avg

        if not (pd.notna(ref_avg) and ref_avg != 0):
            out[f"مضاعف المرتجع ({label}) مقابل المعيار"] = np.nan
            out[f"تصنيف المرتجع ({label})"] = "بيانات غير كافية"
            return out

        ratio = vals / ref_avg
        def label_ratio_arr(x: pd.Series):
            return np.select(
                [x <= m_ok, x <= m_watch, x <= m_high, ~np.isnan(x)],
                ["ضمن المعيار","يحتاج متابعة","مرتفع","مرتفع جدًا"],
                default="بيانات غير كافية"
            )
        out[f"مضاعف المرتجع ({label}) مقابل المعيار"] = ratio
        out[f"تصنيف المرتجع ({label})"] = label_ratio_arr(ratio)
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
        st.download_button("⬇️ تحميل الملف (تصنيفات المرتجع)", out_ret, file_name=f"نتائج_تصنيفات_المرتجع_v5_6_9_{now_stamp()}.xlsx")

# ======================== التبويب 4: تشخيص سريع ========================
with tab_diag:
    st.subheader("🛠️ تشخيص أسباب (0) في نقاط المخاطرة")
    zero_avg = df[col_avgq] == 0
    sample = df.loc[zero_avg, [col_debt, col_avgq]].head(10)
    st.write(f"عدد العملاء الذين متوسط سدادهم = 0: **{int(zero_avg.sum())}**")
    if not sample.empty:
        st.dataframe(sample, use_container_width=True)
    st.caption("إذا كان متوسط السداد = 0 والمديونية > 0 ففي النمط القديم تُعطى -10 نقاط؛ وفي النمط الجديد تُعطى 0 نقاط، بينما المديونية ≤ 0 تعطي 5 نقاط دائمًا.")

# ======================== التبويب 5: تقارير المندوبين ========================
with tab_reps:
    st.subheader("📈 تقارير المندوبين — نظرة عامة وتنزيل Excel")

    rep_col_candidates = ["اسم المندوب", "المندوب", "مندوب", "اسم مندوب"]
    rep_col = next((c for c in rep_col_candidates if c in df.columns), None)

    if rep_col is None:
        st.info("لا يوجد عمود مندوب في الملف. رجاءً أضف عمودًا مثل: (اسم المندوب / المندوب) لاستخدام هذا التبويب.")
    else:
        needed_cols = [rep_col, col_debt, col_avgq, "إجمالي النقاط", "التصنيف النهائي", "مؤشر المخاطرة (مديونية/متوسط)", "نقاط المخاطرة", "نقاط القوة الشرائية", "نقاط الالتزام"]
        for nc in needed_cols:
            if nc not in df.columns:
                st.warning(f"تم فقد عمود ضروري للحساب: {nc}. افتح التبويب الرئيسي أولاً لضبط النتائج.")
                st.stop()

        agg = df.groupby(rep_col).agg(
            عدد_العملاء=("إجمالي النقاط", "size"),
            إجمالي_المديونية=(col_debt, "sum"),
            متوسط_السداد=(col_avgq, "mean"),
            متوسط_المؤشر=("مؤشر المخاطرة (مديونية/متوسط)", "mean"),
            متوسط_النقاط=("إجمالي النقاط", "mean"),
            متوسط_نقاط_المخاطرة=("نقاط المخاطرة", "mean"),
            متوسط_نقاط_الالتزام=("نقاط الالتزام", "mean"),
            متوسط_القوة_الشرائية=("نقاط القوة الشرائية", "mean"),
        ).reset_index()

        cnt_pivot = pd.pivot_table(df, index=rep_col, columns="التصنيف النهائي", values="إجمالي النقاط", aggfunc="size", fill_value=0)
        debt_pivot = pd.pivot_table(df, index=rep_col, columns="التصنيف النهائي", values=col_debt, aggfunc="sum", fill_value=0.0)

        cnt_pct = (cnt_pivot.div(cnt_pivot.sum(axis=1).replace(0, np.nan), axis=0) * 100).round(default_pct_decimals)
        debt_pct = (debt_pivot.div(debt_pivot.sum(axis=1).replace(0, np.nan), axis=0) * 100).round(default_pct_decimals)

        overview = agg.merge(cnt_pivot.add_prefix("عدد/"), on=rep_col, how="left") \
                      .merge(cnt_pct.add_prefix("نسبة_عدد%/"), on=rep_col, how="left") \
                      .merge(debt_pivot.add_prefix("مديونية/"), on=rep_col, how="left") \
                      .merge(debt_pct.add_prefix("نسبة_مديونية%/"), on=rep_col, how="left")

        st.markdown("#### نظرة عامة لكل مندوب")
        st.dataframe(overview, use_container_width=True)

        by_class = df.groupby([rep_col, "التصنيف النهائي"]).agg(
            عدد=("إجمالي النقاط", "size"),
            مديونية=(col_debt, "sum"),
            متوسط_السداد=(col_avgq, "mean"),
            متوسط_المؤشر=("مؤشر المخاطرة (مديونية/متوسط)", "mean"),
            متوسط_النقاط=("إجمالي النقاط", "mean"),
        ).reset_index()

        st.markdown("#### تفصيل حسب (المندوب × الفئة)")
        st.dataframe(by_class, use_container_width=True)

        st.markdown("#### بطاقات ملخّصة")
        reps = overview[rep_col].tolist()
        default_rep = reps[0] if reps else None
        chosen_rep = st.selectbox("اختر المندوب لعرض بطاقات سريعة", reps, index=0 if default_rep else None, key="rep_cards")
        if chosen_rep:
            row = overview[overview[rep_col] == chosen_rep].iloc[0]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("عدد العملاء", int(row["عدد_العملاء"]))
            c2.metric("إجمالي المديونية", round(float(row["إجمالي_المديونية"]), 2))
            c3.metric("متوسط المؤشر", round(float(row["متوسط_المؤشر"]), 2) if pd.notna(row["متوسط_المؤشر"]) else "—")
            c4.metric("متوسط النقاط", round(float(row["متوسط_النقاط"]), 2))

        out = BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            overview.to_excel(writer, sheet_name="Overview", index=False)
            by_class.to_excel(writer, sheet_name="ByClass", index=False)
            cnt_pivot.reset_index().to_excel(writer, sheet_name="CountPivot", index=False)
            cnt_pct.reset_index().to_excel(writer, sheet_name="CountPct", index=False)
            debt_pivot.reset_index().to_excel(writer, sheet_name="DebtPivot", index=False)
            debt_pct.reset_index().to_excel(writer, sheet_name="DebtPct", index=False)
        out.seek(0)
        st.download_button("⬇️ تنزيل تقارير المندوبين (Excel)", out, file_name=f"تقارير_المندوبين_{now_stamp()}.xlsx")

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
        d["[فارق] نسبة من القائد (متوسط)"] = (avg_series / max_avg * 100).round(default_pct_decimals)
        d["[فارق] نسبة من القائد (أعلى)"]   = (high_series / max_high * 100).round(default_pct_decimals)
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
        col_avgpay_local = col_avgq
        col_base_local   = "نسبة المرتجع من المباع"
        col_new_local    = "نسبة نوع جديد من مرتجعات العميل"
        col_comp_local   = "نسبة نوع تعويض من مرتجعات العميل"
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
        if col_avgpay_local not in base_df.columns: return pd.DataFrame()
        avg_series = to_numeric(base_df[col_avgpay_local].map(clean_number)); max_avg = avg_series.max()
        pct_avg = np.where(max_avg>0, (avg_series / max_avg * 100).round(default_pct_decimals), np.nan)
        pp = pd.Series(pct_avg).apply(_score_pp)
        def _one(col_name: str, label: str) -> pd.DataFrame:
            if col_name not in base_df.columns: return pd.DataFrame()
            tmp = to_numeric(base_df[col_name].map(clean_number)).replace([np.inf, -np.inf], np.nan)
            ref_vals = tmp[pp.between(5,10, inclusive="both")]
            ref_avg = ref_vals.mean()
            out = pd.DataFrame(index=base_df.index)
            out[f"[مرتجع] معيار ({label} 10–5)"] = round(ref_avg, 4) if pd.notna(ref_avg) else np.nan
            ratio = tmp / ref_avg if (pd.notna(ref_avg) and ref_avg != 0) else np.nan
            out[f"[مرتجع] مضاعف ({label}) مقابل المعيار"] = ratio
            out[f"[مرتجع] تصنيف ({label})"] = tmp.apply(lambda x: _classify(x, ref_avg))
            out[f"[مرتجع] قيمة ({label})"] = tmp
            return out
        parts = [_one(col_base_local, "المرتجع من المباع"), _one(col_new_local, "النوع الجديد"), _one(col_comp_local, "نوع تعويض")]
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
    st.download_button("⬇️ تحميل الملف الموحّد (Excel)", buf, file_name=f"نتائج_موحّدة_كل_التبويبات_v5_6_9_{now_stamp()}.xlsx")

# ======================== نهاية الملف ========================
