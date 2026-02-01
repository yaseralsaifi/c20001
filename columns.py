import pandas as pd
import streamlit as st
from .utils import match_col

DEBT_ALIASES = ["المديونية","رصيد المديونية","إجمالي المديونية","اجمالي المديونية","رصيد","مستحقات","رصيد مستحق"]
AVGQ_ALIASES = ["متوسط السداد الربعي","متوسط السداد","متوسط السداد 3 اشهر","متوسط السداد ٣ اشهر","متوسط السداد الشهري","المتوسط الشهري للسداد"]
AGE_ALIASES  = ["عمر المديونية (يوم)","عمر المديونية","أيام المديونية","ايام المديونية","عمر الدين","عدد الايام"]
HIGH_ALIASES = ["أعلى متوسط السداد الربعي","اعلى متوسط السداد الربعي","أقصى متوسط السداد الربعي","اقصى متوسط السداد"]

def read_uploaded_file(uploaded_file):
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            return pd.read_csv(uploaded_file)
        return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        st.stop()

def detect_columns(df: pd.DataFrame) -> dict:
    return {
        "debt": match_col(df, DEBT_ALIASES),
        "avgq": match_col(df, AVGQ_ALIASES),
        "age":  match_col(df, AGE_ALIASES),
        "high": match_col(df, HIGH_ALIASES),
    }

def sidebar_column_mapping(df: pd.DataFrame, detected: dict) -> dict:
    st.sidebar.subheader("🧭 تعيين الأعمدة يدويًا (إن لزم)")
    cols = list(df.columns)

    def idx(name):
        return cols.index(name) if name in cols else 0

    col_debt = st.sidebar.selectbox("عمود المديونية", cols, index=(idx(detected["debt"]) if detected["debt"] else 0))
    col_avgq = st.sidebar.selectbox("عمود متوسط السداد الربعي", cols, index=(idx(detected["avgq"]) if detected["avgq"] else 0))

    col_age_opt = ["— (بدون) —"] + cols
    col_age = st.sidebar.selectbox("عمود عمر المديونية (اختياري)", col_age_opt,
                                   index=(1 + idx(detected["age"])) if detected["age"] else 0)
    if col_age == "— (بدون) —":
        col_age = None

    col_high_opt = ["— (غير مستخدم) —"] + cols
    col_high = st.sidebar.selectbox("عمود أعلى متوسط السداد الربعي (للفارق)", col_high_opt,
                                    index=(1 + idx(detected["high"])) if detected["high"] else 0)
    if col_high == "— (غير مستخدم) —":
        col_high = None

    return {"debt": col_debt, "avgq": col_avgq, "age": col_age, "high_avgq": col_high}
