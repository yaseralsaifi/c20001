# -*- coding: utf-8 -*-
import streamlit as st

from customer_ai.utils import normalize
from customer_ai.columns import read_uploaded_file, detect_columns, sidebar_column_mapping
from customer_ai.scoring import build_config_from_sidebar
from customer_ai.main_tab import render_main_tab
from customer_ai.delta_tab import render_delta_tab
from customer_ai.returns_tab import render_returns_tab
from customer_ai.diag_tab import render_diag_tab
from customer_ai.export_unified import render_unified_export

st.set_page_config(page_title="المساعد الذكي لتصنيف العملاء - v5.6.5", layout="wide")
st.title("المساعد الذكي لتصنيف العملاء وتحليل المديونية — v5.6.5")

st.sidebar.header("📂 ملف البيانات")
uploaded_file = st.sidebar.file_uploader("ارفع ملف Excel أو CSV", type=["xlsx", "csv"])

st.sidebar.header("⚙️ الإعدادات الأساسية")
config = build_config_from_sidebar()

if not uploaded_file:
    st.info("⬆️ ارفع ملف العملاء للبدء (Excel/CSV).")
    st.stop()

df = read_uploaded_file(uploaded_file)
df.columns = [normalize(c) for c in df.columns]
df_original = df.copy()

detected = detect_columns(df)
cols = sidebar_column_mapping(df, detected)

st.sidebar.caption(
    f"Detected ➜ المديونية: {detected['debt'] or '—'} | المتوسط: {detected['avgq'] or '—'} | "
    f"العمر: {detected['age'] or '—'} | أعلى متوسط: {detected['high'] or '—'}"
)

st.info(
    "سيتم استخدام الأعمدة: المديونية = **{}** ، متوسط السداد = **{}**{}".format(
        cols["debt"], cols["avgq"], f" ، العمر = **{cols['age']}**" if cols["age"] else ""
    )
)

tab_main, tab_delta, tab_returns, tab_diag = st.tabs([
    "🔎 التصنيف والتحليل الأساسي",
    "🔁 أعمدة الفارق المبسطة",
    "📊 تصنيفات المرتجع (مستقل)",
    "🛠️ تشخيص سريع"
])

with tab_main:
    render_main_tab(df, df_original, cols, config)

with tab_delta:
    render_delta_tab(df, cols, config)

with tab_returns:
    render_returns_tab(df, cols, config)

with tab_diag:
    render_diag_tab(df, cols, config)

render_unified_export(df, df_original, cols, config)
