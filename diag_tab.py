import streamlit as st

def render_diag_tab(df, cols, config):
    st.subheader("🛠️ تشخيص أسباب (0) في نقاط المخاطرة")

    col_avgq = cols["avgq"]
    col_debt = cols["debt"]

    if col_avgq not in df.columns:
        st.info("لا يمكن عرض التشخيص لعدم توفر عمود متوسط السداد.")
        return

    zero_avg = df[col_avgq] == 0
    st.write(f"عدد العملاء الذين متوسط سدادهم = 0: **{zero_avg.sum()}**")

    if col_debt in df.columns:
        sample = df.loc[zero_avg, [col_debt, col_avgq]].head(10)
        if not sample.empty:
            st.dataframe(sample, use_container_width=True)

    st.caption("عند متوسط سداد = 0، يتم إعطاء (0) نقاط مخاطرة حسب منطق الدالة الحالي.")
