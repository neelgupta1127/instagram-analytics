# =========================================================
# INSTAGRAM FOLLOWERS / FOLLOWING ANALYTICS TOOL
# Streamlit Web App Version
# =========================================================

import streamlit as st
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Instagram Analytics",
    page_icon="📊",
    layout="centered"
)

# =========================================================
# CUSTOM STYLE
# =========================================================

st.markdown("""
    <style>
        body { background-color: #0d0d0d; }
        .main { background-color: #0d0d0d; }
        h1 { color: #E1306C; font-family: 'Georgia', serif; }
        .stMetric { background-color: #1a1a2e; border-radius: 10px; padding: 10px; }
        .stDownloadButton button {
            background-color: #E1306C;
            color: white;
            font-weight: bold;
            border-radius: 8px;
            width: 100%;
        }
    </style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================

st.title("📊 Instagram Followers Analytics")
st.markdown("Upload your **Instagram exported HTML files** to analyze your followers and following relationships.")
st.markdown("---")

# =========================================================
# FUNCTION — EXTRACT USERNAMES
# =========================================================

def extract_usernames(html_bytes):
    soup = BeautifulSoup(html_bytes, 'lxml')
    links = soup.find_all('a')
    usernames = list(set(
        link.get_text(strip=True)
        for link in links
        if link.get_text(strip=True)
    ))
    usernames.sort()
    return usernames

# =========================================================
# FILE UPLOAD
# =========================================================

st.subheader("📁 Upload Instagram Export Files")

col1, col2 = st.columns(2)

with col1:
    followers_file = st.file_uploader("📥 followers_1.html", type=["html", "htm"])

with col2:
    following_file = st.file_uploader("📤 following.html", type=["html", "htm"])

# =========================================================
# ANALYSIS
# =========================================================

if followers_file and following_file:

    with st.spinner("🔍 Extracting usernames and analyzing..."):

        # Extract
        followers = extract_usernames(followers_file.read())
        following = extract_usernames(following_file.read())

        # Convert to sets
        followers_set = set(followers)
        following_set = set(following)

        # Set operations
        mutual_followers    = sorted(followers_set & following_set)
        not_following_back  = sorted(following_set - followers_set)
        fans                = sorted(followers_set - following_set)

    st.success("✅ Analysis Complete!")
    st.markdown("---")

    # =========================================================
    # SUMMARY METRICS
    # =========================================================

    st.subheader("📈 Summary")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("👥 Followers",         len(followers_set))
    c2.metric("➡️ Following",          len(following_set))
    c3.metric("🤝 Mutual",            len(mutual_followers))
    c4.metric("❌ Not Following Back", len(not_following_back))
    c5.metric("⭐ Fans",              len(fans))

    st.markdown("---")

    # =========================================================
    # SET THEORY INFO
    # =========================================================

    with st.expander("📐 Set Theory Logic Used"):
        st.markdown("""
        | Category | Formula | Meaning |
        |---|---|---|
        | Mutual Followers | A ∩ B | Both follow each other |
        | Not Following Back | B - A | You follow them, they don't follow back |
        | Fans | A - B | They follow you, you don't follow back |
        """)

    # =========================================================
    # RESULTS TABS
    # =========================================================

    st.subheader("📋 Results")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🤝 Mutual",
        "❌ Not Following Back",
        "⭐ Fans",
        "👥 All Followers",
        "➡️ All Following"
    ])

    with tab1:
        st.write(f"**{len(mutual_followers)} accounts** — Both follow each other")
        st.dataframe(pd.DataFrame(mutual_followers, columns=["Username"]), use_container_width=True)

    with tab2:
        st.write(f"**{len(not_following_back)} accounts** — You follow them but they don't follow you back")
        st.dataframe(pd.DataFrame(not_following_back, columns=["Username"]), use_container_width=True)

    with tab3:
        st.write(f"**{len(fans)} accounts** — They follow you but you don't follow them back")
        st.dataframe(pd.DataFrame(fans, columns=["Username"]), use_container_width=True)

    with tab4:
        st.write(f"**{len(followers_set)} total followers**")
        st.dataframe(pd.DataFrame(sorted(followers_set), columns=["Username"]), use_container_width=True)

    with tab5:
        st.write(f"**{len(following_set)} total following**")
        st.dataframe(pd.DataFrame(sorted(following_set), columns=["Username"]), use_container_width=True)

    st.markdown("---")

    # =========================================================
    # GENERATE EXCEL REPORT
    # =========================================================

    st.subheader("📥 Download Excel Report")

    output = io.BytesIO()

    summary_df = pd.DataFrame({
        "Metric": ["Total Followers", "Total Following", "Mutual Followers", "Not Following Back", "Fans"],
        "Count":  [len(followers_set), len(following_set), len(mutual_followers), len(not_following_back), len(fans)]
    })

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer,                                                              sheet_name='Summary',           index=False)
        pd.DataFrame(sorted(followers_set),  columns=["Followers"]).to_excel(writer,            sheet_name='Followers',         index=False)
        pd.DataFrame(sorted(following_set),  columns=["Following"]).to_excel(writer,            sheet_name='Following',         index=False)
        pd.DataFrame(mutual_followers,        columns=["Mutual Followers"]).to_excel(writer,     sheet_name='Mutual Followers',  index=False)
        pd.DataFrame(not_following_back,      columns=["Not Following Back"]).to_excel(writer,   sheet_name='Not Following Back',index=False)
        pd.DataFrame(fans,                    columns=["Fans"]).to_excel(writer,                 sheet_name='Fans',              index=False)

    # Format Excel
    output.seek(0)
    workbook = load_workbook(output)

    header_fill = PatternFill(start_color="E1306C", end_color="E1306C", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 5

    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)

    st.download_button(
        label="📥 Download Instagram_Analytics_Report.xlsx",
        data=final_output,
        file_name="Instagram_Analytics_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Please upload both **followers_1.html** and **following.html** to begin analysis.")
    st.markdown("""
    **How to get your Instagram export files:**
    1. Open Instagram → Settings → Account → Download your data
    2. Request download and wait for the email
    3. Extract the ZIP file
    4. Find `followers_1.html` and `following.html`
    5. Upload them above ✅
    """)

# =========================================================
# FOOTER
# =========================================================

st.markdown("---")
st.markdown("<center>🔒 100% Offline · No login required · No scraping · Official Instagram data only</center>", unsafe_allow_html=True)
