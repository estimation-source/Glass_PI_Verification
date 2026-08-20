import base64
import os
import sys
import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image

# Root directory path जोडण्यासाठी
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.excel_reader import process_uploaded_files
from src.pdf_reader import read_pdf
from verifier import verify_pi_against_excel

# ============================================================
# 1. STREAMLIT PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="WIN-SQUARE | Glass PI Verification System",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 2. FIX CSS: Header चालू ठेवून Sidebar Toggle Button Visible ठेवणे
# ============================================================
st.markdown("""
    <style>
    header[data-testid="stHeader"] {
        z-index: 99999 !important;
        background: transparent !important;
    }

    button[data-testid="stSidebarCollapsedControl"],
    button[data-testid="stSidebarNavCollapseButton"] {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        background-color: #FF4B4B !important;
        color: white !important;
        border-radius: 8px !important;
        position: fixed !important;
        top: 12px !important;
        left: 12px !important;
        z-index: 999999 !important;
        box-shadow: 0px 3px 8px rgba(0,0,0,0.3) !important;
    }

    button[data-testid="stSidebarCollapsedControl"] svg,
    button[data-testid="stSidebarNavCollapseButton"] svg {
        fill: white !important;
        color: white !important;
        width: 22px !important;
        height: 22px !important;
    }

    [data-testid="stStatusWidget"],
    #MainMenu, 
    footer {
        display: none !important;
        visibility: hidden !important;
    }
    </style>
""", unsafe_allow_html=True)

# State Management Initialization
if "excel_uploader_key" not in st.session_state:
    st.session_state["excel_uploader_key"] = 0

if "pdf_uploader_key" not in st.session_state:
    st.session_state["pdf_uploader_key"] = 0

# ============================================================
# 3. UI LAYOUT, FONTS & EXACT BUTTON CSS
# ============================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #f4f6f9;
        color: #334155;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }

    [data-testid="stSidebar"] {
        background-color: #f1f5f9;
        border-right: 1px solid #e2e8f0;
    }
    
    .quick-guide-title {
        font-size: 15px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 15px;
        margin-bottom: 12px;
    }
    
    .quick-guide-step {
        font-size: 13px;
        color: #475569;
        margin-bottom: 10px;
        line-height: 1.4;
    }

    .hero-container {
        background: #ffffff;
        border-radius: 16px;
        padding: 24px 30px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        margin-bottom: 24px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .hero-title-text {
        font-size: 22px;
        font-weight: 800;
        color: #0f172a;
        margin: 0;
    }

    .hero-sub-text {
        font-size: 13px;
        color: #64748b;
        margin-top: 4px;
    }

    .step-title {
        font-size: 16px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* PRIMARY BLUE BUTTON - COMPACT & NORMAL FONT WEIGHT */
    div.stButton > button[kind="primary"] {
        background-color: #2563eb !important;
        background: #2563eb !important;
        border: 1px solid #2563eb !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(37, 99, 235, 0.2) !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #1d4ed8 !important;
        background: #1d4ed8 !important;
    }

    /* SECONDARY RED BUTTON - COMPACT & NORMAL FONT WEIGHT */
    div.stButton > button[kind="secondary"] {
        background-color: #dc2626 !important;
        background: #dc2626 !important;
        border: 1px solid #dc2626 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(220, 38, 38, 0.2) !important;
    }
    div.stButton > button[kind="secondary"]:hover {
        background-color: #b91c1c !important;
        background: #b91c1c !important;
    }

    /* FORCE NORMAL WEIGHT (NOT BOLD) & WHITE TEXT */
    div.stButton > button p, div.stButton > button span {
        color: #ffffff !important;
        font-weight: 500 !important;
        font-size: 13px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def get_image_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

# =========================================================
# SIDEBAR STRUCTURE MATCH
# =========================================================
with st.sidebar:
    logo_file = get_image_path("logo.png")
    if os.path.exists(logo_file):
        col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
        with col_s2:
            st.image(Image.open(logo_file), width=110)
    else:
        st.markdown("<h2 style='text-align: center; color:#1e293b;'><b>win square</b></h2>", unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("<div class='quick-guide-title'>💡 Quick Guide</div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class='quick-guide-step'><b>1.</b> Upload Excel BOQ file(s) in Step 1.</div>
        <div class='quick-guide-step'><b>2.</b> Click <b>Extract Excel Data</b>.</div>
        <div class='quick-guide-step'><b>3.</b> Upload PI PDF file(s) in Step 2.</div>
        <div class='quick-guide-step'><b>4.</b> Click <b>Extract PDF Data</b>.</div>
        <div class='quick-guide-step'><b>5.</b> Run <b>Verify Data</b> to view mismatch analytics.</div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# HEADER HERO BANNER
# =========================================================
st.markdown(
    """
    <div class="hero-container">
        <div>
            <div class="hero-title-text">Glass PI Verification System</div>
            <div class="hero-sub-text">Automated BOQ vs PI Data Matching & Reconciliation</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# STEP 1 : Extract Data From Excel Sheets
# ============================================================
st.markdown("<div class='step-title'>📂 Step 1: Upload BOQ Excel Files</div>", unsafe_allow_html=True)

uploaded_excel_files = st.file_uploader(
    "Upload Excel File(s) (.xlsx, .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key=f"excel_uploader_{st.session_state['excel_uploader_key']}",
    label_visibility="collapsed"
)

st.markdown("<br>", unsafe_allow_html=True)

col_ex1, col_ex2, _ = st.columns([1, 1, 6])

with col_ex1:
    btn_extract_ex = st.button("🔗 Extract Excel Data", type="primary", use_container_width=False)

with col_ex2:
    btn_reset_ex = st.button("🗑️ Reset Excel", type="secondary", use_container_width=False)

if btn_extract_ex:
    if uploaded_excel_files:
        with st.spinner("Extracting Excel BOQ Data..."):
            excel_df = process_uploaded_files(uploaded_excel_files)
            if excel_df is not None and not excel_df.empty:
                st.session_state["excel_df"] = excel_df
                st.toast(f"Successfully extracted {len(excel_df)} rows from Excel!", icon="✅")
            else:
                st.error("Could not extract valid data from Excel file(s).")
    else:
        st.warning("Please upload Excel file(s) first.")

if btn_reset_ex:
    st.session_state["excel_uploader_key"] += 1
    st.session_state.pop("excel_df", None)
    st.session_state.pop("verification_df", None)
    st.rerun()

if "excel_df" in st.session_state and not st.session_state["excel_df"].empty:
    with st.expander("📄 View Extracted Excel BOQ Data", expanded=False):
        st.dataframe(st.session_state["excel_df"], use_container_width=True)

st.markdown("---")

# ============================================================
# STEP 2 : Extract Data From PI PDF Files
# ============================================================
st.markdown("<div class='step-title'>📂 Step 2: Upload PI PDF Files</div>", unsafe_allow_html=True)

uploaded_pdf_files = st.file_uploader(
    "Upload PI PDF File(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state['pdf_uploader_key']}",
    label_visibility="collapsed"
)

st.markdown("<br>", unsafe_allow_html=True)

col_pdf1, col_pdf2, _ = st.columns([1, 1, 6])

with col_pdf1:
    btn_extract_pdf = st.button("🔍 Extract PDF Data", type="primary", use_container_width=False)

with col_pdf2:
    btn_reset_pdf = st.button("🗑️ Reset PDF", type="secondary", use_container_width=False)

if btn_extract_pdf:
    if uploaded_pdf_files:
        pdf_dfs = []
        with st.spinner("Extracting Data from PDF(s)..."):
            for pdf_file in uploaded_pdf_files:
                temp_path = f"temp_{pdf_file.name}"
                try:
                    with open(temp_path, "wb") as f:
                        f.write(pdf_file.getbuffer())
                    
                    df_pdf = read_pdf(temp_path)
                    if df_pdf is not None and not df_pdf.empty:
                        pdf_dfs.append(df_pdf)
                finally:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

        if pdf_dfs:
            combined_pdf_df = pd.concat(pdf_dfs, ignore_index=True)
            st.session_state["pdf_df"] = combined_pdf_df
            st.toast(f"Successfully extracted {len(combined_pdf_df)} rows from PDF(s)!", icon="✅")
        else:
            st.error("Could not extract valid data from PDF file(s).")
    else:
        st.warning("Please upload PDF file(s) first.")

if btn_reset_pdf:
    st.session_state["pdf_uploader_key"] += 1
    st.session_state.pop("pdf_df", None)
    st.session_state.pop("verification_df", None)
    st.rerun()

if "pdf_df" in st.session_state and not st.session_state["pdf_df"].empty:
    with st.expander("📄 View Extracted PDF Data", expanded=False):
        st.dataframe(st.session_state["pdf_df"], use_container_width=True)

st.markdown("---")

# ============================================================
# STEP 3 : Verify Data & Matching
# ============================================================
st.markdown("<div class='step-title'>📂 Step 3: Verify Data & Run Matching</div>", unsafe_allow_html=True)

excel_ready = "excel_df" in st.session_state and not st.session_state["excel_df"].empty
pdf_ready = "pdf_df" in st.session_state and not st.session_state["pdf_df"].empty

if not excel_ready or not pdf_ready:
    st.info("💡 Please extract data from Step 1 (Excel) and Step 2 (PDF) first.")
else:
    col_v1, _ = st.columns([1, 7])
    with col_v1:
        btn_verify = st.button("⚡ Run Verification", type="primary", use_container_width=False)
    
    if btn_verify:
        with st.spinner("Matching and Verifying Data..."):
            v_df = verify_pi_against_excel(
                st.session_state["excel_df"], 
                st.session_state["pdf_df"]
            )
            st.session_state["verification_df"] = v_df

# ============================================================
# VERIFICATION SUMMARY & ANALYTICS DASHBOARD
# ============================================================
if 'verification_df' in st.session_state and not st.session_state["verification_df"].empty:
    df_res = st.session_state["verification_df"]

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='step-title'>📊 Verification Summary & Analytics</div>", unsafe_allow_html=True)

    total_items = len(df_res)
    exact_matches = len(df_res[df_res["Verification Status"] == "✅ MATCHED"])
    mismatches = total_items - exact_matches
    match_percentage = (exact_matches / total_items) * 100 if total_items > 0 else 0

    col1, col2, col3, _ = st.columns([1.2, 1.2, 1.2, 2.4])
    
    with col1:
        st.markdown(f"""
            <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 18px;">
                <p style="color: #64748b; font-size: 11px; font-weight: 700; margin: 0; text-transform: uppercase;">TOTAL BOQ ITEMS</p>
                <h3 style="color: #0f172a; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{total_items}</h3>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 14px 18px;">
                <p style="color: #166534; font-size: 11px; font-weight: 700; margin: 0; text-transform: uppercase;">EXACT MATCHES</p>
                <div style="display: flex; align-items: baseline; gap: 8px;">
                    <h3 style="color: #15803d; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{exact_matches}</h3>
                    <span style="color: #166534; font-weight: 700; font-size: 12px;">↑ {match_percentage:.1f}%</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        bg_color = "#fef2f2" if mismatches > 0 else "#ffffff"
        border_color = "#fecaca" if mismatches > 0 else "#e2e8f0"
        text_color = "#991b1b" if mismatches > 0 else "#64748b"
        num_color = "#dc2626" if mismatches > 0 else "#0f172a"

        st.markdown(f"""
            <div style="background: {bg_color}; border: 1px solid {border_color}; border-radius: 8px; padding: 14px 18px;">
                <p style="color: {text_color}; font-size: 11px; font-weight: 700; margin: 0; text-transform: uppercase;">MISMATCHES / MISSING</p>
                <h3 style="color: {num_color}; font-size: 24px; font-weight: 800; margin: 4px 0 0 0;">{mismatches}</h3>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns([1, 1])

    with c1:
        st.markdown("<h5 style='color: #334155; font-weight: 700; margin-bottom: 12px;'>Overall Verification Status</h5>", unsafe_allow_html=True)
        status_counts = df_res["Verification Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]

        fig_pie = px.pie(
            status_counts, 
            values="Count", 
            names="Status", 
            hole=0.6,
            color="Status",
            color_discrete_map={
                "✅ MATCHED": "#10b981",
                "⚠️ DIMENSION MISMATCH": "#f59e0b",
                "⚠️ QTY MISMATCH": "#eab308",
                "⚠️ GLASS SPEC MISMATCH": "#ef4444",
                "❌ NOT FOUND IN PDF": "#6b7280"
            }
        )
        
        fig_pie.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
        fig_pie.update_layout(
            showlegend=False, 
            height=300, 
            margin=dict(t=10, b=10, l=10, r=10),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        st.markdown("<h5 style='color: #334155; font-weight: 700; margin-bottom: 12px;'>Breakdown by Issue Type</h5>", unsafe_allow_html=True)
        mismatch_df = df_res[df_res["Verification Status"] != "✅ MATCHED"]
        
        if mismatch_df.empty:
            st.markdown("""
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 18px; text-align: left;">
                    <p style="color: #15803d; font-size: 13px; font-weight: 600; margin: 0;">🎉 Perfect 100% Match! No mismatches to display.</p>
                </div>
            """, unsafe_allow_html=True)
        else:
            bar_data = mismatch_df["Verification Status"].value_counts().reset_index()
            bar_data.columns = ["Issue Type", "Count"]

            fig_bar = px.bar(
                bar_data, 
                x="Issue Type", 
                y="Count", 
                color="Issue Type",
                text="Count",
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig_bar.update_layout(
                xaxis_title="", 
                yaxis_title="Count", 
                showlegend=False, 
                height=280,
                margin=dict(t=10, b=10, l=10, r=10),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # Detailed Table Section
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='step-title'>📑 Detailed Verification Results</div>", unsafe_allow_html=True)
    
    col_f1, _ = st.columns([2, 3])
    with col_f1:
        selected_status = st.selectbox(
            "Filter Results by Status:", 
            options=["ALL"] + list(df_res["Verification Status"].unique())
        )

    display_df = df_res[df_res["Verification Status"] == selected_status] if selected_status != "ALL" else df_res
    st.dataframe(display_df, use_container_width=True, height=400)
