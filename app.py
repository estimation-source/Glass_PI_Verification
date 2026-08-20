import base64
import os
import sys
import pandas as pd
import plotly.express as px
import streamlit as st

# Root directory path जोडण्यासाठी
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.excel_reader import process_uploaded_files
from src.pdf_reader import read_pdf
from verifier import verify_pi_against_excel

# Page Configuration
st.set_page_config(page_title="Glass PI Verification System", page_icon="🔍", layout="wide")

# ============================================================
# BASE64 LOGO HELPER FUNCTION
# ============================================================
def get_base64_image(image_path: str) -> str | None:
    """logo.png फाइल लोड करून Base64 मध्ये कन्व्हर्ट करते."""
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

logo_b64 = get_base64_image("logo.png")
if logo_b64:
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="header-logo">'
else:
    logo_html = '<div class="logo-fallback">💎</div>'

# ============================================================
# CUSTOM EXECUTIVE ENTERPRISE CSS
# ============================================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }

    div[data-testid="stMainBlockContainer"] {
        max-width: 1280px !important;
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
    }

    /* Navbar/Header Styling */
    .navbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 1.25rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15);
    }
    .navbar-brand {
        display: flex;
        align-items: center;
        gap: 16px;
    }
    .header-logo {
        height: 48px;
        width: auto;
        border-radius: 8px;
        background: #ffffff;
        padding: 4px 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .logo-fallback {
        background: #2563eb;
        color: white;
        font-size: 20px;
        padding: 8px 12px;
        border-radius: 8px;
    }
    .nav-title {
        color: #ffffff !important;
        font-size: 22px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        letter-spacing: -0.5px;
    }
    .nav-subtitle {
        color: #94a3b8 !important;
        font-size: 13px !important;
        margin-top: 2px !important;
        font-weight: 500;
    }

    /* Step Card Wrapper */
    .step-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
    }

    .step-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 16px;
    }
    .step-badge {
        background: #eff6ff;
        color: #2563eb;
        font-size: 11px;
        font-weight: 800;
        padding: 4px 10px;
        border-radius: 20px;
        border: 1px solid #bfdbfe;
        letter-spacing: 0.5px;
    }
    .step-title-text {
        color: #0f172a;
        font-size: 16px;
        font-weight: 700;
    }

    /* Alert / Information Notice Box */
    .info-notice {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-radius: 10px;
        padding: 16px 20px;
        color: #0369a1;
        font-size: 14px;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    /* Streamlit Button Restyling */
    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        padding: 0.5rem 1.2rem !important;
        transition: all 0.2s ease !important;
    }

    div.stButton > button[kind="primary"] {
        background: #2563eb !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.15) !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background: #1d4ed8 !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.25) !important;
    }

    /* Hide Header Anchor Elements */
    [data-testid="stHeader"] { display: none; }
    </style>
""", unsafe_allow_html=True)

# Initialize Session State Keys
if "excel_uploader_key" not in st.session_state:
    st.session_state["excel_uploader_key"] = 0

if "pdf_uploader_key" not in st.session_state:
    st.session_state["pdf_uploader_key"] = 0

# ============================================================
# WEBSITE HEADER WITH LOGO
# ============================================================
st.markdown(f"""
    <div class="navbar">
        <div class="navbar-brand">
            {logo_html}
            <div>
                <h1 class="nav-title">Glass PI Verification System</h1>
                <p class="nav-subtitle">Automated BOQ vs PI Data Matching & Reconciliation</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# ============================================================
# STEP 1 : Extract Data From Excel Sheets
# ============================================================
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-badge">STEP 1</span>
            <span class="step-title-text">Extract Data From Excel Sheets</span>
        </div>
""", unsafe_allow_html=True)

uploaded_excel_files = st.file_uploader(
    "Upload Excel File(s) (.xlsx, .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key=f"excel_uploader_{st.session_state['excel_uploader_key']}"
)

col_ex1, col_ex2, _ = st.columns([0.18, 0.18, 0.64])

with col_ex1:
    if st.button("📊 EXTRACT EXCEL", type="primary", use_container_width=True):
        if uploaded_excel_files:
            with st.spinner("Extracting Excel BOQ Data..."):
                excel_df = process_uploaded_files(uploaded_excel_files)
                if excel_df is not None and not excel_df.empty:
                    st.session_state["excel_df"] = excel_df
                    st.success(f"Successfully extracted {len(excel_df)} rows from Excel!")
                else:
                    st.error("Could not extract valid data from Excel file(s).")
        else:
            st.warning("Please upload Excel file(s) first.")

with col_ex2:
    if st.button("CLEAR EXCEL", use_container_width=True):
        st.session_state["excel_uploader_key"] += 1
        st.session_state.pop("excel_df", None)
        st.session_state.pop("verification_df", None)
        st.rerun()

if "excel_df" in st.session_state and not st.session_state["excel_df"].empty:
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📄 View Extracted Excel BOQ Data", expanded=True):
        st.dataframe(st.session_state["excel_df"], use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# STEP 2 : Extract Data From PI PDF Files
# ============================================================
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-badge">STEP 2</span>
            <span class="step-title-text">Extract Data From PI PDF Files</span>
        </div>
""", unsafe_allow_html=True)

uploaded_pdf_files = st.file_uploader(
    "Upload PI PDF File(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state['pdf_uploader_key']}"
)

col_pdf1, col_pdf2, _ = st.columns([0.18, 0.18, 0.64])

with col_pdf1:
    if st.button("🔍 EXTRACT PDF", type="primary", use_container_width=True):
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
                st.success(f"Successfully extracted {len(combined_pdf_df)} rows from PDF(s)!")
            else:
                st.error("Could not extract valid data from PDF file(s).")
        else:
            st.warning("Please upload PDF file(s) first.")

with col_pdf2:
    if st.button("CLEAR PDF", use_container_width=True):
        st.session_state["pdf_uploader_key"] += 1
        st.session_state.pop("pdf_df", None)
        st.session_state.pop("verification_df", None)
        st.rerun()

if "pdf_df" in st.session_state and not st.session_state["pdf_df"].empty:
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📄 View Extracted PDF Data", expanded=True):
        st.dataframe(st.session_state["pdf_df"], use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# STEP 3 : Verify PDF Data Against Excel BOQ
# ============================================================
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-badge">STEP 3</span>
            <span class="step-title-text">Verify PDF Data Against Excel BOQ</span>
        </div>
""", unsafe_allow_html=True)

excel_ready = "excel_df" in st.session_state and not st.session_state["excel_df"].empty
pdf_ready = "pdf_df" in st.session_state and not st.session_state["pdf_df"].empty

if not excel_ready or not pdf_ready:
    st.markdown("""
        <div class="info-notice">
            💡 Please extract data from <b>STEP 1 (Excel)</b> and <b>STEP 2 (PDF)</b> first.
        </div>
    """, unsafe_allow_html=True)
else:
    col_v1, _ = st.columns([0.22, 0.78])
    with col_v1:
        if st.button("⚡ RUN VERIFICATION", type="primary", use_container_width=True):
            with st.spinner("Matching and Verifying Data..."):
                v_df = verify_pi_against_excel(
                    st.session_state["excel_df"], 
                    st.session_state["pdf_df"]
                )
                st.session_state["verification_df"] = v_df

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# VERIFICATION SUMMARY & ANALYTICS DASHBOARD
# ============================================================
if 'verification_df' in st.session_state and not st.session_state["verification_df"].empty:
    df_res = st.session_state["verification_df"]

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
        <h3 style="color: #0f172a; font-size: 20px; font-weight: 700; margin-bottom: 20px;">
            📊 Verification Summary & Analytics
        </h3>
    """, unsafe_allow_html=True)

    total_items = len(df_res)
    exact_matches = len(df_res[df_res["Verification Status"] == "✅ MATCHED"])
    mismatches = total_items - exact_matches
    match_percentage = (exact_matches / total_items) * 100 if total_items > 0 else 0

    col1, col2, col3, _ = st.columns([1.2, 1.2, 1.2, 2.4])
    
    with col1:
        st.markdown(f"""
            <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.02);">
                <p style="color: #64748b; font-size: 11px; font-weight: 700; margin: 0; text-transform: uppercase; letter-spacing: 0.5px;">TOTAL BOQ ITEMS</p>
                <h2 style="color: #0f172a; font-size: 28px; font-weight: 800; margin: 4px 0 0 0;">{total_items}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 16px 20px;">
                <p style="color: #166534; font-size: 11px; font-weight: 700; margin: 0; text-transform: uppercase; letter-spacing: 0.5px;">EXACT MATCHES</p>
                <div style="display: flex; align-items: baseline; gap: 8px;">
                    <h2 style="color: #15803d; font-size: 28px; font-weight: 800; margin: 4px 0 0 0;">{exact_matches}</h2>
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
            <div style="background: {bg_color}; border: 1px solid {border_color}; border-radius: 12px; padding: 16px 20px;">
                <p style="color: {text_color}; font-size: 11px; font-weight: 700; margin: 0; text-transform: uppercase; letter-spacing: 0.5px;">MISMATCHES / MISSING</p>
                <h2 style="color: {num_color}; font-size: 28px; font-weight: 800; margin: 4px 0 0 0;">{mismatches}</h2>
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
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 20px; text-align: left;">
                    <p style="color: #15803d; font-size: 14px; font-weight: 600; margin: 0;">🎉 Perfect 100% Match! No mismatches to display.</p>
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
    st.markdown("<h4 style='color: #0f172a; font-weight: 700;'>📑 Detailed Verification Results</h4>", unsafe_allow_html=True)
    
    col_f1, _ = st.columns([2, 3])
    with col_f1:
        selected_status = st.selectbox(
            "Filter Results by Status:", 
            options=["ALL"] + list(df_res["Verification Status"].unique())
        )

    display_df = df_res[df_res["Verification Status"] == selected_status] if selected_status != "ALL" else df_res
    st.dataframe(display_df, use_container_width=True, height=400)
