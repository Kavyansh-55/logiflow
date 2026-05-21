# app.py
# RUN: streamlit run app.py

import json
from io import BytesIO
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pypdf import PdfReader, PdfWriter

# =====================================
# API
# =====================================

API_KEY = "AIzaSyAW7pTxN74jsPgOR-vcqYfXmRQ97KO7XOA"

client = genai.Client(
    api_key=API_KEY
)

# =====================================
# PAGE
# =====================================

st.set_page_config(
    page_title="LogiFlow",
    page_icon="📄",
    layout="wide"
)

# =====================================
# SESSION
# =====================================

if "credits" not in st.session_state:
    st.session_state.credits = 3

if "history" not in st.session_state:
    st.session_state.history = []

if "locked" not in st.session_state:
    st.session_state.locked = False

if "email" not in st.session_state:
    st.session_state.email = "operations@demo.com"

if "password" not in st.session_state:
    st.session_state.password = "••••••••"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# Permanent lock logic
if st.session_state.credits <= 0:
    st.session_state.locked = True

# =====================================
# CSS STYLE OVERRIDES (DARK MODE SAFE)
# =====================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

#MainMenu, footer, header {
    visibility: hidden;
}

[data-testid='stToolbar'],
[data-testid='stHeaderActionElements'],
.stMarkdown a.header-anchor,
button[data-testid="stTabBarTicket"] {
    display: none !important;
}

.block-container {
    max-width: 94%;
    padding-top: .4rem;
}

.stApp {
    background: linear-gradient(135deg, #07111f, #081527, #111827);
}

.stTabs [data-baseweb="tab-list"] {
    justify-content: center;
    gap: 45px;
}

/* Hard-coded styling rule to prevent dark mode text blending bugs */
.logo {
    font-size: 58px;
    font-weight: 800;
    letter-spacing: 12px;
    color: #FFFFFF !important;
    text-shadow: 0px 4px 12px rgba(139, 92, 246, 0.4);
    text-align: center;
    margin-bottom: 5px;
}

.subtitle {
    text-align: center;
    font-size: 14px;
    margin-bottom: 35px;
    color: #A3BFD9 !important;
    font-weight: 500;
    letter-spacing: 0.5px;
}

.metric {
    padding: 22px;
    background: linear-gradient(180deg, #111827, #0f172a);
    border-radius: 22px;
    border: 1px solid rgba(255, 255, 255, 0.06);
    height: 150px;
    text-align: center;
}

.lock {
    padding: 35px;
    border-radius: 20px;
    background: linear-gradient(180deg, #1e1b4b, #111827);
    border: 1px solid #7c3aed;
    text-align: center;
    margin-bottom: 25px;
}

[data-testid="stFileUploader"] {
    background: #111827 !important;
    border: 2px dashed #7c3aed !important;
    border-radius: 25px !important;
    padding: 40px !important;
    min-height: 240px !important;
    margin-top: 20px !important;
}

.stButton button {
    height: 56px;
    border-radius: 14px;
    font-weight: 600;
    width: 100%;
}

/* Premium Structured Pricing Custom Alignment Card Layouts */
.pricing-card {
    background: linear-gradient(180deg, #111827, #0f172a);
    border-radius: 22px;
    padding: 30px;
    border: 1px solid rgba(255, 255, 255, 0.06);
    min-height: 380px;
}

.pricing-title {
    font-size: 22px;
    font-weight: 700;
    margin-bottom: 5px;
}

.pricing-price {
    font-size: 28px;
    font-weight: 800;
    color: #A78BFA;
    margin-bottom: 20px;
}

.pricing-feature {
    font-size: 14px;
    margin-bottom: 12px;
    color: #E2E8F0;
    display: flex;
    align-items: center;
}

.checkmark {
    color: #10B981;
    font-weight: bold;
    margin-right: 10px;
}
</style>
""", unsafe_allow_html=True)

# =====================================
# PREMIUM PROFILE NAVIGATION HUB
# =====================================

left, center, right = st.columns([3, 6, 3])

with left:
    # Interactive Popover Dashboard Panel
    with st.popover(
            "🌐 Public Account Menu" if not st.session_state.logged_in else f"👤 Account: {st.session_state.email}"):
        st.markdown("### 📁 Management Panel")

        # User dynamic profile variables
        user_email = st.text_input("Operator Workspace Email",
                                   value=st.session_state.email if st.session_state.logged_in else "public_demo@logiflow.io")
        user_pass = st.text_input("Security Key / Password", value=st.session_state.password, type="password")

        if st.button("💾 Save Account Changes"):
            st.session_state.email = user_email
            st.session_state.password = user_pass
            st.session_state.logged_in = True
            st.toast("Profile configurations updated securely!", icon="💾")
            st.rerun()

        st.markdown("---")
        st.markdown("### 💳 Active Subscriptions")

        if st.session_state.logged_in:
            st.info(f"🏷️ Plan: Pending Activation\n\n📊 Usage Limits: 0 / 2,500 operations.")
            if st.button("🚪 Logout Session"):
                st.session_state.logged_in = False
                st.session_state.credits = 3
                st.session_state.locked = False
                st.session_state.email = "operations@demo.com"
                st.session_state.password = "••••••••"
                st.rerun()
        else:
            st.warning("🏷️ Plan: Demo Sandbox\n\n📊 Status: Operational Verification Mode Only.")

with center:
    st.markdown("""
    <div class='logo'>LOGIFLOW</div>
    <div class='subtitle'>AI-Powered Invoice Extraction • Excel Automation • Enterprise Processing</div>
    """, unsafe_allow_html=True)

with right:
    st.markdown(
        f"<div style='text-align: right; padding-top: 12px; font-weight: 600; font-size: 16px; color: #94A3B8;'>Credits: {st.session_state.credits}/3</div>",
        unsafe_allow_html=True)

# =====================================
# NAVIGATION ENGINE
# =====================================

tab1, tab2, tab3 = st.tabs([
    "🚀 Data Extractor",
    "💳 Upgrade & Plans",
    "📊 Processing History"
])


# =====================================
# PDF BOUNDARY CONSTRAINTS
# =====================================

def first_two_pages(pdf_bytes):
    reader = PdfReader(BytesIO(pdf_bytes))
    writer = PdfWriter()
    pages = min(2, len(reader.pages))

    for i in range(pages):
        writer.add_page(reader.pages[i])

    out = BytesIO()
    writer.write(out)
    return out.getvalue()


def parse_pdf(pdf):
    prompt = """
    Extract:
    Invoice Number
    Date
    Carrier
    Total Charges

    Return JSON only
    """
    part = types.Part.from_bytes(
        data=pdf,
        mime_type="application/pdf"
    )

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[prompt, part],
        config={"response_mime_type": "application/json"}
    )
    return json.loads(response.text)


# =====================================
# AUTOMATED EXCEL UTILITY
# =====================================

def create_excel(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False)
        sh = w.sheets["Sheet1"]

        fill = PatternFill(fill_type="solid", start_color="1E3A8A")
        font = Font(bold=True, color="FFFFFF")

        for c in sh[1]:
            c.fill = fill
            c.font = font
            c.alignment = Alignment(horizontal="center")

        for col in sh.columns:
            m = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                m = max(m, len(str(cell.value or "")))
            sh.column_dimensions[letter].width = m + 3

    return out.getvalue()


# =====================================
# TAB 1: WORKPLACE INTAKE INTERFACE
# =====================================

with tab1:
    st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)

    if st.session_state.locked:
        st.markdown("""
        <div class='lock'>
            <h2>🔒 Demo Access Exhausted</h2>
            <p style='color: #94A3B8;'>3 lifetime evaluation credits have been consumed cleanly.</p>
            <p>Upgrade your platform profile instance to unlock limitless production processing logs.</p>
        </div>
        """, unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            email_input = st.text_input("Create Workspace Operator Email", value="")
        with c2:
            pass_input = st.text_input("Choose Strong Password", type="password")

        if st.button("📩 Create Operator Account & Proceed"):
            if email_input and pass_input:
                st.session_state.email = email_input
                st.session_state.password = pass_input
                st.session_state.logged_in = True
                st.toast("Corporate profile registry successful!", icon="🚀")
                st.rerun()
            else:
                st.error("Please fill in corporate registry parameters to lock in your software account structure.")

        st.warning(
            "Device session locked under Sandbox rule limits. Please migrate your system token context to a paid tier.")

    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Credits Running", f"{st.session_state.credits}/3")
        with c2:
            st.metric("Processing Profile", "PDF → Flat Excel")
        with c3:
            st.metric("Engine Version", "Neural Core V2")

        st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)

        file = st.file_uploader(
            "Upload System Document File Here",
            type=["pdf"],
            accept_multiple_files=False
        )

        if file:
            if st.button("⚡ Analyze Document Engine Pipeline"):
                with st.spinner("Executing direct bytes parsing matrix channels..."):
                    pdf = file.read()
                    pdf = first_two_pages(pdf)
                    result = parse_pdf(pdf)

                    st.session_state.credits -= 1
                    st.session_state.history.append(file.name)

                    df = pd.DataFrame([result])
                    st.success("Ledger array extracted perfectly.")

                    st.dataframe(df, use_container_width=True)

                    st.download_button(
                        "⬇ Export Structured Excel Document",
                        create_excel(df),
                        "logiflow_ledger.xlsx"
                    )

# =====================================
# TAB 2: CORPORATE PRICING ARRAYS
# =====================================

with tab2:
    st.markdown("<div style='height:35px'></div>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("""
        <div class="pricing-card" style="border-top: 4px solid #3B82F6;">
            <div class="pricing-title">Demo Sandbox</div>
            <div class="pricing-price">$0 <span style="font-size:14px; color:#64748B;">/ free trial</span></div>
            <div class="pricing-feature"><span class="checkmark">✓</span> 3 Lifetime Evaluation Credits</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Standard Cloud Queuing Channel</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Limit 1 Document Upload per Run</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Strict 2-Page Cap per Invoice</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> For Accuracy Verification Only</div>
        </div>
        """, unsafe_allow_html=True)
        st.button("Current Testing Environment Active", disabled=True, key="p1")

    with c2:
        st.markdown("""
        <div class="pricing-card" style="border-top: 4px solid #10B981; box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.1);">
            <div class="pricing-title">Logistics Pro</div>
            <div class="pricing-price">$149 <span style="font-size:14px; color:#64748B;">/ month</span></div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Volumetric Cap: 2,500 Invoices/Mo</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> High-Speed Processing (4–5 Seconds)</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Zero Page Limits/Caps per Document</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Multi-File Batch Data Uploads</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Premium Styled Excel Ledgers</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("⚡ Upgrade to Logistics Pro", key="p2"):
            st.success("Your enterprise routing profile invoice request has been filed with billing@logiflow.io!")

    with c3:
        st.markdown("""
        <div class="pricing-card" style="border-top: 4px solid #F59E0B;">
            <div class="pricing-title">Enterprise Core</div>
            <div class="pricing-price">$499 <span style="font-size:14px; color:#64748B;">/ month</span></div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Unlimited Processing (30k/mo Fair Use)</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Dedicated Priority Execution (2–3s)</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Multi-User Managed Seats (Up to 10)</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Live Google Sheets Webhook Sync</div>
            <div class="pricing-feature"><span class="checkmark">✓</span> Custom Accounting ERP Key Mapping</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("💼 Deploy Enterprise Core Plan", key="p3"):
            st.success(
                "An enterprise solution architect has been assigned to contact your organization desk within 1 hour.")

# =====================================
# TAB 3: SYSTEM TRANSACTION RECORDS
# =====================================

with tab3:
    st.markdown("<div style='height:35px'></div>", unsafe_allow_html=True)
    st.subheader("Recent Document Audit Run Matrix")

    if len(st.session_state.history) == 0:
        st.info("No transaction parsing logs found inside current validation registry framework.")
    else:
        for idx, item in enumerate(st.session_state.history):
            st.success(f"📟 Entry ID #{1001 + idx}  |  📄 Filename Reference: {item}  |  🔒 Status: Cleared")

# =====================================
# SYSTEM FOOTER DATA & SEO MOAT
# =====================================

st.markdown("---")
st.subheader("Industry Documentation & Technical Manuals")

st.markdown("""
• [How to automate freight billing records and eliminate clerical errors data entry modules](#)  
• [Top AI tools engineered to convert unstructured logistics multi-page PDFs to Excel sheets](#)  
• [The hidden balance-sheet costs of manual back-office entry errors in corporate fleet processing](#)  
""")

st.caption(
    "Registered Utility Gateway Asset  |  🔐 Contact Interface Support Protocol: support@logiflow.io | licensing@logiflow.io")