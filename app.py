import streamlit as st
import io
from pdf_generator import FinancialModelData, PDFGenerator

st.set_page_config(page_title="Startup Financial Model", page_icon="📊", layout="wide")

# Custom CSS for a professional look
st.markdown("""
<style>
    /* Add some professional padding and colors */
    .stButton>button {
        background-color: #1F4E78;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #153654;
        color: white;
    }
    h1, h2, h3 {
        color: #1F4E78;
    }
    .stDownloadButton>button {
        background-color: #28a745;
        color: white;
        font-weight: bold;
    }
    .stDownloadButton>button:hover {
        background-color: #218838;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 3-Statement Financial Model Generator")
st.markdown("""
Welcome to the professional startup financial model generator. 
Adjust the assumptions below, and the app will generate a clean and professional PDF report containing your Income Statement, Balance Sheet, Cash Flow Statement, Capex Schedule, and Debt Schedule.
""")
st.divider()

user_assumptions = {
    "REVENUE ASSUMPTIONS": {},
    "SG&A ASSUMPTIONS": {},
    "CAPEX & DEPRECIATION": {},
    "DEBT SCHEDULE": {},
    "OTHER": {}
}

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Revenue Assumptions")
    with st.container(border=True):
        user_assumptions["REVENUE ASSUMPTIONS"]["Year 1 Revenue (INR)"] = st.number_input("Year 1 Revenue (INR)", value=5400000.0, step=100000.0, format="%.2f")
        user_assumptions["REVENUE ASSUMPTIONS"]["Revenue Growth Rate - Y2"] = st.number_input("Revenue Growth Rate - Y2", value=0.50, step=0.01, format="%.2f")
        user_assumptions["REVENUE ASSUMPTIONS"]["Revenue Growth Rate - Y3"] = st.number_input("Revenue Growth Rate - Y3", value=0.35, step=0.01, format="%.2f")
        user_assumptions["REVENUE ASSUMPTIONS"]["Revenue Growth Rate - Y4"] = st.number_input("Revenue Growth Rate - Y4", value=0.25, step=0.01, format="%.2f")
        user_assumptions["REVENUE ASSUMPTIONS"]["Revenue Growth Rate - Y5"] = st.number_input("Revenue Growth Rate - Y5", value=0.15, step=0.01, format="%.2f")
        user_assumptions["REVENUE ASSUMPTIONS"]["Gross Margin Change"] = st.number_input("Gross Margin Change", value=0.07, step=0.01, format="%.2f")
        user_assumptions["REVENUE ASSUMPTIONS"]["Gross Margin %"] = st.number_input("Gross Margin %", value=0.60, step=0.01, format="%.2f")

with col2:
    st.subheader("SG&A Assumptions")
    with st.container(border=True):
        user_assumptions["SG&A ASSUMPTIONS"]["Sales & Marketing (% Revenue)"] = st.number_input("Sales & Marketing (% Revenue)", value=0.20, step=0.01, format="%.2f")
        user_assumptions["SG&A ASSUMPTIONS"]["General & Admin (% Revenue)"] = st.number_input("General & Admin (% Revenue)", value=0.10, step=0.01, format="%.2f")
        user_assumptions["SG&A ASSUMPTIONS"]["R&D (% Revenue)"] = st.number_input("R&D (% Revenue)", value=0.08, step=0.01, format="%.2f")

    st.subheader("Capex & Depreciation")
    with st.container(border=True):
        user_assumptions["CAPEX & DEPRECIATION"]["Capex (% Revenue)"] = st.number_input("Capex (% Revenue)", value=0.05, step=0.01, format="%.2f")
        user_assumptions["CAPEX & DEPRECIATION"]["Useful Life (years)"] = st.number_input("Useful Life (years)", value=5.0, step=1.0, format="%.1f")
        user_assumptions["CAPEX & DEPRECIATION"]["Beginning PP&E, Net (INR)"] = st.number_input("Beginning PP&E, Net (INR)", value=1400000.0, step=100000.0, format="%.2f")

with col3:
    st.subheader("Debt Schedule")
    with st.container(border=True):
        user_assumptions["DEBT SCHEDULE"]["Beginning Debt (INR)"] = st.number_input("Beginning Debt (INR)", value=600000.0, step=50000.0, format="%.2f")
        user_assumptions["DEBT SCHEDULE"]["Annual Interest Rate"] = st.number_input("Annual Interest Rate", value=0.07, step=0.01, format="%.2f")
        user_assumptions["DEBT SCHEDULE"]["Annual Debt Repayment (INR)"] = st.number_input("Annual Debt Repayment (INR)", value=100000.0, step=10000.0, format="%.2f")

    st.subheader("Other General")
    with st.container(border=True):
        user_assumptions["OTHER"]["Tax Rate"] = st.number_input("Tax Rate", value=0.30, step=0.01, format="%.2f")
        user_assumptions["OTHER"]["Beginning Cash (INR)"] = st.number_input("Beginning Cash (INR)", value=800000.0, step=50000.0, format="%.2f")
        user_assumptions["OTHER"]["Beginning Equity (INR)"] = st.number_input("Beginning Equity (INR)", value=2000000.0, step=100000.0, format="%.2f")

st.divider()

col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])

with col_btn2:
    st.markdown("<h3 style='text-align: center'>Generate Your PDF Report</h3>", unsafe_allow_html=True)
    if st.button("Generate Financial Model", use_container_width=True):
        with st.spinner("Generating financial model PDF..."):
            model_data = FinancialModelData(custom_assumptions=user_assumptions)
            pdf = PDFGenerator(model_data)
            
            output = io.BytesIO()
            pdf.build_pdf(output)
            output.seek(0)
            
            st.success("Financial Model PDF generated successfully!")
            
            st.download_button(
                label="📥 Download PDF Report",
                data=output,
                file_name="financial_model_output.pdf",
                mime="application/pdf",
                use_container_width=True
            )

st.markdown('''
---
<div style='text-align: center; color: gray; font-size: small'>
Built with Streamlit • 3-Statement Financial Modeling
</div>
''', unsafe_allow_html=True)
