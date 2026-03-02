# Startup Financial Model Generator

This repository contains a professional Streamlit web application that generates a fully integrated 3-statement financial model (Income Statement, Balance Sheet, Cash Flow Statement, Capex Schedule, and Debt Schedule) in a downloadable Excel file.

## Features
- **Interactive UI**: Enter real-time assumptions for Revenue, SG&A, Capex, Debt, and other key inputs.
- **Dynamic Excel Generation**: Leveraging `openpyxl` to build beautifully formatted Excel files on the fly.
- **Instant Download**: Provides the user with a direct `.xlsx` download button for the generated model.

## File Structure
- `app.py`: The main Streamlit web application.
- `generator.py`: The core financial modeling engine that constructs the Excel file based on assumptions.
- `requirements.txt`: Python dependencies.

## How to Run Locally

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the Streamlit application:
   ```bash
   streamlit run app.py
   ```

3. Open your browser to the local URL provided (usually `http://localhost:8501`).
