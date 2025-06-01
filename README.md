# Personal Finance Analysis Dashboard

## Overview

This is an interactive web application built with Streamlit that helps users analyze their personal spending patterns from credit card statements in CSV format. It provides visualizations, spending advice, and options to export reports.

## Features

*   **CSV Upload:** Securely upload your credit card statement (CSV file) directly in the app.
*   **Data Processing:** Automatically cleans and preprocesses financial data, parsing dates and amounts.
*   **Transaction Categorization:** Assigns spending categories to transactions based on keywords found in their descriptions (e.g., "餐飲", "交通").
*   **Interactive Dashboard:**
    *   **Summary Statistics:** View total spending, total number of transactions, and average transaction value.
    *   **Spending by Category:** Interactive bar chart showing spending distribution across different categories.
    *   **Spending by Month:** Interactive line chart illustrating spending trends over time.
    *   **Outlier Transactions:** Table highlighting single transactions exceeding a predefined threshold (e.g., NT$5,000).
*   **Personalized Spending Advice:** Generates simple advice based on your spending habits (e.g., highest spending category, changes in monthly spending).
*   **Report Export:**
    *   **Excel:** Download a multi-sheet Excel file containing raw transactions, spending by category, and spending by month.
    *   **PDF:** Download a comprehensive PDF report including summary statistics, charts, outlier transactions, and spending advice. (Supports Chinese characters with appropriate font).

## Requirements

*   Python (3.7+ recommended)
*   The following Python libraries:
    *   streamlit
    *   pandas
    *   plotly
    *   fpdf2
    *   xlsxwriter
    *   kaleido

## Setup and Installation

1.  **Clone the repository (if applicable) or download the `app.py` file.**

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3.  **Install dependencies:**
    ```bash
    pip install streamlit pandas plotly fpdf2 xlsxwriter kaleido
    ```
    (Alternatively, if a `requirements.txt` file is provided: `pip install -r requirements.txt`)

## Running the Application

1.  Navigate to the directory where `app.py` is located.
2.  Run the following command in your terminal:
    ```bash
    streamlit run app.py
    ```
3.  The application should open in your web browser.

## CSV File Format

The application expects a CSV file with the following columns:

*   `消費日`: Transaction date (e.g., YYYY/MM/DD or YYYY-MM-DD).
*   `交易說明`: Description of the transaction. Keywords from this field are used for categorization.
*   `新臺幣金額`: Transaction amount (should be numeric; commas as thousands separators are handled).

Other columns can be present but will be ignored by the core analysis.

## PDF Export Font Note

For optimal PDF export with Chinese characters, the application attempts to use the `fireflysung.ttf` font. If this font is not available in the environment, it will fall back to Arial, and Chinese characters may not render correctly in the PDF. You might need to install this font or ensure the TTF file is accessible by the application.

## Future Enhancements (Ideas)

*   User-defined categories and keyword mapping.
*   More sophisticated spending advice and trend analysis.
*   Support for multiple accounts or currencies.
*   Data persistence (e.g., saving data to a database).
*   Direct deployment to cloud platforms.
```
