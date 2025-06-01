import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF
import plotly.io as pio

# --- Analysis Functions ---

def calculate_total_spending(df):
    """Calculates the sum of the '新臺幣金額' column."""
    if '新臺幣金額' in df.columns and not df.empty:
        return df['新臺幣金額'].sum()
    return 0.0

def calculate_spending_by_category(df):
    """Groups by 'Category' and sums '新臺幣金額', sorted by spending."""
    if 'Category' in df.columns and '新臺幣金額' in df.columns and not df.empty:
        return df.groupby('Category')['新臺幣金額'].sum().sort_values(ascending=False)
    return pd.Series(name='新臺幣金額') # Return empty Series with correct name if columns missing

def calculate_spending_by_month(df):
    """Groups by 'YearMonth' and sums '新臺幣金額', sorted by 'YearMonth'."""
    if 'YearMonth' in df.columns and '新臺幣金額' in df.columns and not df.empty:
        # Ensure YearMonth is in a sortable format if it's not already (e.g. Period)
        # If YearMonth is text, convert to Period or Datetime for correct sorting
        temp_df = df.copy()
        if not pd.api.types.is_period_dtype(temp_df['YearMonth']) and not pd.api.types.is_datetime64_any_dtype(temp_df['YearMonth']):
             temp_df['YearMonth'] = pd.to_datetime(temp_df['YearMonth'], errors='coerce').dt.to_period('M')

        return temp_df.groupby('YearMonth')['新臺幣金額'].sum().sort_index()
    return pd.Series(name='新臺幣金額') # Return empty Series

def identify_outlier_transactions(df, threshold=5000):
    """Filters for transactions where '新臺幣金額' > threshold."""
    if '新臺幣金額' in df.columns and not df.empty:
        outliers = df[df['新臺幣金額'] > threshold]
        # Ensure relevant columns exist before selecting them
        relevant_cols = ['消費日', '交易說明', '新臺幣金額', 'Category']
        existing_cols = [col for col in relevant_cols if col in outliers.columns]
        return outliers[existing_cols].sort_values(by='新臺幣金額', ascending=False)
    return pd.DataFrame() # Return empty DataFrame

def generate_spending_advice(df, spending_by_category, spending_by_month):
    """Generates a list of spending advice strings based on the data."""
    advice_list = []
    default_advice = "目前尚無足夠的資料來產生具體的消費建議。"

    # Advice 1: Highest Spending Category (Overall)
    if spending_by_category is not None and not spending_by_category.empty:
        top_category_name = spending_by_category.index[0]
        top_category_amount = spending_by_category.iloc[0]
        advice_list.append(f"您在「{top_category_name}」類別的總支出最高，金額為 NT$ {top_category_amount:,.0f}。請特別留意此類別的開銷。")

    # Advice 2: Monthly Spending Change (Comparison to Average)
    if spending_by_month is not None and len(spending_by_month) >= 2:
        # Ensure index is sorted for correct latest month identification
        spending_by_month_sorted = spending_by_month.sort_index()
        latest_month_spending = spending_by_month_sorted.iloc[-1]
        latest_month_name = spending_by_month_sorted.index[-1] # Could be Period object

        if len(spending_by_month_sorted) > 1:
            previous_months_spending = spending_by_month_sorted.iloc[:-1]
            avg_monthly_spending = previous_months_spending.mean()

            if avg_monthly_spending > 0: # Avoid division by zero or meaningless comparison
                percentage_diff = ((latest_month_spending - avg_monthly_spending) / avg_monthly_spending) * 100
                latest_month_str = str(latest_month_name) # Convert Period to string for display

                if percentage_diff > 20: # More than 20% increase
                    advice_list.append(f"您最近一個月 ({latest_month_str}) 的總支出 (NT$ {latest_month_spending:,.0f}) 比前幾個月的平均支出 (NT$ {avg_monthly_spending:,.0f}) 高出 {percentage_diff:.0f}%。建議檢視是否有非預期的大額支出。")
                elif percentage_diff < -20: # More than 20% decrease
                    advice_list.append(f"您最近一個月 ({latest_month_str}) 的總支出 (NT$ {latest_month_spending:,.0f}) 比前幾個月的平均支出 (NT$ {avg_monthly_spending:,.0f}) 低了 {abs(percentage_diff):.0f}%。這可能是個好現象！")

    # Advice 3: High spending on "Miscellaneous"
    if spending_by_category is not None and "雜項" in spending_by_category.index:
        misc_spending = spending_by_category.loc["雜項"]
        total_spending_val = calculate_total_spending(df) # Recalculate or pass as arg
        if total_spending_val > 0 and (misc_spending / total_spending_val) > 0.15: # If misc is > 15% of total
             advice_list.append(f"「雜項」類別的支出佔比較高 (NT$ {misc_spending:,.0f})。嘗試更詳細地記錄這些消費，以便更好地追蹤資金流向。")

    if not advice_list:
        return [default_advice]
    return advice_list

def to_excel(df_main, series_by_category, series_by_month):
    """Exports DataFrames to an in-memory Excel file."""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # Main transactions sheet
    # Select and rename columns for the main transactions sheet for better readability
    df_export = df_main.copy()
    if '消費日' in df_export.columns:
        df_export['消費日'] = df_export['消費日'].dt.strftime('%Y-%m-%d') # Format date as string

    # Define a subset of columns and their desired names for export
    export_columns_map = {
        '消費日': '消費日期 (Date)',
        '交易說明': '交易說明 (Description)',
        '新臺幣金額': '金額 (Amount NTD)',
        'Category': '類別 (Category)',
        'YearMonth': '消費月份 (YearMonth)',
        'Year': '年份 (Year)',
        'Month': '月份 (Month)'
    }
    # Filter out columns not present in df_export and select existing ones
    columns_to_export = {k: v for k, v in export_columns_map.items() if k in df_export.columns}
    df_to_sheet = df_export[list(columns_to_export.keys())].rename(columns=columns_to_export)
    df_to_sheet.to_excel(writer, sheet_name='所有交易 (All Transactions)', index=False)

    # Spending by Category sheet
    if series_by_category is not None and not series_by_category.empty:
        sbc_df = series_by_category.reset_index()
        sbc_df.columns = ['類別 (Category)', '總支出 (Total Spending NTD)']
        sbc_df.to_excel(writer, sheet_name='各類別支出 (Spending by Category)', index=False)

    # Spending by Month sheet
    if series_by_month is not None and not series_by_month.empty:
        sbm_df = series_by_month.reset_index()
        sbm_df.columns = ['消費月份 (YearMonth)', '總支出 (Total Spending NTD)']
        if '消費月份 (YearMonth)' in sbm_df.columns: # Ensure column exists
             sbm_df['消費月份 (YearMonth)'] = sbm_df['消費月份 (YearMonth)'].astype(str) # Convert Period to string
        sbm_df.to_excel(writer, sheet_name='各月份支出 (Spending by Month)', index=False)

    writer.close() # Saves the workbook
    processed_data = output.getvalue()
    return processed_data

def create_pdf_report(df_main, total_spending_val, avg_transaction_val, num_transactions_val,
                        sbc_df_for_pdf, sbm_df_for_pdf, outlier_transactions_df,
                        advice_list_for_pdf, fig_category_obj, fig_month_obj):
    """Creates a PDF report with summary, charts, tables, and advice."""
    pdf = FPDF()
    pdf.add_page()

    font_path_available = False
    try:
        # Attempt to add a CJK font. This requires the font file to be discoverable.
        # For example, if 'fireflysung.ttf' is in the same directory or a known font path.
        pdf.add_font('fireflysung', '', 'fireflysung.ttf', uni=True)
        pdf.set_font('fireflysung', '', 12)
        font_path_available = True
    except RuntimeError as e:
        # Fallback to Arial if CJK font is not found
        pdf.set_font('Arial', '', 12)
        # This st.warning will only work if called from within a Streamlit execution context.
        # Ideally, font checking/warning should be handled in the Streamlit part of the code.
        # For now, following prompt to put it here.
        st.warning(f"CJK font (fireflysung) not found or error loading: {e}. PDF text may not render Chinese characters correctly. Using Arial as fallback.")

    # Report Title
    pdf.set_font_size(16)
    pdf.cell(0, 10, '個人財務報告 (Personal Finance Report)', ln=True, align='C')
    pdf.set_font_size(12) # Reset font size

    # Summary Statistics
    pdf.ln(10)
    pdf.set_font_size(14)
    pdf.cell(0, 10, '總覽 (Summary Statistics)', ln=True)
    pdf.set_font_size(12)
    pdf.cell(0, 8, f"總支出 (Total Spending): NT$ {total_spending_val:,.2f}", ln=True)
    pdf.cell(0, 8, f"總交易筆數 (Total Transactions): {num_transactions_val}", ln=True)
    pdf.cell(0, 8, f"平均交易金額 (Average Transaction): NT$ {avg_transaction_val:,.2f}", ln=True)

    # Charts as Images
    if fig_category_obj:
        try:
            img_bytes_category = pio.to_image(fig_category_obj, format='png', width=800, height=600, scale=2)
            pdf.add_page()
            pdf.set_font_size(14)
            pdf.cell(0, 10, '各類別消費分佈 (Spending by Category Chart)', ln=True)
            pdf.set_font_size(12)
            pdf.image(BytesIO(img_bytes_category), x=10, y=None, w=180) # Adjust width as needed
            pdf.ln(5)
        except Exception as e:
            pdf.ln(5)
            pdf.cell(0, 8, f"無法生成類別消費圖表: {e}", ln=True)


    if fig_month_obj:
        try:
            img_bytes_month = pio.to_image(fig_month_obj, format='png', width=800, height=600, scale=2)
            pdf.add_page()
            pdf.set_font_size(14)
            pdf.cell(0, 10, '每月消費趨勢 (Spending by Month Chart)', ln=True)
            pdf.set_font_size(12)
            pdf.image(BytesIO(img_bytes_month), x=10, y=None, w=180) # Adjust width
            pdf.ln(5)
        except Exception as e:
            pdf.ln(5)
            pdf.cell(0, 8, f"無法生成月份消費圖表: {e}", ln=True)

    # Outlier Transactions Table
    pdf.add_page()
    pdf.set_font_size(14)
    pdf.cell(0, 10, '高額交易明細 (High-Value Transactions) (前10筆)', ln=True)
    pdf.set_font_size(10) # Smaller font for table

    if outlier_transactions_df is not None and not outlier_transactions_df.empty:
        col_widths = [25, 85, 30, 40] # Adjusted for A4 width (approx 190 units available for cells)

        # Table Header
        pdf.set_font(pdf.font_family, 'B', pdf.font_size_pt) # Bold for header
        pdf.cell(col_widths[0], 8, '日期 (Date)', border=1, align='C')
        pdf.cell(col_widths[1], 8, '說明 (Description)', border=1, align='C')
        pdf.cell(col_widths[2], 8, '金額 (Amount)', border=1, align='C')
        pdf.cell(col_widths[3], 8, '類別 (Category)', border=1, align='C')
        pdf.ln()
        pdf.set_font(pdf.font_family, '', pdf.font_size_pt) # Regular for data

        for _, row in outlier_transactions_df.head(10).iterrows(): # Limit to 10 rows for PDF
            date_str = str(row['消費日'].date()) if pd.notnull(row['消費日']) else ''
            desc_str = str(row['交易說明'])[:28] if font_path_available else str(row['交易說明'])[:45] # Shorter for CJK, longer for Arial
            amount_str = f"{row['新臺幣金額']:,.0f}"
            cat_str = str(row['Category'])[:10] if font_path_available else str(row['Category'])[:18]

            pdf.cell(col_widths[0], 8, date_str, border=1)
            pdf.cell(col_widths[1], 8, desc_str, border=1)
            pdf.cell(col_widths[2], 8, amount_str, border=1, align='R')
            pdf.cell(col_widths[3], 8, cat_str, border=1)
            pdf.ln()
    else:
        pdf.cell(0, 8, '未發現高於 NT$ 5,000 的交易。 (No transactions found above NT$ 5,000.)', ln=True)
    pdf.set_font_size(12) # Reset font size

    # Spending Advice
    pdf.add_page()
    pdf.set_font_size(14)
    pdf.cell(0, 10, '消費建議 (Spending Advice)', ln=True)
    pdf.set_font_size(12)
    if advice_list_for_pdf:
        for advice in advice_list_for_pdf:
            if advice == "目前尚無足夠的資料來產生具體的消費建議。": # Default message
                 pdf.multi_cell(0, 8, advice, ln=True) # Use multi_cell for long advice
            else:
                 pdf.multi_cell(0, 8, f"- {advice}", ln=True)

    return pdf.output(dest='S').encode('latin-1') # Output as bytes string

# --- Main Application ---

def main():
    st.set_page_config(layout="wide", page_title="個人理財儀表板") # Wide layout and page title

    st.title("💳 個人理財儀表板 (Personal Finance Dashboard)")
    st.markdown("""
    歡迎使用您的個人理財儀表板！請上傳您的信用卡帳單 CSV 檔案，以進行消費分析、趨勢觀察及取得理財建議。
    請確保 CSV 檔案包含以下欄位：`消費日` (日期格式), `交易說明` (文字), `新臺幣金額` (數值)。
    """)

    # --- Sidebar Setup ---
    st.sidebar.header("⚙️ 控制面板 (Control Panel)")
    uploaded_file = st.sidebar.file_uploader("請上傳您的信用卡帳單 CSV 檔案", type=["csv", "txt"]) # Allow txt for flexibility

    # --- About App Section in Sidebar ---
    with st.sidebar.expander("ℹ️ 關於此應用 (About this App)", expanded=False):
        st.info("這是一個使用 Streamlit 建立的互動式個人財務儀表板。")
        st.markdown("""
        **如何在本機執行:**
        1. 確認已安裝 Python 和 pip。
        2. 安裝必要的套件: `pip install streamlit pandas plotly fpdf2 xlsxwriter kaleido`
        3. 將此腳本儲存為 `app.py`。
        4. 在終端機中執行: `streamlit run app.py`

        **部署到雲端:**
        若要讓應用程式無需手動執行腳本即可持續訪問，可以將其部署到:
        - Streamlit Community Cloud
        - Heroku
        - AWS (EC2, Elastic Beanstalk)
        - Google Cloud (App Engine, Cloud Run)
        - 等其他雲端平台。
        """)

    # --- Main content area ---
    if uploaded_file is not None:
        df = None # Initialize df
        try:
            # More robust CSV reading
            df = pd.read_csv(uploaded_file, thousands=',')

            # Basic validation of essential columns
            required_cols = {'消費日', '新臺幣金額', '交易說明'}
            missing_cols = required_cols - set(df.columns)
            if missing_cols:
                st.error(f"上傳的 CSV 檔案缺少必要的欄位：{', '.join(missing_cols)}。請確認檔案內容。")
                st.stop() # Halt execution if critical columns are missing

            # Attempt to parse dates, can be error-prone
            try:
                df['消費日'] = pd.to_datetime(df['消費日'])
            except Exception as e:
                st.error(f"解析 '消費日' 欄位時發生錯誤: {e}。請確保該欄位為有效的日期格式。")
                st.stop()

        except pd.errors.ParserError as pe:
            st.error(f"無法解析此 CSV 檔案：{pe}。請確認檔案格式是否正確（例如，是否為逗號分隔），並符合預期的欄位結構。")
            st.stop()
        except Exception as e: # Catch other potential errors during file reading
            st.error(f"讀取或初步處理檔案時發生未預期錯誤：{e}")
            st.stop()

        if df.empty:
            st.warning("CSV 檔案為空或不包含有效數據。請上傳有效的資料檔案。")
            st.stop()

        # --- DATA PREPROCESSING AND FEATURE ENGINEERING ---
        st.header("🛠️ 資料處理與特徵工程")
        with st.expander("查看資料處理詳情", expanded=False):
            st.write("正在進行資料清理、格式轉換及特徵提取...")

            # 1. Handle Missing Values for '新臺幣金額'
            if '新臺幣金額' in df.columns:
                df['新臺幣金額'] = df['新臺幣金額'].fillna(0.0)
            # else: already handled by initial column check

            # 2. Ensure Numeric Amount for '新臺幣金額'
            if '新臺幣金額' in df.columns:
                try:
                    df['新臺幣金額'] = df['新臺幣金額'].astype(str).str.replace(',', '', regex=False)
                    df['新臺幣金額'] = pd.to_numeric(df['新臺幣金額'], errors='coerce').fillna(0.0)
                except ValueError as ve: # Should be caught by pd.to_numeric with errors='coerce' mostly
                    st.error(f"轉換 '新臺幣金額' 為數值時出錯: {ve}。")
                    df['新臺幣金額'] = 0.0 # Fallback or stop
                except Exception as ex:
                    st.error(f"處理 '新臺幣金額' 時發生非預期錯誤: {ex}。")
                    df['新臺幣金額'] = 0.0 # Fallback or stop

            # Date feature engineering (assuming '消費日' is now datetime)
            if pd.api.types.is_datetime64_any_dtype(df['消費日']):
                df['YearMonth'] = df['消費日'].dt.to_period('M')
                df['Year'] = df['消費日'].dt.year
                df['Month'] = df['消費日'].dt.month
                st.write("已成功提取 年-月、年、月 資訊。")
            else:
                # This case should ideally be caught earlier, but as a safeguard:
                st.error("'消費日' 欄位未能成功轉換為日期格式，無法提取日期特徵。")
                # Create dummy columns if they are essential for later code not to break
                df['YearMonth'] = pd.NaT
                df['Year'] = pd.NA
                df['Month'] = pd.NA


            # 4. Basic Categorization
            df['Category'] = '雜項' # Initialize to Miscellaneous
            if '交易說明' in df.columns:
                df['交易說明'] = df['交易說明'].astype(str) # Ensure string type
                category_keywords = {
                    '餐飲': ['餐飲', '餐廳', '美食', 'FOOD', 'RESTAURANT', '小吃', '咖啡', '飲料', '茶餐廳', 'CAFE', 'STARBUCKS', '星巴克', '麥當勞', '肯德基', 'SUBWAY', '摩斯漢堡', '早午餐', '晚餐', '午餐', 'Uber Eats', 'Foodpanda'],
                    '交通': ['交通', '計程車', '油資', '停車', '高鐵', '台鐵', '捷運', 'UBER', 'TAXI', '客運', '航空', '車票', '加油', '移動', 'LINE TAXI', '台灣大車隊', '罰單'],
                    '購物': ['購物', '百貨', '超市', '網購', '服飾', 'PCHOME', 'MOMO', '蝦皮', 'SHOP', '線上購物', '便利商店', '藥妝', 'AMAZON', '淘寶', 'IHERB', '博客來', '誠品', 'COSTCO', 'PX MART'],
                    '娛樂': ['電影', 'KTV', '遊戲', 'NETFLIX', 'SPOTIFY', '娛樂', '演唱會', '展覽', 'DISNEY+', 'YOUTUBE PREMIUM', 'STEAM', 'Nintendo', '愛奇藝', 'KKBOX'],
                    '家居': ['IKEA', '家樂福', '全聯', '大潤發', '水電', '管理費', '電信費', '網路費', '手機費', '瓦斯費', '生活用品', '傢俱', '特力屋', '水費', '房租', '房貸'],
                    '健保/醫療': ['醫院', '診所', '藥局', '健保費', '醫療', '醫生', '掛號費', '保健', '手術', '牙醫'],
                    '進修學習': ['學費', '課程', '書籍', '講座', '線上課程', 'UDEMY', 'COURSERA', 'Hahow', 'TIPS', '補習班'],
                    '保險': ['保險', '保費', '壽險', '產險', '醫療險', '意外險', '國泰人壽', '富邦人壽', '南山人壽'],
                    '投資理財': ['基金', '股票', '股利', '利息收入', '手續費', '證券', 'ETF'],
                    '其他': ['手續費', '利息', '捐款', '雜支', '政府規費', '稅', '禮金', '紅包']
                }
                for category, keywords in category_keywords.items():
                    for keyword in keywords:
                        df.loc[df['交易說明'].str.contains(keyword, case=False, na=False, regex=False), 'Category'] = category
                st.write("已完成自動交易分類。")
            else:
                st.warning("缺少 '交易說明' 欄位，無法進行自動分類。所有交易將歸類為 '雜項'。")

            st.success("資料處理與特徵工程完成！")

            # Display DataFrame Info and Preview in Expander
            st.subheader("預覽處理後資料 (Processed Data Preview)")
            st.dataframe(df.head())

            buffer = BytesIO() # Changed from io.StringIO for df.info() to work with BytesIO as well
            df.info(buf=buffer)
            s = buffer.getvalue().decode('utf-8', errors='replace') # Decode bytes to string
            st.text_area("DataFrame Info (資料型態與缺失值)", s, height=250)

        # --- Sidebar: Summary Statistics & Categorization Info ---
        if not df.empty and '新臺幣金額' in df.columns:
            st.sidebar.markdown("---")
            st.sidebar.subheader("📊 財務總覽 (Summary)")
            total_spending_val = calculate_total_spending(df)
            num_transactions = len(df)
            avg_transaction = total_spending_val / num_transactions if num_transactions > 0 else 0.0

            st.sidebar.metric(label="總支出 (Total Spending)", value=f"NT$ {total_spending_val:,.2f}")
            st.sidebar.metric(label="總交易筆數 (Total Transactions)", value=f"{num_transactions:,}")
            st.sidebar.metric(label="平均交易金額 (Avg. Transaction)", value=f"NT$ {avg_transaction:,.2f}")

            if 'Category' in df.columns:
                misc_count = df[df['Category'] == '雜項'].shape[0]
                if num_transactions > 0:
                    misc_percentage = (misc_count / num_transactions) * 100
                    st.sidebar.info(f"未分類交易 (雜項): {misc_count} 筆 ({misc_percentage:.1f}%)")
                    if misc_percentage > 30:
                        st.sidebar.warning("大量交易未被自動分類。可考慮擴充關鍵字列表或檢查交易說明內容。")
        st.sidebar.markdown("---")


        # --- Main Dashboard Area ---
        st.header("📈 消費分析儀表板 (Spending Analysis Dashboard)")

        if not df.empty and '新臺幣金額' in df.columns:
            # Variables for charts, advice, and export - ensure they are defined before use
            spending_by_category = None
            spending_by_month = None
            fig_category = None
            fig_month = None
            outlier_transactions = None
            advice_list = []

            # Spending by Category Chart
            if 'Category' in df.columns:
                st.subheader("🍰 各類別消費分佈 (Spending by Category)")
                spending_by_category = calculate_spending_by_category(df) # This is a Series
                if not spending_by_category.empty:
                    spending_by_category_df = spending_by_category.reset_index()
                    spending_by_category_df.columns = ['Category', 'Amount']
                    fig_category = px.bar(spending_by_category_df, x='Category', y='Amount', title="各類別消費長條圖", labels={'Amount': '金額 (NT$)', 'Category': '消費類別'}, color='Category')
                    st.plotly_chart(fig_category, use_container_width=True)
                else:
                    st.info("尚無足夠資料顯示各類別消費分佈。")
            else:
                st.warning("缺少 'Category' 欄位，無法生成各類別消費圖表。")

            st.markdown("---")

            # Spending by Month Chart
            if 'YearMonth' in df.columns:
                st.subheader("📅 每月消費趨勢 (Spending by Month)")
                spending_by_month = calculate_spending_by_month(df) # This is a Series
                if not spending_by_month.empty:
                    spending_by_month_df = spending_by_month.reset_index()
                    spending_by_month_df.columns = ['YearMonth', 'Amount']
                    spending_by_month_df['YearMonth'] = spending_by_month_df['YearMonth'].astype(str)
                    fig_month = px.line(spending_by_month_df, x='YearMonth', y='Amount', title="每月消費趨勢折線圖", markers=True, labels={'Amount': '總金額 (NT$)', 'YearMonth': '月份'})
                    fig_month.update_xaxes(type='category')
                    st.plotly_chart(fig_month, use_container_width=True)
                else:
                    st.info("尚無足夠資料顯示每月消費趨勢。")
            else:
                st.warning("缺少 'YearMonth' 欄位，無法生成每月消費趨勢圖表。")

            st.markdown("---")

            # Outlier Transactions Table
            st.subheader("💰 高額交易 (> NT$ 5,000) (High-Value Transactions)")
            required_outlier_cols = ['消費日', '交易說明', '新臺幣金額', 'Category']
            if all(col in df.columns for col in required_outlier_cols):
                outlier_transactions = identify_outlier_transactions(df)
                if not outlier_transactions.empty:
                    st.dataframe(outlier_transactions)
                else:
                    st.info("未發現高於 NT$ 5,000 的交易。")
            else:
                st.warning(f"缺少必要欄位進行高額交易分析 (需要: {', '.join(required_outlier_cols)})。")

            st.markdown("---")

            # Spending Advice Section
            st.subheader("💡 消費建議 (Spending Advice)")
            if 'Category' in df.columns and 'YearMonth' in df.columns and '新臺幣金額' in df.columns and \
               spending_by_category is not None and spending_by_month is not None:
                advice_list = generate_spending_advice(df, spending_by_category, spending_by_month)
                if advice_list:
                    for advice_item in advice_list:
                        if advice_item == "目前尚無足夠的資料來產生具體的消費建議。":
                            st.info(advice_item)
                        else:
                            st.markdown(f"- {advice_item}")
            else:
                st.info("無法產生消費建議，因為缺少必要的資料欄位或分析結果。")

            st.markdown("---")
            st.header("📤 匯出報告 (Export Report)")

            # Excel Download
            if spending_by_category is not None and spending_by_month is not None:
                excel_data = to_excel(df.copy(), spending_by_category, spending_by_month)
                st.download_button(label="📥 下載 Excel 報告", data=excel_data, file_name="financial_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("主要分析數據未能成功計算，無法匯出 Excel 報告。")

            # PDF Download
            # Variables needed: df, total_spending_val, avg_transaction, num_transactions,
            # spending_by_category (Series), spending_by_month (Series),
            # outlier_transactions (DataFrame), advice_list, fig_category (Plotly fig), fig_month (Plotly fig)

            # Ensure all variables are defined before attempting PDF export
            pdf_export_possible = (
                df is not None and not df.empty and
                'total_spending_val' in locals() and
                'avg_transaction' in locals() and
                'num_transactions' in locals() and
                spending_by_category is not None and
                spending_by_month is not None and
                outlier_transactions is not None and # Ensure this is defined
                advice_list is not None and # Ensure this is defined
                fig_category is not None and # Ensure this is defined
                fig_month is not None # Ensure this is defined
            )

            if pdf_export_possible:
                sbc_df_for_pdf = spending_by_category.reset_index(); sbc_df_for_pdf.columns = ['Category', 'Amount']
                sbm_df_for_pdf = spending_by_month.reset_index(); sbm_df_for_pdf.columns = ['YearMonth', 'Amount']
                sbm_df_for_pdf['YearMonth'] = sbm_df_for_pdf['YearMonth'].astype(str)

                pdf_data = create_pdf_report(
                    df.copy(), total_spending_val, avg_transaction, num_transactions,
                    sbc_df_for_pdf, sbm_df_for_pdf,
                    outlier_transactions, advice_list,
                    fig_category, fig_month
                )
                st.download_button(label="📥 下載 PDF 報告", data=pdf_data, file_name="financial_report.pdf", mime="application/pdf")
            else:
                st.info("部分資料不完整或圖表未生成，無法產生 PDF 報告。")

        elif df.empty: # This case is after successful load but df became empty through processing
            st.warning("資料處理後無有效數據可供分析。")
        else: # Missing '新臺幣金額' or other critical columns for dashboard display
             st.error("核心欄位缺失或無法處理，無法顯示完整的財務分析儀表板。請檢查CSV檔案與資料處理步驟。")

    else: # No file uploaded
        st.info("👈 請從左方側邊欄上傳 CSV 檔案以開始分析。")
        st.markdown("""
        ### 使用說明
        1.  **準備您的 CSV 檔案**：確保檔案包含 `消費日`, `交易說明`, `新臺幣金額` 等欄位。
        2.  **上傳檔案**：點擊左方側邊欄的 "瀏覽檔案" 按鈕，選擇您的 CSV 檔案。
        3.  **查看分析**：儀表板將自動處理數據並顯示各項分析結果。
        4.  **下載報告**：您可以下載 Excel 或 PDF 格式的詳細報告。

        若遇到問題，請檢查 CSV 檔案格式或洽詢應用程式管理員。
        """)

if __name__ == "__main__":
    main()