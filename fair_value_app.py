import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import io
import base64

# Function to create Excel template with sample data
def create_excel_template():
    """
    Creates an Excel template with two sheets: Perusahaan_Target and Perusahaan_Sektor
    """
    # Sample data for Perusahaan_Target (single company)
    target_data = {
        'Ticker': ['BBRI'],
        'Nama_Perusahaan': ['Bank Rakyat Indonesia (Persero) Tbk'],
        'Sektor': ['Perbankan'],
        'Harga_Saham_Saat_Ini': [4520],
        'Jumlah_Saham_Beredar': [125_000_000_000], # Added underscore for readability
        'Net_Income_Terbaru': [56_000_000_000_000],
        'Total_Pendapatan_Terbaru': [150_000_000_000_000],
        'Total_Ekuitas_Terbaru': [280_000_000_000_000]
    }
    
    # Sample data for Perusahaan_Sektor (multiple companies in same sector)
    sector_data = {
        'Ticker': ['BBCA', 'BMRI', 'BBNI'],
        'Nama_Perusahaan': [
            'Bank Central Asia Tbk',
            'Bank Mandiri (Persero) Tbk', 
            'Bank Negara Indonesia (Persero) Tbk'
        ],
        'Sektor': ['Perbankan', 'Perbankan', 'Perbankan'],
        'Harga_Saham_Saat_Ini': [8750, 6225, 4890],
        'Jumlah_Saham_Beredar': [25_000_000_000, 23_000_000_000, 19_000_000_000],
        'Net_Income_Terbaru': [32_000_000_000_000, 34_000_000_000_000, 17_000_000_000_000],
        'Total_Pendapatan_Terbaru': [80_000_000_000_000, 95_000_000_000_000, 65_000_000_000_000],
        'Total_Ekuitas_Terbaru': [180_000_000_000_000, 200_000_000_000_000, 140_000_000_000_000]
    }
    
    # Create DataFrames
    df_target = pd.DataFrame(target_data)
    df_sector = pd.DataFrame(sector_data)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write target company sheet
        df_target.to_excel(writer, sheet_name='Perusahaan_Target', index=False)
        
        # Write sector companies sheet
        df_sector.to_excel(writer, sheet_name='Perusahaan_Sektor', index=False)
        
        # Get workbook and worksheets
        workbook = writer.book
        worksheet_target = writer.sheets['Perusahaan_Target']
        worksheet_sector = writer.sheets['Perusahaan_Sektor']
        
        # Format headers
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Format numbers
        number_format_currency = workbook.add_format({
            'num_format': '#,##0',
            'border': 1
        })
        
        # Apply formatting to target sheet
        for col_num, value in enumerate(df_target.columns.values):
            worksheet_target.write(0, col_num, value, header_format)
            worksheet_target.set_column(col_num, col_num, 20)
        
        # Apply formatting to sector sheet
        for col_num, value in enumerate(df_sector.columns.values):
            worksheet_sector.write(0, col_num, value, header_format)
            worksheet_sector.set_column(col_num, col_num, 20)
        
        # Format number columns for target sheet
        for row in range(1, len(df_target) + 1):
            # Apply currency format to relevant columns
            for col_idx in [3, 4, 5, 6, 7]: # Harga_Saham_Saat_Ini, Jumlah_Saham_Beredar, Net_Income_Terbaru, Total_Pendapatan_Terbaru, Total_Ekuitas_Terbaru
                worksheet_target.write(row, col_idx, df_target.iloc[row-1, col_idx], number_format_currency)
        
        # Format number columns for sector sheet
        for row in range(1, len(df_sector) + 1):
            # Apply currency format to relevant columns
            for col_idx in [3, 4, 5, 6, 7]: # Harga_Saham_Saat_Ini, Jumlah_Saham_Beredar, Net_Income_Terbaru, Total_Pendapatan_Terbaru, Total_Ekuitas_Terbaru
                worksheet_sector.write(row, col_idx, df_sector.iloc[row-1, col_idx], number_format_currency)
    
    processed_data = output.getvalue()
    return processed_data

# Function to create empty Excel template
def create_empty_excel_template():
    """
    Creates an empty Excel template with headers only
    """
    # Headers only
    headers = ['Ticker', 'Nama_Perusahaan', 'Sektor', 'Harga_Saham_Saat_Ini', 
               'Jumlah_Saham_Beredar', 'Net_Income_Terbaru', 'Total_Pendapatan_Terbaru', 
               'Total_Ekuitas_Terbaru']
    
    # Create empty DataFrames with headers
    df_target = pd.DataFrame(columns=headers)
    df_sector = pd.DataFrame(columns=headers)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write target company sheet
        df_target.to_excel(writer, sheet_name='Perusahaan_Target', index=False)
        
        # Write sector companies sheet
        df_sector.to_excel(writer, sheet_name='Perusahaan_Sektor', index=False)
        
        # Get workbook and worksheets
        workbook = writer.book
        worksheet_target = writer.sheets['Perusahaan_Target']
        worksheet_sector = writer.sheets['Perusahaan_Sektor']
        
        # Format headers
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Apply formatting to both sheets
        for col_num, value in enumerate(headers):
            worksheet_target.write(0, col_num, value, header_format)
            worksheet_target.set_column(col_num, col_num, 20)
            
            worksheet_sector.write(0, col_num, value, header_format)
            worksheet_sector.set_column(col_num, col_num, 20)
    
    processed_data = output.getvalue()
    return processed_data

# Function to read data from the uploaded Excel file
@st.cache_data
def read_excel_data(uploaded_file):
    """
    Reads target company and sector comparable company data from an Excel file.
    Assumptions:
    - The first sheet is named 'Perusahaan_Target' and contains single-row data.
    - The second sheet is named 'Perusahaan_Sektor' and contains multi-row data.
    """
    try:
        xls = pd.ExcelFile(uploaded_file)

        # Read 'Perusahaan_Target' sheet
        if 'Perusahaan_Target' in xls.sheet_names:
            df_target = pd.read_excel(xls, sheet_name='Perusahaan_Target')
            if not df_target.empty:
                # Take the first row (assuming only one target company)
                target_data = df_target.iloc[0].to_dict()
            else:
                st.error("Sheet 'Perusahaan_Target' is empty or has no data.")
                return None, None
        else:
            st.error("Excel file must have a sheet named 'Perusahaan_Target'.")
            return None, None

        # Read 'Perusahaan_Sektor' sheet
        if 'Perusahaan_Sektor' in xls.sheet_names:
            df_sector = pd.read_excel(xls, sheet_name='Perusahaan_Sektor')
            if df_sector.empty:
                st.warning("Sheet 'Perusahaan_Sektor' is empty. Sector comparison will not be performed.")
        else:
            st.warning("Excel file does not have a sheet named 'Perusahaan_Sektor'. Sector comparison will not be performed.")
            df_sector = pd.DataFrame() # Create an empty DataFrame

        return target_data, df_sector
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}. Please ensure the file format and sheet names are correct.")
        return None, None

# Function to calculate key financial ratios
def calculate_key_ratios(net_income, revenue, total_equity, current_price, shares_outstanding):
    """
    Calculates key financial ratios (Net Profit Margin, ROE, P/E Ratio, P/B Ratio)
    from direct inputs.
    """
    ratios = {}

    # Net Profit Margin
    if net_income is not None and revenue is not None and revenue != 0:
        ratios['Net Profit Margin (%)'] = (net_income / revenue) * 100
    else:
        ratios['Net Profit Margin (%)'] = float('nan')

    # ROE (Return on Equity)
    if net_income is not None and total_equity is not None and total_equity != 0:
        ratios['ROE (%)'] = (net_income / total_equity) * 100
    else:
        ratios['ROE (%)'] = float('nan')

    # P/E Ratio (Price-to-Earnings Ratio)
    if current_price is not None and shares_outstanding is not None and shares_outstanding != 0:
        if net_income is not None and net_income != 0:
            eps = net_income / shares_outstanding
            ratios['P/E Ratio'] = current_price / eps if eps != 0 else float('inf')
        else:
            ratios['P/E Ratio'] = float('inf')

    # P/B Ratio (Price-to-Book Ratio)
    if current_price is not None and shares_outstanding is not None and shares_outstanding != 0:
        if total_equity is not None and total_equity != 0:
            book_value_per_share = total_equity / shares_outstanding if shares_outstanding != 0 else 0
            ratios['P/B Ratio'] = current_price / book_value_per_share if book_value_per_share != 0 else float('inf')
        else:
            ratios['P/B Ratio'] = float('inf')

    return {k: v for k, v in ratios.items() if not pd.isna(v)}

# Function to filter comparable companies by sector from Excel data
def get_sector_comparables_from_excel(comparables_df, target_sector):
    """
    Filters the comparable companies DataFrame to get only those in the same sector.
    """
    if comparables_df.empty or 'Sektor' not in comparables_df.columns:
        return pd.DataFrame() # Return empty if no data or 'Sektor' column

    # Ensure case-insensitive sector comparison
    filtered_comparables = comparables_df[
        comparables_df['Sektor'].astype(str).str.lower() == target_sector.lower()
    ]
    return filtered_comparables

# Function to calculate average sector ratios from Excel data
def calculate_sector_averages_from_excel(comparable_companies_data):
    """
    Calculates average financial ratios for a specific sector
    from the filtered comparable companies DataFrame.
    """
    all_pe_ratios = []
    all_pb_ratios = []
    all_roe_ratios = []
    all_npm_ratios = []

    if comparable_companies_data.empty:
        return {}

    for index, row in comparable_companies_data.iterrows():
        current_price = row.get('Harga_Saham_Saat_Ini')
        shares_outstanding = row.get('Jumlah_Saham_Beredar')
        net_income = row.get('Net_Income_Terbaru')
        revenue = row.get('Total_Pendapatan_Terbaru')
        total_equity = row.get('Total_Ekuitas_Terbaru')

        if (current_price is not None and shares_outstanding is not None and shares_outstanding != 0 and
            net_income is not None and revenue is not None and total_equity is not None):

            ratios = calculate_key_ratios(net_income, revenue, total_equity, current_price, shares_outstanding)

            if 'P/E Ratio' in ratios and ratios['P/E Ratio'] != float('inf') and not pd.isna(ratios['P/E Ratio']):
                all_pe_ratios.append(ratios['P/E Ratio'])
            if 'P/B Ratio' in ratios and ratios['P/B Ratio'] != float('inf') and not pd.isna(ratios['P/B Ratio']):
                all_pb_ratios.append(ratios['P/B Ratio'])
            if 'ROE (%)' in ratios and not pd.isna(ratios['ROE (%)']):
                all_roe_ratios.append(ratios['ROE (%)'])
            if 'Net Profit Margin (%)' in ratios and not pd.isna(ratios['Net Profit Margin (%)']):
                all_npm_ratios.append(ratios['Net Profit Margin (%)'])
        else:
            # Suppress warnings for individual missing comparables data
            pass

    sector_averages = {}
    if all_pe_ratios:
        sector_averages['P/E Ratio'] = sum(all_pe_ratios) / len(all_pe_ratios)
    if all_pb_ratios:
        sector_averages['P/B Ratio'] = sum(all_pb_ratios) / len(all_pb_ratios)
    if all_roe_ratios:
        sector_averages['ROE (%)'] = sum(all_roe_ratios) / len(all_roe_ratios)
    if all_npm_ratios:
        sector_averages['Net Profit Margin (%)'] = sum(all_npm_ratios) / len(all_npm_ratios)

    return sector_averages

# Function to calculate fair value using the multiplier method
def calculate_fair_value_multiplier(ticker_ratios, sector_avg_ratios, current_price, method='P/E'):
    """
    Calculates fair value using the multiplier method (P/E or P/B).
    Assumption: The fair price of the company is when its ratio equals the sector average.
    """
    fair_value = None
    if current_price is None:
        return None

    # Fair value calculation based on P/E Ratio
    if method == 'P/E' and 'P/E Ratio' in ticker_ratios and 'P/E Ratio' in sector_avg_ratios:
        if ticker_ratios['P/E Ratio'] != 0 and ticker_ratios['P/E Ratio'] != float('inf') and not pd.isna(ticker_ratios['P/E Ratio']):
            # Ensure sector_avg_ratios['P/E Ratio'] is not zero to avoid division by zero
            if sector_avg_ratios['P/E Ratio'] != 0:
                fair_value_multiplier = sector_avg_ratios['P/E Ratio'] / ticker_ratios['P/E Ratio']
                fair_value = current_price * fair_value_multiplier
            else:
                fair_value = float('inf') # Sector average P/E is zero, implies infinite fair value if current P/E is not zero
    # Fair value calculation based on P/B Ratio
    elif method == 'P/B' and 'P/B Ratio' in ticker_ratios and 'P/B Ratio' in sector_avg_ratios:
        if ticker_ratios['P/B Ratio'] != 0 and ticker_ratios['P/B Ratio'] != float('inf') and not pd.isna(ticker_ratios['P/B Ratio']):
            # Ensure sector_avg_ratios['P/B Ratio'] is not zero to avoid division by zero
            if sector_avg_ratios['P/B Ratio'] != 0:
                fair_value_multiplier = sector_avg_ratios['P/B Ratio'] / ticker_ratios['P/B Ratio']
                fair_value = current_price * fair_value_multiplier
            else:
                fair_value = float('inf') # Sector average P/B is zero, implies infinite fair value if current P/B is not zero
    return fair_value

def create_gauge_chart(current_price, fair_value, title):
    """Create a gauge chart for fair value visualization"""
    if fair_value is None or pd.isna(fair_value) or fair_value <= 0:
        return None
    
    # Define min and max for the gauge to ensure visibility and meaningful range
    # Ensure current_price is also positive for meaningful gauge
    if current_price <= 0:
        return None

    # Determine a reasonable max value for the gauge axis
    # It should be greater than both current_price and fair_value
    max_value = max(current_price, fair_value) * 1.2
    min_value = min(current_price, fair_value) * 0.8
    if min_value < 0: min_value = 0 # Ensure minimum is not negative

    # Avoid zero range if fair_value is extremely small
    if max_value - min_value < 1: 
        max_value = max(current_price, fair_value) + 10 # Small buffer
        min_value = min(current_price, fair_value) - 10 
        if min_value < 0: min_value = 0

    fig = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = current_price,
        delta = {'reference': fair_value, 'increasing': {'color': "#FF6B6B"}, 'decreasing': {'color': "#4ECDC4"}},
        title = {'text': title, 'font': {'size': 12, 'color': '#2E3440'}},
        gauge = {
            'axis': {'range': [min_value, max_value], 'tickcolor': '#5E81AC'},
            'bar': {'color': "#5E81AC"},
            'steps': [
                {'range': [min_value, fair_value * 0.8], 'color': "#A3BE8C"}, # Undervalued (Greenish)
                {'range': [fair_value * 0.8, fair_value * 1.2], 'color': "#EBCB8B"}, # Fair (Yellowish)
                {'range': [fair_value * 1.2, max_value], 'color': "#BF616A"} # Overvalued (Reddish)
            ],
            'threshold': {
                'line': {'color': "#D08770", 'width': 4},
                'thickness': 0.75,
                'value': fair_value
            }
        }
    ))
    
    fig.update_layout(
        height=300,
        margin=dict(l=20, r=20, t=60, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': '#2E3440'}
    )
    
    return fig

def create_comparison_chart(target_ratios, sector_ratios):
    """Create a comparison chart between target and sector ratios"""
    if not target_ratios or not sector_ratios:
        return None
    
    # Find common ratios
    common_ratios = set(target_ratios.keys()) & set(sector_ratios.keys())
    if not common_ratios:
        return None
    
    ratios = list(common_ratios)
    target_values = [target_ratios[ratio] for ratio in ratios]
    sector_values = [sector_ratios[ratio] for ratio in ratios]
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name='Perusahaan Target',
        x=ratios,
        y=target_values,
        marker_color='#5E81AC',
        text=[f'{v:.2f}' for v in target_values],
        textposition='auto',
    ))
    
    fig.add_trace(go.Bar(
        name='Rata-rata Sektor',
        x=ratios,
        y=sector_values,
        marker_color='#88C999',
        text=[f'{v:.2f}' for v in sector_values],
        textposition='auto',
    ))
    
    fig.update_layout(
        title='Perbandingan Rasio Keuangan',
        xaxis_title='Rasio',
        yaxis_title='Nilai',
        barmode='group',
        height=400,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': '#2E3440'},
        title_font={'size': 15, 'color': '#2E3440'}
    )
    
    return fig

# --- Streamlit User Interface ---
st.set_page_config(
    layout="wide", 
    page_title="Kalkulator Fair Value Saham",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern design
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 0;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 3rem 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    }
    
    .main-header h1 {
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 1rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    
    .main-header p {
        font-size: 1rem;
        font-weight: 300;
        opacity: 0.9;
    }
    
    .info-card {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 5px 20px rgba(0,0,0,0.08);
        margin-bottom: 1.5rem;
        border-left: 4px solid #667eea;
    }
    
    /* New styles for data display */
    .data-container {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        color: #ffffff;
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
        max-width: 500px; /* Limit width for better readability */
        width: 100%;
        display: flex;
        flex-direction: column;
        gap: 1.5rem; /* Space between each data item */
        margin: 0 auto 1.5rem auto; /* Center the container and add bottom margin */
    }
    .data-item {
        text-align: center;
        padding: 1rem;
        border-bottom: 1px solid rgba(255, 255, 255, 0.3); /* Subtle separator */
    }
    .data-item:last-child {
        border-bottom: none; /* No border for the last item */
    }
    .data-value {
        font-size: 1.6rem; /* Larger font for the value */
        font-weight: 700;
        margin-bottom: 0.25rem; /* Small space between value and label */
        line-height: 1.2; /* Adjust line height for better spacing */
    }
    .data-label {
        font-size: 0.9rem; /* Smaller font for the label */
        font-weight: 400;
        opacity: 0.9; /* Slightly transparent for distinction */
    }

    /* Responsive adjustments for smaller screens */
    @media (max-width: 640px) {
        .data-container {
            padding: 1.5rem;
            gap: 1rem;
        }
        .data-value {
            font-size: 1.5rem;
        }
        .data-label {
            font-size: 0.8rem;
        }
    }
    /* End of new styles for data display */

    .status-card {
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        font-weight: 500;
        text-align: center;
        font-size: 1.1rem;
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
    }
    
    .status-undervalued {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        color: white;
    }
    
    .status-overvalued {
        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
        color: white;
    }
    
    .status-fair {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
        color: #2c3e50;
    }
    
    .sidebar .sidebar-content {
        background: white;
        border-radius: 15px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .stFileUploader > div > div {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
        padding: 2rem;
        border: none;
        color: white;
        text-align: center;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.3);
    }
    
    .stDataFrame {
        background: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
    }
    
    .footer {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 15px;
        margin-top: 3rem;
        text-align: center;
    }
    
    .divider {
        height: 3px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 2px;
        margin: 2rem 0;
    }
    
    .download-section {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        text-align: center;
    }
    
    .download-section h3 {
        color: white;
        margin-bottom: 1rem;
    }
    
    .download-section p {
        color: rgba(255,255,255,0.8);
        font-size: 0.9rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📊 Kalkulator Fair Value Saham</h1>
    <p>Analisis kelayakan investasi saham berdasarkan data fundamental perusahaan</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 1rem;">
        <h2 style="color: #667eea;">📁 Upload Data</h2>
        <p>Unggah file Excel dengan data perusahaan target dan pembanding</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Excel Template Download Section
    st.markdown("""
    <div class="download-section">
        <h3>📥 Download Template Excel</h3>
        <p>Download template Excel dengan format yang sesuai dan contoh data</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Create download buttons
    col1, col2 = st.columns(2)
    
    with col1:
        # Template with sample data
        template_data = create_excel_template()
        st.download_button(
            label="📊 Template + Contoh",
            data=template_data,
            file_name="template_fair_value_dengan_contoh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Template Excel dengan contoh data perusahaan perbankan"
        )
    
    with col2:
        # Empty template
        empty_template_data = create_empty_excel_template()
        st.download_button(
            label="📄 Template Kosong",
            data=empty_template_data,
            file_name="template_fair_value_kosong.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Template Excel kosong untuk diisi dengan data Anda"
        )
    
    st.markdown("---")
    
    uploaded_file = st.file_uploader(
        "Pilih file Excel (.xlsx)",
        type=["xlsx"],
        help="File harus berisi sheet 'Perusahaan_Target' dan 'Perusahaan_Sektor'"
    )
    
    if uploaded_file is not None:
        st.success("✅ File berhasil diunggah!")
        
        # Add some info about the expected format
        with st.expander("ℹ️ Format File Excel"):
            st.markdown("""
            **Sheet 1 'Perusahaan_Target':** Data satu perusahaan target
            
            **Sheet 2 'Perusahaan_Sektor':** Data 3+ perusahaan pembanding 1 sektor
            
            **Kolom yang diperlukan di kedua sheet:**
            - **Ticker**: Kode saham (contoh: BBRI, BBCA)
            - **Nama_Perusahaan**: Nama lengkap perusahaan
            - **Sektor**: Sektor industri (contoh: Perbankan, Teknologi)
            - **Harga_Saham_Saat_Ini**: Harga saham terkini (dalam Rupiah)
            - **Jumlah_Saham_Beredar**: Jumlah saham yang beredar
            - **Net_Income_Terbaru**: Laba bersih terbaru (dalam Rupiah)
            - **Total_Pendapatan_Terbaru**: Total pendapatan terbaru (dalam Rupiah)
            - **Total_Ekuitas_Terbaru**: Total ekuitas terbaru (dalam Rupiah)
            """)

if uploaded_file is not None:
    with st.spinner("🔄 Memproses data..."):
        target_data, df_sector_comparables = read_excel_data(uploaded_file)

    if target_data is not None:
        # Extract target company data
        ticker_target = target_data.get('Ticker')
        company_name_target = target_data.get('Nama_Perusahaan')
        sector_target = target_data.get('Sektor')
        shares_outstanding_target = target_data.get('Jumlah_Saham_Beredar')
        current_price_target = target_data.get('Harga_Saham_Saat_Ini')
        net_income_target = target_data.get('Net_Income_Terbaru')
        revenue_target = target_data.get('Total_Pendapatan_Terbaru')
        total_equity_target = target_data.get('Total_Ekuitas_Terbaru')

        # Ensure all key data is present
        if all(x is not None for x in [ticker_target, company_name_target, sector_target,
                                       shares_outstanding_target, current_price_target,
                                       net_income_target, revenue_target, total_equity_target]):
            
            # Company info section
            st.markdown(f"""
            <div class="info-card">
                <h2 style="color: #667eea; margin-bottom: 1rem;">🏢 {company_name_target} ({ticker_target})</h2>
                <p style="color: #666; font-size: 1.1rem; margin-bottom: 0;"><strong>Sektor:</strong> {sector_target}</p>
            </div>
            """, unsafe_allow_html=True)

            # Key metrics using the new data-container and data-item structure
            st.markdown(f"""
            <div class="data-container">
                <!-- Harga Saham Saat Ini -->
                <div class="data-item">
                    <div class="data-value">Rp {current_price_target:,.0f}</div>
                    <div class="data-label">Harga Saham Saat Ini</div>
                <div class="data-item">
                    <div class="data-value">{shares_outstanding_target:,.0f}</div>
                    <div class="data-label">Jumlah Saham Beredar</div>
                <div class="data-item">
                    <div class="data-value">Rp {net_income_target:,.0f}</div>
                    <div class="data-label">Net Income Terbaru</div>
                <div class="data-item">
                    <div class="data-value">Rp {revenue_target:,.0f}</div>
                    <div class="data-label">Total Pendapatan</div>
                <div class="data-item">
                    <div class="data-value">Rp {total_equity_target:,.0f}</div>
                    <div class="data-label">Total Ekuitas</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # Calculate target company ratios
            ratios_target = calculate_key_ratios(
                net_income_target, revenue_target, total_equity_target,
                current_price_target, shares_outstanding_target
            )

            # Financial ratios section
            st.markdown("""
            <div class="info-card">
                <h2 style="color: #667eea; margin-bottom: 1rem;">📊 Rasio Keuangan</h2>
            </div>
            """, unsafe_allow_html=True)

            if ratios_target:
                # Display ratios in cards
                ratio_cols = st.columns(len(ratios_target))
                for i, (ratio, value) in enumerate(ratios_target.items()):
                    with ratio_cols[i]:
                        if 'Ratio' in ratio:
                            display_value = f"{value:.2f}x"
                        else:
                            display_value = f"{value:.2f}%"
                        
                        # Define help text for each ratio
                        help_text = ""
                        if ratio == 'Net Profit Margin (%)':
                            help_text = "Persentase laba bersih dari setiap rupiah pendapatan. Semakin tinggi, semakin efisien perusahaan mengelola biaya dan mengubah penjualan menjadi keuntungan."
                        elif ratio == 'ROE (%)':
                            help_text = "Return on Equity: Mengukur seberapa efisien perusahaan menggunakan modal (ekuitas) pemegang saham untuk menghasilkan laba bersih. Semakin tinggi, semakin baik profitabilitas bagi pemilik."
                        elif ratio == 'P/E Ratio':
                            help_text = "Price-to-Earnings Ratio: Menunjukkan berapa kali investor bersedia membayar untuk setiap Rp1 laba bersih yang dihasilkan perusahaan. P/E tinggi bisa berarti ekspektasi pertumbuhan tinggi atau saham dinilai premium. P/E rendah bisa berarti undervalue atau prospek pertumbuhan rendah."
                        elif ratio == 'P/B Ratio':
                            help_text = "Price-to-Book Ratio: Membandingkan harga saham dengan nilai aset bersih (ekuitas) perusahaan. Ini menunjukkan berapa kali investor bersedia membayar untuk setiap Rp1 nilai buku perusahaan. P/B > 1x umum jika perusahaan diharapkan menghasilkan laba di atas nilai asetnya, atau memiliki aset tak berwujud seperti merek."

                        st.markdown(f"""
                        <div style="background: white; padding: 1rem; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem;">
                            <h4 style="color: #667eea; margin: 0;">{display_value}</h4>
                            <p style="color: #666; margin: 0; font-size: 0.9rem;">{ratio} <span title="{help_text}" style="cursor: help; color: #5E81AC;">ⓘ</span></p>
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.info("ℹ️ Tidak dapat menghitung rasio keuangan. Pastikan data 'Net_Income_Terbaru', 'Total_Pendapatan_Terbaru', 'Total_Ekuitas_Terbaru', 'Harga_Saham_Saat_Ini', dan 'Jumlah_Saham_Beredar' terisi dengan angka positif.")

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # Sector comparison section
            st.markdown(f"""
            <div class="info-card">
                <h2 style="color: #667eea; margin-bottom: 1rem;">🔄 Analisis Sektor: {sector_target}</h2>
            </div>
            """, unsafe_allow_html=True)

            if not df_sector_comparables.empty:
                # Filter comparable companies from Excel by sector
                comparable_companies_in_sector = get_sector_comparables_from_excel(df_sector_comparables, sector_target)

                if not comparable_companies_in_sector.empty:
                    st.markdown(f"""
                    <div style="background: white; padding: 1.5rem; border-radius: 10px; margin-bottom: 1rem;">
                        <h3 style="color: #667eea;">📋 Perusahaan Pembanding ({len(comparable_companies_in_sector)} perusahaan)</h3>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Display comparable companies
                    display_cols = ['Ticker', 'Nama_Perusahaan', 'Harga_Saham_Saat_Ini', 'Net_Income_Terbaru', 'Total_Ekuitas_Terbaru'] # Added Total_Ekuitas_Terbaru
                    st.dataframe(
                        comparable_companies_in_sector[display_cols].style.format({
                            'Harga_Saham_Saat_Ini': "Rp {:,.0f}",
                            'Net_Income_Terbaru': "Rp {:,.0f}",
                            'Total_Ekuitas_Terbaru': "Rp {:,.0f}" # Formatted Total_Ekuitas_Terbaru
                        }),
                        use_container_width=True
                    )

                    # Calculate average sector ratios
                    sector_avg_ratios = calculate_sector_averages_from_excel(comparable_companies_in_sector)

                    if sector_avg_ratios:
                        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                        
                        # Comparison chart
                        comparison_chart = create_comparison_chart(ratios_target, sector_avg_ratios)
                        if comparison_chart:
                            st.plotly_chart(comparison_chart, use_container_width=True)

                        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

                        # Fair Value Analysis
                        st.markdown("""
                        <div class="info-card">
                            <h2 style="color: #667eea; margin-bottom: 1rem;">💎 Analisis Fair Value</h2>
                        </div>
                        """, unsafe_allow_html=True)

                        # Fair value calculations
                        fair_value_pe = calculate_fair_value_multiplier(ratios_target, sector_avg_ratios, current_price_target, method='P/E')
                        fair_value_pb = calculate_fair_value_multiplier(ratios_target, sector_avg_ratios, current_price_target, method='P/B')

                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if fair_value_pe is not None and not pd.isna(fair_value_pe) and fair_value_pe > 0:
                                # P/E Fair Value
                                st.markdown(f"""
                                <div style="background: white; padding: 1.5rem; border-radius: 15px; text-align: center; box-shadow: 0 5px 15px rgba(0,0,0,0.1);">
                                    <h3 style="color: #667eea;">Fair Value (P/E)</h3>
                                    <h2 style="color: #2c3e50; font-size: 2rem;">Rp {fair_value_pe:,.0f}</h2>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                # Status for P/E
                                if current_price_target < fair_value_pe * 0.95: # Adding a small buffer for "undervalued"
                                    st.markdown("""
                                    <div class="status-card status-undervalued">
                                        🚀 UNDERVALUED - Berdasarkan P/E<br>
                                        <small>Harga di bawah fair value, potensi profit tinggi</small>
                                    </div>
                                    """, unsafe_allow_html=True)
                                elif current_price_target > fair_value_pe * 1.05: # Adding a small buffer for "overvalued"
                                    st.markdown("""
                                    <div class="status-card status-overvalued">
                                        ⚠️ OVERVALUED - Berdasarkan P/E<br>
                                        <small>Harga di atas fair value, pertimbangkan risiko</small>
                                    </div>
                                    """, unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="status-card status-fair">
                                        ✅ FAIR VALUE - Berdasarkan P/E<br>
                                        <small>Harga sesuai dengan fair value</small>
                                    </div>
                                    """, unsafe_allow_html=True)
                                
                                # Gauge chart for P/E
                                pe_gauge = create_gauge_chart(current_price_target, fair_value_pe, "P/E Fair Value")
                                if pe_gauge:
                                    st.plotly_chart(pe_gauge, use_container_width=True)
                            else:
                                st.info("ℹ️ Fair Value (P/E) tidak dapat dihitung atau tidak valid. Pastikan Net Income Target tidak nol dan rata-rata P/E sektor tidak nol.")
                        
                        with col2:
                            if fair_value_pb is not None and not pd.isna(fair_value_pb) and fair_value_pb > 0:
                                # P/B Fair Value
                                st.markdown(f"""
                                <div style="background: white; padding: 1.5rem; border-radius: 15px; text-align: center; box-shadow: 0 5px 15px rgba(0,0,0,0.1);">
                                    <h3 style="color: #667eea;">Fair Value (P/B)</h3>
                                    <h2 style="color: #2c3e50; font-size: 2rem;">Rp {fair_value_pb:,.0f}</h2>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                # Status for P/B
                                if current_price_target < fair_value_pb * 0.95: # Adding a small buffer
                                    st.markdown("""
                                    <div class="status-card status-undervalued">
                                        🚀 UNDERVALUED - Berdasarkan P/B<br>
                                        <small>Harga di bawah fair value, potensi profit tinggi</small>
                                    </div>
                                    """, unsafe_allow_html=True)
                                elif current_price_target > fair_value_pb * 1.05: # Adding a small buffer
                                    st.markdown("""
                                    <div class="status-card status-overvalued">
                                        ⚠️ OVERVALUED - Berdasarkan P/B<br>
                                        <small>Harga di atas fair value, pertimbangkan risiko</small>
                                    </div>
                                    """, unsafe_allow_html=True)
                                else:
                                    st.markdown("""
                                    <div class="status-card status-fair">
                                        ✅ FAIR VALUE - Berdasarkan P/B<br>
                                        <small>Harga sesuai dengan fair value</small>
                                    </div>
                                    """, unsafe_allow_html=True)
                                
                                # Gauge chart for P/B
                                pb_gauge = create_gauge_chart(current_price_target, fair_value_pb, "P/B Fair Value")
                                if pb_gauge:
                                    st.plotly_chart(pb_gauge, use_container_width=True)
                            else:
                                st.info("ℹ️ Fair Value (P/B) tidak dapat dihitung atau tidak valid. Pastikan Total Ekuitas Target tidak nol dan rata-rata P/B sektor tidak nol.")

                        # Summary analysis
                        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                        
                        # Investment recommendation
                        st.markdown("""
                        <div class="info-card">
                            <h2 style="color: #667eea; margin-bottom: 1rem;">📈 Rekomendasi Investasi</h2>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Calculate overall recommendation
                        undervalued_count = 0
                        overvalued_count = 0
                        
                        if fair_value_pe is not None and not pd.isna(fair_value_pe) and fair_value_pe > 0:
                            if current_price_target < fair_value_pe * 0.95:
                                undervalued_count += 1
                            elif current_price_target > fair_value_pe * 1.05:
                                overvalued_count += 1
                        
                        if fair_value_pb is not None and not pd.isna(fair_value_pb) and fair_value_pb > 0:
                            if current_price_target < fair_value_pb * 0.95:
                                undervalued_count += 1
                            elif current_price_target > fair_value_pb * 1.05:
                                overvalued_count += 1
                        
                        # Display recommendation
                        if undervalued_count > overvalued_count:
                            st.markdown("""
                            <div class="status-card status-undervalued" style="font-size: 1.2rem; padding: 2rem;">
                                <h3 style="margin-bottom: 1rem;">🎯 REKOMENDASI: BUY</h3>
                                <p>Mayoritas indikator menunjukkan saham ini **undervalued**. Ini bisa menjadi peluang investasi yang menarik.</p>
                                <small><strong>Catatan:</strong> Selalu lakukan riset mendalam sebelum berinvestasi</small>
                            </div>
                            """, unsafe_allow_html=True)
                        elif overvalued_count > undervalued_count:
                            st.markdown("""
                            <div class="status-card status-overvalued" style="font-size: 1.2rem; padding: 2rem;">
                                <h3 style="margin-bottom: 1rem;">⚠️ REKOMENDASI: HOLD/SELL</h3>
                                <p>Mayoritas indikator menunjukkan saham ini **overvalued**. Pertimbangkan untuk menunggu atau menjual.</p>
                                <small><strong>Catatan:</strong> Evaluasi faktor fundamental lainnya</small>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown("""
                            <div class="status-card status-fair" style="font-size: 1.2rem; padding: 2rem;">
                                <h3 style="margin-bottom: 1rem;">⚖️ REKOMENDASI: HOLD</h3>
                                <p>Indikator menunjukkan hasil yang **seimbang**. Pertimbangkan faktor lain sebelum mengambil keputusan.</p>
                                <small><strong>Catatan:</strong> Analisis lebih dalam diperlukan</small>
                            </div>
                            """, unsafe_allow_html=True)

                        # Risk factors
                        st.markdown("""
                        <div class="info-card">
                            <h3 style="color: #667eea;">⚠️ Faktor Risiko yang Perlu Dipertimbangkan</h3>
                            <ul style="color: #666; line-height: 1.6;">
                                <li>Kondisi ekonomi makro dan industri</li>
                                <li>Kinerja manajemen dan strategi perusahaan</li>
                                <li>Kompetisi dan posisi pasar</li>
                                <li>Regulasi pemerintah yang dapat mempengaruhi bisnis</li>
                                <li>Volatilitas pasar saham secara keseluruhan</li>
                            </ul>
                        </div>
                        """, unsafe_allow_html=True)

                    else:
                        st.warning("❌ Tidak dapat menghitung rata-rata rasio sektor. Pastikan data perusahaan pembanding lengkap dan valid.")
                else:
                    st.warning(f"❌ Tidak ditemukan perusahaan pembanding di sektor **'{sector_target}'** dalam file Excel Anda. Pastikan nama sektor di sheet 'Perusahaan_Sektor' sesuai dengan 'Perusahaan_Target'.")
            else:
                st.info("ℹ️ Tidak ada data perusahaan pembanding di sheet 'Perusahaan_Sektor'. Analisis perbandingan sektor tidak dapat dilakukan.")

        else:
            st.error("❌ Data perusahaan target tidak lengkap. Pastikan semua kolom wajib (Ticker, Nama_Perusahaan, Sektor, Harga_Saham_Saat_Ini, Jumlah_Saham_Beredar, Net_Income_Terbaru, Total_Pendapatan_Terbaru, Total_Ekuitas_Terbaru) terisi di sheet 'Perusahaan_Target'.")

# Footer
st.markdown("""
<div class="footer">
    <h3>📚 Disclaimer</h3>
    <p>Kalkulator ini adalah alat bantu analisis yang menggunakan metode valuasi dasar. Hasil analisis bukan merupakan rekomendasi investasi yang pasti.</p>
    <p><strong>Selalu lakukan riset mendalam, konsultasi dengan advisor keuangan, dan pertimbangkan toleransi risiko Anda sebelum berinvestasi.</strong></p>
    <br>
    <p style="opacity: 0.8;">💡 Dikembangkan untuk membantu investor retail dalam analisis fundamental saham</p>
</div>
""", unsafe_allow_html=True)

# Add some additional spacing
st.markdown("<br><br>", unsafe_allow_html=True)
