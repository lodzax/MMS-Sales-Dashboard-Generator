import streamlit as st
import pandas as pd
import plotly.express as px

# Page configuration
st.set_page_config(
    page_title="MMS Sales Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better card styling
st.markdown("""
<style>
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
    }
    .stMetric label {
        font-size: 1.1rem !important;
        color: #1e3c72 !important;
    }
    .stMetric [data-testid="stMetricValue"] {
        font-size: 2.5rem !important;
        font-weight: 700 !important;
        color: #0e2b4f !important;
    }
    h1, h2, h3 {
        color: #1e3c72;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .uploadedFile {
        background-color: #e8f0fe;
        border-radius: 5px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar for file upload and export
with st.sidebar:
    st.header("📁 Data Import")
    uploaded_file = st.file_uploader(
        "Upload MMS Sales Tracker Excel file",
        type=["xlsx", "xls"],
        help="Select the Excel file exported from MMS Sales Tracker."
    )
    st.markdown("---")
    st.markdown("### ℹ️ Instructions")
    st.markdown("""
    1. Upload your **MMS Sales Tracker.xlsx** file.
    2. The dashboard will automatically update with:
       - Key metrics
       - Charts by sales rep (stacked vertically for clarity)
       - Detailed summary table
    3. Download the summary as CSV or the full dashboard as interactive HTML.
    """)
    st.markdown("---")
    st.caption("Built with Streamlit & Plotly")

# Main area
st.title("📊 MMS Sales Dashboard")
st.markdown("---")

# Function to process uploaded file
def process_uploaded_file(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        # Clean and prepare
        df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
        df = df.dropna(subset=['Amount', 'Sales Rep Number'])
        df['Sales Rep Number'] = df['Sales Rep Number'].astype(str)
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

# Check if file is uploaded
if uploaded_file is not None:
    df = process_uploaded_file(uploaded_file)
    if df is None:
        st.stop()
else:
    st.info("👈 Please upload an MMS Sales Tracker Excel file using the sidebar to begin.")
    st.stop()

# Compute metrics
total_sales = df['Amount'].sum()
total_invoices = len(df)
total_sales_reps = df['Sales Rep Number'].nunique()

# Aggregate by sales rep
rep_stats = df.groupby('Sales Rep Number').agg(
    invoice_count=('Invoice Number', 'count'),
    total_sales=('Amount', 'sum')
).reset_index().sort_values('total_sales', ascending=False)

# ----- METRICS SECTION -----
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        label="💰 Total Sales",
        value=f"${total_sales:,.2f}",
        delta=None
    )

with col2:
    st.metric(
        label="📄 Total Invoices",
        value=f"{total_invoices:,}",
        delta=None
    )

with col3:
    st.metric(
        label="👥 Total Sales Reps",
        value=f"{total_sales_reps}",
        delta=None
    )

st.markdown("---")

# ----- CHARTS SECTION (VERTICAL LAYOUT) -----

# Chart 1: Sales by Sales Rep (top)
st.subheader("💰 Sales by Sales Rep")
fig_sales = px.bar(
    rep_stats,
    x='Sales Rep Number',
    y='total_sales',
    color_discrete_sequence=['#e6550d'],
    text=rep_stats['total_sales'].round(2),
    labels={'total_sales': 'Total Sales ($)', 'Sales Rep Number': 'Sales Rep'}
)
fig_sales.update_traces(
    texttemplate='$%{text}',
    textposition='outside',
    textfont=dict(size=11, color='#8b2c0d')
)
fig_sales.update_layout(
    height=450,
    margin=dict(l=20, r=20, t=30, b=120),  # extra bottom margin
    plot_bgcolor='white',
    hovermode='x unified',
    xaxis=dict(
        type='category',      # force categorical axis (no numeric binning)
        tickangle=-45,
        automargin=True,
        tickfont=dict(size=10)
    )
)
st.plotly_chart(fig_sales, use_container_width=True)

# Chart 2: Invoices by Sales Rep (bottom)
st.subheader("📊 Invoices by Sales Rep")
fig_invoices = px.bar(
    rep_stats,
    x='Sales Rep Number',
    y='invoice_count',
    color_discrete_sequence=['#3182bd'],
    text='invoice_count',
    labels={'invoice_count': 'Number of Invoices', 'Sales Rep Number': 'Sales Rep'}
)
fig_invoices.update_traces(textposition='outside', textfont=dict(size=11, color='#1e3c72'))
fig_invoices.update_layout(
    height=450,
    margin=dict(l=20, r=20, t=30, b=120),
    plot_bgcolor='white',
    hovermode='x unified',
    xaxis=dict(
        type='category',      # force categorical axis
        tickangle=-45,
        automargin=True,
        tickfont=dict(size=10)
    )
)
st.plotly_chart(fig_invoices, use_container_width=True)

st.markdown("---")

# ----- TABLE SECTION -----
st.subheader("📋 Sales Rep Summary")
# Format the dataframe for display
display_df = rep_stats.copy()
display_df['total_sales'] = display_df['total_sales'].map('${:,.2f}'.format)
display_df.columns = ['Sales Rep', 'Total Invoices', 'Total Sales']

# Use st.dataframe with styling
st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Sales Rep": st.column_config.TextColumn("Sales Rep"),
        "Total Invoices": st.column_config.NumberColumn("Total Invoices", format="%d"),
        "Total Sales": st.column_config.TextColumn("Total Sales"),
    }
)

# Download button for the summary
csv = rep_stats.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 Download Summary as CSV",
    data=csv,
    file_name="sales_rep_summary.csv",
    mime="text/csv"
)

# ----- EXPORT DASHBOARD AS HTML (updated layout) -----
def generate_export_html():
    # Convert rep_stats to HTML table
    table_html = rep_stats.copy()
    table_html['total_sales'] = table_html['total_sales'].map('${:,.2f}'.format)
    table_html.columns = ['Sales Rep', 'Total Invoices', 'Total Sales']
    table_html = table_html.to_html(index=False, escape=False, classes='summary-table')

    # Get Plotly figures as HTML without plotlyjs (we'll add it manually)
    invoices_html = fig_invoices.to_html(include_plotlyjs=False, full_html=False)
    sales_html = fig_sales.to_html(include_plotlyjs=False, full_html=False)

    # Build full HTML document with vertical layout
    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>MMS Sales Dashboard Export</title>
        <script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
        <style>
            body {{
                font-family: 'Arial', sans-serif;
                background-color: #f0f2f6;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
            }}
            h1 {{
                color: #1e3c72;
                text-align: center;
                font-weight: 600;
                margin-bottom: 30px;
            }}
            .metrics-container {{
                display: flex;
                justify-content: space-between;
                gap: 20px;
                margin-bottom: 30px;
            }}
            .metric-card {{
                background-color: white;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                padding: 20px;
                flex: 1;
                text-align: center;
            }}
            .metric-card.blue {{ background-color: #e3f2fd; }}
            .metric-card.green {{ background-color: #e8f5e8; }}
            .metric-card.orange {{ background-color: #fff3e0; }}
            .metric-label {{
                font-size: 1.2rem;
                color: #1e3c72;
                margin-bottom: 10px;
            }}
            .metric-value {{
                font-size: 2.5rem;
                font-weight: 700;
                color: #0e2b4f;
            }}
            .chart-container {{
                background-color: white;
                border-radius: 10px;
                padding: 15px;
                margin-bottom: 20px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .chart-container h3 {{
                margin-top: 0;
                color: #1e3c72;
                font-weight: 500;
            }}
            .table-container {{
                background-color: white;
                border-radius: 10px;
                padding: 20px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .summary-table {{
                width: 100%;
                border-collapse: collapse;
            }}
            .summary-table th {{
                background-color: #1e3c72;
                color: white;
                padding: 12px;
                text-align: left;
            }}
            .summary-table td {{
                padding: 10px 12px;
                border-bottom: 1px solid #ddd;
            }}
            .summary-table tr:nth-child(even) {{
                background-color: #f5f7fa;
            }}
            .summary-table tr:hover {{
                background-color: #e9ecef;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 MMS Sales Dashboard</h1>
            <div class="metrics-container">
                <div class="metric-card blue">
                    <div class="metric-label">💰 Total Sales</div>
                    <div class="metric-value">${total_sales:,.2f}</div>
                </div>
                <div class="metric-card green">
                    <div class="metric-label">📄 Total Invoices</div>
                    <div class="metric-value">{total_invoices:,}</div>
                </div>
                <div class="metric-card orange">
                    <div class="metric-label">👥 Total Sales Reps</div>
                    <div class="metric-value">{total_sales_reps}</div>
                </div>
            </div>
            <div class="chart-container">
                <h3>💰 Sales by Sales Rep</h3>
                {sales_html}
            </div>
            <div class="chart-container">
                <h3>📊 Invoices by Sales Rep</h3>
                {invoices_html}
            </div>
            <div class="table-container">
                <h3>📋 Sales Rep Summary</h3>
                {table_html}
            </div>
        </div>
    </body>
    </html>
    """
    return html_template

# Add export button to sidebar
with st.sidebar:
    st.markdown("---")
    st.download_button(
        label="📥 Export Dashboard as HTML",
        data=generate_export_html(),
        file_name="mms_dashboard_export.html",
        mime="text/html",
        help="Download an interactive HTML version of the dashboard (charts remain interactive)."
    )