import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Load the Excel file
file_path = "MMS Sales Tracker.xlsx"
df = pd.read_excel(file_path)

# Clean and prepare data
df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
df = df.dropna(subset=['Amount', 'Sales Rep Number'])
df['Sales Rep Number'] = df['Sales Rep Number'].astype(str)

# --- Compute metrics ---
total_sales = df['Amount'].sum()
total_invoices = len(df)
total_sales_reps = df['Sales Rep Number'].nunique()

# --- Prepare data for "by Sales Rep" charts and table ---
rep_stats = df.groupby('Sales Rep Number').agg(
    invoice_count=('Invoice Number', 'count'),
    total_sales=('Amount', 'sum')
).reset_index().sort_values('total_sales', ascending=False)

# --- Create the dashboard figure with subplots ---
# Layout: 2 rows, 3 columns
# Row 1: three indicator metrics
# Row 2: col1 = invoices bar chart, col2 = sales bar chart, col3 = summary table
fig = make_subplots(
    rows=2, cols=3,
    specs=[
        [{'type': 'indicator'}, {'type': 'indicator'}, {'type': 'indicator'}],
        [{'type': 'xy'}, {'type': 'xy'}, {'type': 'table'}]
    ],
    column_widths=[0.3, 0.3, 0.4],  # allocate more width to table
    row_heights=[0.3, 0.7],
    subplot_titles=('', '', '', 'Invoices by Sales Rep', 'Sales by Sales Rep', 'Summary by Sales Rep')
)

# --- Add indicator traces ---
fig.add_trace(go.Indicator(
    mode="number",
    value=total_sales,
    title={"text": "Total Sales<br><span style='font-size:0.8em;color:gray'>$</span>"},
    number={"prefix": "$", "font": {"size": 40}},
), row=1, col=1)

fig.add_trace(go.Indicator(
    mode="number",
    value=total_invoices,
    title={"text": "Total Invoices"},
    number={"font": {"size": 40}},
), row=1, col=2)

fig.add_trace(go.Indicator(
    mode="number",
    value=total_sales_reps,
    title={"text": "Total Sales Reps"},
    number={"font": {"size": 40}},
), row=1, col=3)

# --- Add bar chart: Invoices by Sales Rep ---
fig.add_trace(
    go.Bar(
        x=rep_stats['Sales Rep Number'],
        y=rep_stats['invoice_count'],
        name='Invoice Count',
        marker_color='royalblue',
        text=rep_stats['invoice_count'],
        textposition='outside'
    ),
    row=2, col=1
)

# --- Add bar chart: Sales by Sales Rep ---
fig.add_trace(
    go.Bar(
        x=rep_stats['Sales Rep Number'],
        y=rep_stats['total_sales'],
        name='Total Sales ($)',
        marker_color='firebrick',
        text=rep_stats['total_sales'].round(2),
        textposition='outside',
        texttemplate='$%{text}'
    ),
    row=2, col=2
)

# --- Add summary table ---
fig.add_trace(
    go.Table(
        header=dict(
            values=['Sales Rep', 'Total Invoices', 'Total Sales ($)'],
            fill_color='paleturquoise',
            align='left',
            font=dict(size=12)
        ),
        cells=dict(
            values=[
                rep_stats['Sales Rep Number'],
                rep_stats['invoice_count'],
                rep_stats['total_sales'].round(2)
            ],
            fill_color='lavender',
            align='left',
            format=[None, None, '.2f'],
            suffix=[None, None, ' $']
        )
    ),
    row=2, col=3
)

# --- Update layout and axes ---
fig.update_layout(
    title_text="MMS Sales Dashboard",
    title_font_size=24,
    height=800,
    showlegend=False  # legends are redundant for these separate charts
)

# Set y-axis labels for the two bar charts
fig.update_yaxes(title_text="Invoice Count", row=2, col=1)
fig.update_yaxes(title_text="Sales Amount ($)", row=2, col=2)
fig.update_xaxes(title_text="Sales Rep Number", row=2, col=1)
fig.update_xaxes(title_text="Sales Rep Number", row=2, col=2)

# Export to interactive HTML
output_file = "sales_dashboard.html"
fig.write_html(output_file)
print(f"Dashboard saved as {output_file}")