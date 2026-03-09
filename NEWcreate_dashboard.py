import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Load and prepare data
file_path = "MMS Sales Tracker.xlsx"
df = pd.read_excel(file_path)

df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
df = df.dropna(subset=['Amount', 'Sales Rep Number'])
df['Sales Rep Number'] = df['Sales Rep Number'].astype(str)

# --- Metrics ---
total_sales = df['Amount'].sum()
total_invoices = len(df)
total_sales_reps = df['Sales Rep Number'].nunique()

# --- Aggregated data for charts and table ---
rep_stats = df.groupby('Sales Rep Number').agg(
    invoice_count=('Invoice Number', 'count'),
    total_sales=('Amount', 'sum')
).reset_index().sort_values('total_sales', ascending=False)

# --- Create subplot grid ---
fig = make_subplots(
    rows=2, cols=3,
    specs=[
        [{'type': 'indicator'}, {'type': 'indicator'}, {'type': 'indicator'}],
        [{'type': 'xy'}, {'type': 'xy'}, {'type': 'table'}]
    ],
    column_widths=[0.3, 0.3, 0.4],
    row_heights=[0.3, 0.7],
    subplot_titles=('', '', '', 'Invoices by Sales Rep', 'Sales by Sales Rep', 'Summary by Sales Rep')
)

# --- Add indicators with styled backgrounds (shapes) ---
# Define colors
card_colors = ['#e3f2fd', '#e8f5e8', '#fff3e0']  # light blue, green, orange
icon_map = {1: '💰', 2: '📄', 3: '👥'}  # icons for each card

for i, (value, title, icon) in enumerate(zip(
    [total_sales, total_invoices, total_sales_reps],
    ['Total Sales', 'Total Invoices', 'Total Sales Reps'],
    ['💰', '📄', '👥']
), start=1):
    # Add indicator trace
    fig.add_trace(go.Indicator(
        mode="number+delta" if i==1 else "number",
        value=value,
        number={'prefix': '$ ' if i==1 else '', 'font': {'size': 45, 'color': '#1e3c72'}},
        title={'text': f"<b>{title}</b><br><span style='font-size:0.9em;color:#666'>{icon}</span>"},
        domain={'row': 0, 'column': i-1}
    ), row=1, col=i)

# Add background rectangles for cards (using paper coordinates)
# Card positions (relative): row1 (y from 0.7 to 1.0), columns with widths 0.3,0.3,0.4
# Column boundaries: col1: 0 to 0.3, col2: 0.3 to 0.6, col3: 0.6 to 1.0
card_x0 = [0.02, 0.32, 0.62]  # left edges with small padding
card_x1 = [0.28, 0.58, 0.98]  # right edges
card_y0, card_y1 = 0.72, 0.98  # vertical range

for i in range(3):
    fig.add_shape(
        type="rect",
        xref="paper", yref="paper",
        x0=card_x0[i], y0=card_y0, x1=card_x1[i], y1=card_y1,
        fillcolor=card_colors[i],
        line=dict(color="#ffffff", width=2),
        layer="below",
        opacity=0.9,
        name=f"card_bg_{i+1}"
    )
    # Add subtle shadow
    fig.add_shape(
        type="rect",
        xref="paper", yref="paper",
        x0=card_x0[i]+0.002, y0=card_y0-0.005, x1=card_x1[i]+0.002, y1=card_y1-0.005,
        fillcolor="rgba(0,0,0,0.1)",
        line=dict(width=0),
        layer="below",
        name=f"card_shadow_{i+1}"
    )

# --- Bar chart: Invoices by Sales Rep ---
fig.add_trace(
    go.Bar(
        x=rep_stats['Sales Rep Number'],
        y=rep_stats['invoice_count'],
        name='Invoice Count',
        marker_color='#3182bd',  # blue
        text=rep_stats['invoice_count'],
        textposition='outside',
        textfont=dict(size=11, color='#1e3c72'),
        hovertemplate='Rep: %{x}<br>Invoices: %{y}<extra></extra>'
    ),
    row=2, col=1
)

# --- Bar chart: Sales by Sales Rep ---
fig.add_trace(
    go.Bar(
        x=rep_stats['Sales Rep Number'],
        y=rep_stats['total_sales'],
        name='Total Sales ($)',
        marker_color='#e6550d',  # orange
        text=rep_stats['total_sales'].round(2),
        textposition='outside',
        texttemplate='$%{text}',
        textfont=dict(size=11, color='#8b2c0d'),
        hovertemplate='Rep: %{x}<br>Sales: $%{y:,.2f}<extra></extra>'
    ),
    row=2, col=2
)

# --- Summary table with alternating row colors ---
# Prepare cell colors: header is different, then alternate rows
header_color = '#1e3c72'
header_font_color = 'white'
odd_row_color = '#f5f7fa'
even_row_color = '#e9ecef'

cell_colors = []
for i in range(len(rep_stats)):
    cell_colors.append(odd_row_color if i % 2 == 0 else even_row_color)

fig.add_trace(
    go.Table(
        header=dict(
            values=['<b>Sales Rep</b>', '<b>Invoices</b>', '<b>Total Sales ($)</b>'],
            fill_color=header_color,
            font=dict(color=header_font_color, size=13),
            align='left',
            height=30
        ),
        cells=dict(
            values=[
                rep_stats['Sales Rep Number'],
                rep_stats['invoice_count'],
                rep_stats['total_sales'].round(2)
            ],
            fill_color=[cell_colors, cell_colors, cell_colors],
            font=dict(color='#333', size=12),
            align='left',
            format=[None, None, '.2f'],
            suffix=[None, None, ' $'],
            height=25
        )
    ),
    row=2, col=3
)

# --- Overall layout styling ---
fig.update_layout(
    title={
        'text': '📊 MMS Sales Dashboard',
        'font': {'size': 28, 'family': 'Arial, sans-serif', 'color': '#1e3c72'},
        'x': 0.5,
        'xanchor': 'center'
    },
    paper_bgcolor='#f0f2f6',
    plot_bgcolor='white',
    font=dict(family='Arial, sans-serif', size=12, color='#333'),
    height=800,
    showlegend=False,
    margin=dict(t=100, b=50, l=50, r=50)
)

# --- Style axes for bar charts ---
for col in [1, 2]:
    fig.update_xaxes(
        title_text="Sales Rep Number",
        row=2, col=col,
        tickangle=-45,
        gridcolor='lightgray',
        showline=True, linewidth=1, linecolor='lightgray'
    )
    fig.update_yaxes(
        title_text="Invoice Count" if col==1 else "Sales Amount ($)",
        row=2, col=col,
        gridcolor='lightgray',
        showline=True, linewidth=1, linecolor='lightgray'
    )

# Add a subtle background to the entire dashboard
fig.add_shape(
    type="rect",
    xref="paper", yref="paper",
    x0=0, y0=0, x1=1, y1=1,
    fillcolor="rgba(0,0,0,0)",
    line=dict(color="#ccc", width=1),
    layer="below"
)

# --- Export to HTML ---
output_file = "sales_dashboard_enhanced.html"
fig.write_html(output_file)
print(f"✨ Enhanced dashboard saved as {output_file}")