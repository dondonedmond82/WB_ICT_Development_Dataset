"""World Bank ICT Development Dashboard

Run:
    pip install -r requirements.txt
    python app.py

Keep violence_data.xlsx in the same folder as this file. The filename is
retained for compatibility with the user's current project folder; the
workbook itself contains World Bank ICT indicators, not VAWG survey data.
Open http://127.0.0.1:8050
"""
from pathlib import Path
import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, State, dash_table

BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = Path(os.getenv("ICT_DATA_PATH", BASE_DIR / "violence_data.xlsx"))
SHEET_NAME = os.getenv("ICT_SHEET_NAME", "WB_ICT_Development_Dataset")

REQUIRED_COLUMNS = [
    "Country_Code", "Country_Name", "Region", "Income_Group", "Year",
    "Internet_Users_Pct", "Mobile_Subscriptions_Per100",
    "Fixed_Broadband_Per100", "GDP_Per_Capita_USD", "GDP_Growth_Rate_Pct",
    "School_Enrollment_Secondary_Pct", "Govt_Expenditure_Education_Pct_GDP",
    "Digital_Skills_Workforce_Pct", "IT_Project_Completion_Rate_Pct",
    "Cybersecurity_Index_Score", "Data_Status", "Data_Method",
]

if not DATA_PATH.exists():
    raise FileNotFoundError(
        f"Dataset not found: {DATA_PATH}\n"
        "Place violence_data.xlsx beside app.py or set ICT_DATA_PATH."
    )
if DATA_PATH.suffix.lower() in {".xlsx", ".xls"}:
    try:
        df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)
    except ValueError as exc:
        raise ValueError(
            f"Worksheet '{SHEET_NAME}' was not found in {DATA_PATH.name}. "
            "Set ICT_SHEET_NAME to a valid worksheet name."
        ) from exc
elif DATA_PATH.suffix.lower() == ".csv":
    df = pd.read_csv(DATA_PATH)
else:
    raise ValueError("Supported dataset formats are .xlsx, .xls and .csv.")

missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
if missing:
    raise ValueError(
        "The selected file does not match the World Bank ICT dataset schema.\n"
        f"Missing columns: {missing}\nFound columns: {list(df.columns)}"
    )

NUMERIC_COLUMNS = [c for c in REQUIRED_COLUMNS if c not in {
    "Country_Code", "Country_Name", "Region", "Income_Group", "Data_Status", "Data_Method"
}]
for col in NUMERIC_COLUMNS:
    df[col] = pd.to_numeric(df[col], errors="coerce")
df["Year"] = df["Year"].astype("Int64")
df = df.dropna(subset=["Year", "Country_Name"]).copy()
df["Year"] = df["Year"].astype(int)
df = df.sort_values(["Country_Name", "Year"])

METRICS = {
    "Internet users (%)": "Internet_Users_Pct",
    "Mobile subscriptions (per 100 people)": "Mobile_Subscriptions_Per100",
    "Fixed broadband (per 100 people)": "Fixed_Broadband_Per100",
    "GDP per capita (US$)": "GDP_Per_Capita_USD",
    "GDP growth (%)": "GDP_Growth_Rate_Pct",
    "Secondary school enrollment (%)": "School_Enrollment_Secondary_Pct",
    "Government education spending (% GDP)": "Govt_Expenditure_Education_Pct_GDP",
    "Digital skills workforce (%)": "Digital_Skills_Workforce_Pct",
    "IT project completion (%)": "IT_Project_Completion_Rate_Pct",
    "Cybersecurity index score": "Cybersecurity_Index_Score",
}
METRIC_LABELS = {v: k for k, v in METRICS.items()}
NAVY = "#14213D"
BLUE = "#3977B8"
TEAL = "#0F8B8D"
ORANGE = "#D97706"
GREEN = "#27856A"
RED = "#C94C4C"
PURPLE = "#8064A2"
LIGHT = "#F3F6FA"
MUTED = "#687386"
PALETTE = [NAVY, TEAL, ORANGE, BLUE, GREEN, PURPLE, RED]

app = Dash(__name__, title="World Bank ICT Development Dashboard")
server = app.server


def options_for(col):
    vals = sorted(df[col].dropna().unique().tolist(), key=lambda x: str(x))
    return [{"label": "All", "value": "__ALL__"}] + [
        {"label": str(v), "value": str(v)} for v in vals
    ]


def card(label, value, sub=""):
    return html.Div([
        html.Div(label, className="kpi-label"),
        html.Div(value, className="kpi-value"),
        html.Div(sub, className="kpi-sub"),
    ], className="kpi-card")


def base_layout(fig, title, height=360):
    fig.update_layout(
        template="plotly_white", title={"text": title, "x": 0.025, "xanchor": "left",
        "font": {"size": 15, "color": NAVY}}, height=height,
        margin={"l": 48, "r": 24, "t": 62, "b": 55},
        font={"family": "Arial, sans-serif", "size": 11, "color": NAVY},
        paper_bgcolor="white", plot_bgcolor="white",
        legend={"orientation": "h", "y": -0.22},
    )
    fig.update_xaxes(showgrid=False, automargin=True)
    fig.update_yaxes(showgrid=True, gridcolor="#E8ECF2", zeroline=False, automargin=True)
    return fig


def empty_fig(message="No records match the selected filters."):
    fig = go.Figure()
    fig.add_annotation(text=message, x=.5, y=.5, xref="paper", yref="paper", showarrow=False)
    return base_layout(fig, "No data available")


def filtered_data(region, income, country, status, years):
    data = df.copy()
    if region != "__ALL__":
        data = data[data.Region.astype(str) == region]
    if income != "__ALL__":
        data = data[data.Income_Group.astype(str) == income]
    if country != "__ALL__":
        data = data[data.Country_Name.astype(str) == country]
    if status != "__ALL__":
        data = data[data.Data_Status.astype(str) == status]
    if years and len(years) == 2:
        data = data[data.Year.between(int(years[0]), int(years[1]))]
    return data


min_year, max_year = int(df.Year.min()), int(df.Year.max())
status_options = options_for("Data_Status")

app.layout = html.Div([
    html.Header([
        html.Div([
            html.Div("WORLD BANK | DIGITAL DEVELOPMENT", className="eyebrow"),
            html.H1("ICT Development Dashboard"),
            html.P("Explore connectivity, digital capacity and supporting development indicators across countries and years."),
        ]),
        html.Div([html.Span("DATASET", className="header-tag"),
                  html.Strong(f"{df.Country_Name.nunique()} countries · {min_year}–{max_year}")], className="header-meta")
    ], className="hero"),
    html.Div([
        html.Aside([
            html.H3("Filters"),
            html.Label("Region"),
            dcc.Dropdown(id="region-filter", options=options_for("Region"), value="__ALL__", clearable=False),
            html.Label("Income group"),
            dcc.Dropdown(id="income-filter", options=options_for("Income_Group"), value="__ALL__", clearable=False),
            html.Label("Country"),
            dcc.Dropdown(id="country-filter", options=options_for("Country_Name"), value="__ALL__", clearable=False),
            html.Label("Data status"),
            dcc.Dropdown(id="status-filter", options=status_options, value="__ALL__", clearable=False),
            html.Label("Year range"),
            dcc.RangeSlider(id="year-filter", min=min_year, max=max_year, step=1,
                value=[min_year, max_year], marks={y: str(y) for y in range(min_year, max_year + 1, 5)} | {max_year: str(max_year)},
                tooltip={"placement": "bottom", "always_visible": False}),
            html.Div(id="filter-count", className="filter-count"),
            html.Button("Reset filters", id="reset-filters", n_clicks=0, className="reset-button"),
        ], className="sidebar"),
        html.Main([
            html.Div(id="kpi-row", className="kpi-grid"),
            html.Div([
                html.Div([
                    html.Label("Indicator for trend and comparison"),
                    dcc.Dropdown(id="metric-filter", options=[{"label": k, "value": v} for k, v in METRICS.items()],
                                 value="Internet_Users_Pct", clearable=False),
                ], className="metric-picker"),
                html.Div(id="metric-note", className="metric-note"),
            ], className="metric-toolbar"),
            html.Div([
                dcc.Graph(id="trend-chart", config={"displayModeBar": False}),
                dcc.Graph(id="region-chart", config={"displayModeBar": False}),
            ], className="chart-grid"),
            html.Div([
                dcc.Graph(id="scatter-chart", config={"displayModeBar": False}),
                dcc.Graph(id="heatmap-chart", config={"displayModeBar": False}),
            ], className="chart-grid"),
            html.Section([
                html.Div([html.H3("Country-year records"),
                          html.Button("Download filtered data (CSV)", id="download-button", className="download-button")], className="table-heading"),
                dash_table.DataTable(id="data-table", page_size=12, sort_action="native", filter_action="native",
                    style_table={"overflowX": "auto"},
                    style_header={"backgroundColor": "#EEF2F7", "fontWeight": "700", "color": NAVY, "border": "0"},
                    style_cell={"padding": "9px", "fontFamily": "Arial", "fontSize": "11px", "textAlign": "left", "border": "0", "borderBottom": "1px solid #EDF0F4", "minWidth": "105px", "maxWidth": "220px", "whiteSpace": "normal"},
                    style_data_conditional=[{"if": {"row_index": "odd"}, "backgroundColor": "#FAFBFD"}]),
                dcc.Download(id="download-data"),
            ], className="table-panel"),
            html.Div([
                html.Strong("Data interpretation: "),
                "The workbook includes original 2018–2023 observations and modeled historical estimates / projections for other years. Check Data Status and Data Method before interpreting trends. Modeled values are not official observed statistics."
            ], className="disclaimer"),
        ], className="main-content")
    ], className="dashboard-body")
], className="app-shell")


@app.callback(
    Output("region-filter", "value"), Output("income-filter", "value"),
    Output("country-filter", "value"), Output("status-filter", "value"),
    Output("year-filter", "value"), Input("reset-filters", "n_clicks"),
    prevent_initial_call=True,
)
def reset_filters(_):
    return "__ALL__", "__ALL__", "__ALL__", "__ALL__", [min_year, max_year]


@app.callback(
    Output("kpi-row", "children"), Output("filter-count", "children"),
    Output("metric-note", "children"), Output("trend-chart", "figure"),
    Output("region-chart", "figure"), Output("scatter-chart", "figure"),
    Output("heatmap-chart", "figure"), Output("data-table", "data"),
    Output("data-table", "columns"),
    Input("region-filter", "value"), Input("income-filter", "value"),
    Input("country-filter", "value"), Input("status-filter", "value"),
    Input("year-filter", "value"), Input("metric-filter", "value"),
)
def update_dashboard(region, income, country, status, years, metric):
    data = filtered_data(region, income, country, status, years)
    metric_label = METRIC_LABELS.get(metric, metric)
    count_text = f"{len(data):,} records · {data.Country_Name.nunique()} countries"
    if data.empty:
        blank = empty_fig()
        return [card("Countries", "0"), card("Years", "—"), card("Selected indicator", "—"), card("Regions", "0")], count_text, "No records available for this selection.", blank, blank, blank, blank, [], []

    metric_values = data[metric].dropna()
    avg = metric_values.mean() if not metric_values.empty else float("nan")
    latest_year = int(data.Year.max())
    latest = data[data.Year == latest_year][metric].mean()
    avg_text = f"{avg:,.2f}" if pd.notna(avg) else "N/A"
    latest_text = f"{latest:,.2f}" if pd.notna(latest) else "N/A"
    kpis = [
        card("Countries", f"{data.Country_Name.nunique():,}", "In current selection"),
        card("Regions", f"{data.Region.nunique():,}", "Represented in selection"),
        card("Year coverage", f"{data.Year.min()}–{data.Year.max()}", f"{data.Year.nunique()} annual periods"),
        card(f"Average · {metric_label}", avg_text, f"Latest year ({latest_year}): {latest_text}"),
    ]
    note = f"Selected indicator: {metric_label}. Values are averaged across available country-year records in the current filter selection; missing values are excluded."

    trend_data = data.groupby(["Year", "Region"], as_index=False)[metric].mean()
    trend = px.line(trend_data, x="Year", y=metric, color="Region", markers=True,
                    color_discrete_sequence=PALETTE)
    trend.update_layout(yaxis_title=metric_label, xaxis_title="Year")
    trend = base_layout(trend, f"{metric_label} · trend by region", 370)

    region_data = data.groupby("Region", as_index=False)[metric].mean().sort_values(metric)
    region_fig = px.bar(region_data, x=metric, y="Region", orientation="h", color="Region",
                         color_discrete_sequence=PALETTE, text_auto=".2s")
    region_fig.update_layout(xaxis_title=metric_label, yaxis_title=None, showlegend=False)
    region_fig = base_layout(region_fig, f"Regional comparison · {metric_label}", 370)

    scatter_data = data.dropna(subset=["GDP_Per_Capita_USD", metric]).copy()
    if len(scatter_data) > 3000:
        scatter_data = scatter_data.sample(3000, random_state=42)
    scatter = px.scatter(scatter_data, x="GDP_Per_Capita_USD", y=metric, color="Income_Group",
                         hover_data=["Country_Name", "Year", "Region", "Data_Status"],
                         color_discrete_sequence=PALETTE, opacity=.72)
    scatter.update_layout(xaxis_title="GDP per capita (US$)", yaxis_title=metric_label)
    scatter = base_layout(scatter, f"GDP per capita vs. {metric_label}", 390)

    heat_data = data.groupby(["Region", "Year"], as_index=False)[metric].mean()
    pivot = heat_data.pivot(index="Region", columns="Year", values=metric)
    heat = go.Figure(data=go.Heatmap(z=pivot.values, x=pivot.columns, y=pivot.index,
                                     colorscale="Blues", colorbar={"title": metric_label}))
    heat = base_layout(heat, f"Regional heatmap · {metric_label}", 390)

    display_cols = ["Country_Code", "Country_Name", "Region", "Income_Group", "Year", metric, "Data_Status"]
    table = data[display_cols].sort_values(["Year", "Country_Name"], ascending=[False, True]).copy()
    table[metric] = table[metric].round(2)
    table = table.where(pd.notna(table), None)
    columns = [{"name": (metric_label if c == metric else c.replace("_", " ")), "id": c} for c in display_cols]
    return kpis, count_text, note, trend, region_fig, scatter, heat, table.to_dict("records"), columns


@app.callback(
    Output("download-data", "data"), Input("download-button", "n_clicks"),
    State("region-filter", "value"), State("income-filter", "value"),
    State("country-filter", "value"), State("status-filter", "value"),
    State("year-filter", "value"), prevent_initial_call=True,
)
def download_filtered(_, region, income, country, status, years):
    data = filtered_data(region, income, country, status, years)
    return dcc.send_data_frame(data.to_csv, "world_bank_ict_filtered.csv", index=False)


app.index_string = """<!DOCTYPE html>
<html><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}
<style>
*{box-sizing:border-box}body{margin:0;background:#F3F6FA;color:#14213D;font-family:Arial,sans-serif}.app-shell{min-height:100vh}.hero{background:#14213D;color:white;padding:28px 34px;display:flex;justify-content:space-between;align-items:center;gap:20px}.eyebrow{font-size:10px;letter-spacing:1.8px;color:#8DD2D0;font-weight:700}.hero h1{font-size:27px;margin:8px 0 7px}.hero p{margin:0;color:#D6DEEA;font-size:13px}.header-meta{text-align:right;min-width:180px}.header-meta strong{display:block;margin-top:8px;font-size:13px}.header-tag{font-size:9px;letter-spacing:1.5px;color:#9DAFC7}.dashboard-body{display:grid;grid-template-columns:260px minmax(0,1fr);gap:20px;padding:20px;max-width:1700px;margin:auto}.sidebar,.table-panel,.kpi-card,.metric-toolbar,.chart-grid>div{background:white;border:1px solid #E5EAF1;border-radius:10px;box-shadow:0 2px 8px rgba(20,33,61,.035)}.sidebar{padding:19px;height:max-content;position:sticky;top:16px}.sidebar h3{margin:0 0 18px;font-size:16px}.sidebar label,.metric-picker label{display:block;font-size:11px;font-weight:700;margin:14px 0 7px;color:#4E5B70}.sidebar .Select-control{border-color:#DDE3EC}.filter-count{font-size:11px;color:#687386;margin:20px 0 12px}.reset-button,.download-button{border:0;border-radius:6px;background:#14213D;color:white;padding:10px 13px;font-size:11px;font-weight:700;cursor:pointer}.reset-button{width:100%;background:#EEF2F7;color:#14213D}.main-content{min-width:0}.kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:15px}.kpi-card{padding:16px}.kpi-label{font-size:10px;color:#687386;font-weight:700}.kpi-value{font-size:24px;font-weight:700;margin:9px 0 4px;color:#14213D}.kpi-sub{font-size:10px;color:#8A94A5}.metric-toolbar{padding:13px 16px;margin-bottom:14px}.metric-picker{max-width:450px}.metric-picker label{margin:0 0 7px}.metric-note{font-size:10px;color:#687386;margin-top:8px}.chart-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:14px}.chart-grid>div{min-width:0;overflow:hidden}.chart-grid .dash-graph{width:100%}.table-panel{padding:17px;margin-top:4px}.table-heading{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:13px}.table-heading h3{margin:0;font-size:15px}.disclaimer{font-size:11px;line-height:1.6;color:#5D687A;background:#EAF0F6;border-left:3px solid #0F8B8D;padding:12px 14px;margin-top:14px;border-radius:4px}.disclaimer strong{color:#14213D}
@media(max-width:1050px){.dashboard-body{grid-template-columns:1fr}.sidebar{position:static}.kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:700px){.hero{padding:22px;display:block}.header-meta{text-align:left;margin-top:18px}.dashboard-body{padding:12px}.chart-grid{grid-template-columns:1fr}.kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.table-heading{align-items:flex-start;flex-direction:column}}
</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>"""

if __name__ == "__main__":
    app.run(debug=os.getenv("DASH_DEBUG", "false").lower() == "true", host="127.0.0.1", port=int(os.getenv("PORT", "8050")))
