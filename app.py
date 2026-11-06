"""
World Bank Group ITS Technical Assessment — Dash implementation

Based on "Instructions - Technical Assessment.docx" and the supplied
WB_ICT_Development_Dataset.xlsx workbook.

Run:
    pip install -r requirements.txt
    python app.py

Open:
    http://127.0.0.1:8050

The workbook contains 150 country-year rows (25 countries, 2018–2023).
Missing metric values are not silently imputed. Aggregations use available
observations and display coverage; the readiness score is blank if any of its
three component indicators is missing.
"""

from pathlib import Path
import os
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, dash_table

BASE = Path(__file__).resolve().parent
SHEET = os.getenv("WB_SHEET_NAME", "WB_ICT_Development_Dataset")

# Resolve the workbook from an explicit environment variable first.
# Otherwise search the app folder and common data subfolders.
configured_path = os.getenv("WB_DATA_PATH")
if configured_path:
    DATA_PATH = Path(configured_path).expanduser()
    if not DATA_PATH.is_absolute():
        DATA_PATH = (BASE / DATA_PATH).resolve()
else:
    workbook_name = "WB_ICT_Development_Dataset.xlsx"
    candidates = [
        BASE / workbook_name,
        BASE / "data" / workbook_name,
        BASE / "dataset" / workbook_name,
        Path.cwd() / workbook_name,
        Path.cwd() / "data" / workbook_name,
        Path.cwd() / "dataset" / workbook_name,
    ]
    DATA_PATH = next((p.resolve() for p in candidates if p.is_file()), candidates[0].resolve())

if not DATA_PATH.is_file():
    raise FileNotFoundError(
        "\\nWorld Bank dataset file was not found.\\n"
        f"Expected filename: WB_ICT_Development_Dataset.xlsx\\n"
        f"Application folder: {BASE}\\n"
        "Fix: extract the COMPLETE project ZIP (not app.py alone), so the Excel "
        "workbook is beside app.py; alternatively put it in a 'data' folder, "
        "or set WB_DATA_PATH to the full workbook path.\\n"
        "Example in Windows CMD:\\n"
        'set WB_DATA_PATH=C:\\\\projectArena\\\\dondon\\\\machinelearning\\\\worldbank\\\\WB_ICT_Development_Dataset.xlsx'
    )

df = pd.read_excel(DATA_PATH, sheet_name=SHEET)
df.columns = [" ".join(str(c).strip().split()) for c in df.columns]
if "Country_Code" not in df.columns:
    df = df.rename(columns={df.columns[0]: "Country_Code"})

EXPECTED = [
    "Country_Code", "Country_Name", "Region", "Income_Group", "Year",
    "Internet_Users_Pct", "Mobile_Subscriptions_Per100",
    "Fixed_Broadband_Per100", "GDP_Per_Capita_USD", "GDP_Growth_Rate_Pct",
    "School_Enrollment_Secondary_Pct", "Govt_Expenditure_Education_Pct_GDP",
    "Digital_Skills_Workforce_Pct", "IT_Project_Completion_Rate_Pct",
    "Cybersecurity_Index_Score"
]
missing_cols = [c for c in EXPECTED if c not in df.columns]
if missing_cols:
    raise ValueError(f"Dataset is missing required columns: {missing_cols}")

METRICS = [
    "Internet_Users_Pct", "Mobile_Subscriptions_Per100", "Fixed_Broadband_Per100",
    "GDP_Per_Capita_USD", "GDP_Growth_Rate_Pct",
    "School_Enrollment_Secondary_Pct", "Govt_Expenditure_Education_Pct_GDP",
    "Digital_Skills_Workforce_Pct", "IT_Project_Completion_Rate_Pct",
    "Cybersecurity_Index_Score"
]
LABELS = {
    "Internet_Users_Pct": "Internet users (%)",
    "Mobile_Subscriptions_Per100": "Mobile subscriptions per 100",
    "Fixed_Broadband_Per100": "Fixed broadband per 100",
    "GDP_Per_Capita_USD": "GDP per capita (USD)",
    "GDP_Growth_Rate_Pct": "GDP growth (%)",
    "School_Enrollment_Secondary_Pct": "Secondary enrollment (%)",
    "Govt_Expenditure_Education_Pct_GDP": "Education spending (% GDP)",
    "Digital_Skills_Workforce_Pct": "Digital skills workforce (%)",
    "IT_Project_Completion_Rate_Pct": "IT project completion (%)",
    "Cybersecurity_Index_Score": "Cybersecurity score"
}
for c in METRICS + ["Year"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
for c in ["Country_Code", "Country_Name", "Region", "Income_Group"]:
    df[c] = df[c].astype("string").str.strip()
df = df.dropna(subset=["Country_Name", "Year"]).copy()
df["Year"] = df["Year"].astype(int)
df = df.sort_values(["Country_Name", "Year"]).reset_index(drop=True)

# Transparent complete-case score from the three weighted components.
# Clamp to the assessment's requested 30–85 display range.
raw_score = (
    .40 * df["Internet_Users_Pct"] +
    .35 * df["Digital_Skills_Workforce_Pct"] +
    .25 * df["Cybersecurity_Index_Score"]
)
complete_score = df[["Internet_Users_Pct", "Digital_Skills_Workforce_Pct",
                     "Cybersecurity_Index_Score"]].notna().all(axis=1)
df["Digital_Readiness_Score"] = np.where(
    complete_score, raw_score.clip(lower=30, upper=85), np.nan
)

NAVY = "#203864"
BLUE = "#4472C4"
TEAL = "#16A085"
GREEN = "#70AD47"
RED = "#C00000"
GOLD = "#ED7D31"
BG = "#F3F6FA"
INK = "#172B4D"
MUTED = "#65758B"
PALETTE = [NAVY, BLUE, TEAL, GOLD, GREEN, "#8064A2", "#C0504D"]
GRID = "#E5EAF1"

def records(frame):
    # Convert directly to dictionaries; avoid deprecated read_json(literal_json)
    # and normalize pandas/NumPy missing values for Dash serialization.
    clean = frame.astype(object).where(pd.notna(frame), None)
    return clean.to_dict("records")

def fmt(value, suffix=""):
    return "—" if pd.isna(value) else f"{value:,.2f}{suffix}"

def filtered(year, region, income, country):
    d = df.copy()
    if year not in (None, "ALL"):
        d = d[d.Year == int(year)]
    if region not in (None, "ALL"):
        d = d[d.Region == region]
    if income not in (None, "ALL"):
        d = d[d.Income_Group == income]
    if country not in (None, "ALL"):
        d = d[d.Country_Name == country]
    return d

def chart_layout(fig, title, height=430):
    fig.update_layout(
        template="plotly_white", title={"text": title, "x": .02, "font": {"size": 16, "color": NAVY}},
        height=height, margin={"l": 55, "r": 25, "t": 65, "b": 65},
        font={"family": "Arial", "size": 11, "color": INK},
        paper_bgcolor="white", plot_bgcolor="white",
        legend={"orientation": "h", "y": -0.2}
    )
    fig.update_xaxes(showgrid=False, automargin=True)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, automargin=True)
    return fig

def card(title, value_id, sub):
    return html.Div(className="kpi", children=[
        html.Div(title, className="kpi-label"),
        html.Div(id=value_id, className="kpi-value"),
        html.Div(sub, className="kpi-sub")
    ])

def section(title, subtitle=None):
    return html.Div([html.H2(title, className="section-title"),
                     html.P(subtitle, className="section-sub") if subtitle else None])

def table_component(columns, data, page_size=12):
    return dash_table.DataTable(
        columns=columns, data=data, page_size=page_size,
        sort_action="native", filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"fontFamily":"Arial","fontSize":11,"padding":"9px",
                    "textAlign":"left","minWidth":"100px","maxWidth":"260px",
                    "whiteSpace":"normal"},
        style_header={"backgroundColor":NAVY,"color":"white","fontWeight":"bold"},
        style_data={"borderColor":GRID},
        style_data_conditional=[
            {"if":{"row_index":"odd"},"backgroundColor":"#F7F9FC"}
        ]
    )

app = Dash(__name__, title="World Bank ITS | Technical Assessment", suppress_callback_exceptions=True)
server = app.server
app.index_string = """<!DOCTYPE html><html><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}
<style>
*{box-sizing:border-box}body{margin:0;background:#F3F6FA;color:#172B4D;font-family:Arial,sans-serif}
.header{background:#203864;color:white;padding:22px 4%;border-bottom:5px solid #16A085}
.header h1{font-size:24px;margin:0 0 6px}.header p{font-size:12px;color:#DCE6F1;margin:0}
.wrap{max-width:1500px;margin:auto;padding:20px 3% 45px}
.section-title{font-size:17px;color:#203864;margin:22px 0 6px}
.section-sub{font-size:12px;color:#65758B;margin:0 0 12px;line-height:1.5}
.filters,.panel{background:white;border:1px solid #E1E7EF;border-radius:9px;padding:15px}
.filters{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:12px}
.flab{font-size:10px;text-transform:uppercase;font-weight:bold;color:#65758B;margin-bottom:5px;display:block}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.kpi{background:white;border:1px solid #E1E7EF;border-radius:9px;padding:15px;min-height:105px}
.kpi-label{font-size:10px;color:#65758B;font-weight:bold;text-transform:uppercase}
.kpi-value{font-size:25px;font-weight:bold;color:#203864;margin:9px 0 4px}
.kpi-sub{font-size:10px;color:#65758B}
.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-top:14px}
.panel h3{font-size:14px;color:#203864;margin:0 0 8px}
.callout{background:#EAF4F8;border-left:4px solid #4472C4;border-radius:6px;padding:14px 16px;font-size:12px;line-height:1.65}
.answer{background:white;border:1px solid #E1E7EF;border-radius:9px;padding:16px;margin:12px 0}
.answer h3{font-size:14px;margin:0 0 8px;color:#203864}
.answer p{font-size:12px;line-height:1.65;color:#33445B;margin:0}
.code{background:#F3F6FA;border:1px solid #E1E7EF;padding:10px;border-radius:6px;font-family:Consolas,monospace;font-size:11px;white-space:pre-wrap;overflow-wrap:anywhere}
.tab-content{padding-top:8px}.footer{font-size:10px;color:#65758B;text-align:center;margin-top:25px}
@media(max-width:850px){.filters,.kpis,.grid2{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:560px){.filters,.kpis,.grid2{grid-template-columns:1fr}.header h1{font-size:19px}}
</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>"""

year_options = [{"label":"All years","value":"ALL"}] + [
    {"label":str(y),"value":int(y)} for y in sorted(df.Year.unique())
]
region_options = [{"label":"All regions","value":"ALL"}] + [
    {"label":r,"value":r} for r in sorted(df.Region.dropna().unique())
]
income_options = [{"label":"All income groups","value":"ALL"}] + [
    {"label":r,"value":r} for r in sorted(df.Income_Group.dropna().unique())
]
country_options = [{"label":"All countries","value":"ALL"}] + [
    {"label":r,"value":r} for r in sorted(df.Country_Name.dropna().unique())
]

def render_page(page):
    if page == "overview":
        return html.Div([
            section("Executive filters","All KPI cards and visuals respond to the selected year, region, income group and country."),
            section("Executive KPIs","The selected year defaults to 2023; choose All years to view the full period."),
            html.Div(className="kpis",children=[
                card("Average internet users (%)","kpi-internet","Mean of available observations"),
                card("Average digital readiness score","kpi-readiness","Weighted score, 30–85 display range"),
                card("Average IT project completion (%)","kpi-project","Mean of available observations"),
                card("YoY internet change (pp)","kpi-yoy","Selected year versus prior year")
            ]),
            html.Div(id="yoy-note",className="callout",style={"marginTop":"12px"}),
            html.Div(className="grid2",children=[
                html.Div(className="panel",children=[dcc.Graph(id="map-chart",config={"displayModeBar":False})]),
                html.Div(className="panel",children=[dcc.Graph(id="readiness-chart",config={"displayModeBar":False})]),
                html.Div(className="panel",children=[dcc.Graph(id="trend-chart",config={"displayModeBar":False})]),
                html.Div(className="panel",children=[dcc.Graph(id="skills-chart",config={"displayModeBar":False})])
            ]),
            section("Digital readiness score — DAX definition"),
            html.Div(className="panel",children=[
                html.Div("Calculated column (row-level):",className="flab"),
                html.Div(
"""Digital_Readiness_Score =
VAR Internet = 'WB_ICT_Development_Dataset'[Internet_Users_Pct]
VAR Skills = 'WB_ICT_Development_Dataset'[Digital_Skills_Workforce_Pct]
VAR Cyber = 'WB_ICT_Development_Dataset'[Cybersecurity_Index_Score]
VAR WeightedScore = 0.40 * Internet + 0.35 * Skills + 0.25 * Cyber
RETURN
IF(
    ISBLANK(Internet) || ISBLANK(Skills) || ISBLANK(Cyber),
    BLANK(),
    MIN(85, MAX(30, WeightedScore))
)""",className="code"),
                html.P("The score is calculated only when all three components are present. The 30–85 cap is applied to meet the assessment's stated display range; it should be disclosed because it compresses values outside that interval.",className="section-sub")
            ])
        ])
    return html.Div([
        section("Analysis & written answers","This page covers the Excel-style summaries and the three written questions in Part C of the assessment."),
        section("Part A — Excel analysis","The following tables reproduce the requested calculations in an interactive web view."),
        html.Div(className="grid2",children=[
            html.Div(className="panel",children=[
                html.H3("Average Internet Users % by income group — 2023"),
                html.Div(id="income-summary")
            ]),
            html.Div(className="panel",children=[
                html.H3("Data quality — missing values by column"),
                html.Div(id="quality-table")
            ])
        ]),
        html.Div(className="panel",style={"marginTop":"14px"},children=[
            html.H3("Pivot-style table — average IT project completion by region and year"),
            html.Div(id="pivot-table")
        ]),
        section("Part C — written responses","Each response follows the assessment's requested concise format."),
        html.Div(className="answer",children=[
            html.H3("Question 1 — Which region should be prioritized for digital skills investment?"),
            html.Div(id="answer-q1")
        ]),
        html.Div(className="answer",children=[
            html.H3("Question 2 — How were missing values identified and handled?"),
            html.P(
                f"The dataset contains {int(df[METRICS].isna().sum().sum())} missing metric cells out of "
                f"{len(df) * len(METRICS)} metric cells ({df[METRICS].isna().sum().sum()/(len(df)*len(METRICS))*100:.2f}%). "
                "I identified them with a column-level null count and missing-percentage check. "
                "I retained the original missing values rather than replacing them with invented values; "
                "each indicator average uses available observations, and the digital readiness score is left blank "
                "when any of its three components is missing. This preserves observed data but means coverage varies "
                "by indicator, so leaders should review sample counts alongside averages."
            )
        ]),
        html.Div(className="answer",children=[
            html.H3("Question 3 — What additional data source would add the most value?"),
            html.P(
                "A country-level, year-by-year dataset on public-sector ICT project delivery from the World Bank's "
                "project portfolio (including approved budget, actual expenditure, implementation status, completion "
                "date and results indicators) would add the most value. It would let ITS leadership connect digital "
                "readiness and skills gaps to the performance, cost and outcomes of actual technology investments, "
                "helping distinguish countries that need foundational capacity from those that need stronger project execution."
            )
        ]),
        section("Assessment notes"),
        html.Div(className="callout",children=[
            "Source basis: the uploaded World Bank Group ITS Technical Assessment (3 pages) and the supplied "
            "WB_ICT_Development_Dataset workbook. The assessment specifies 150 rows, 25 countries, six years "
            "(2018–2023), and approximately 5% intentionally missing values. The data supplied here contains "
            "150 rows and 10 metric columns; the missing rate is calculated over those metric cells."
        ])
    ])

# Keep both page layouts mounted in the DOM at all times. The tab control
# only changes display styles, so every callback target always exists.
shared_filters = html.Div(className="filters", children=[
    html.Div([html.Label("Year", className="flab"),
              dcc.Dropdown(id="year-filter", options=year_options, value=2023, clearable=False)]),
    html.Div([html.Label("Region", className="flab"),
              dcc.Dropdown(id="region-filter", options=region_options, value="ALL", clearable=False)]),
    html.Div([html.Label("Income group", className="flab"),
              dcc.Dropdown(id="income-filter", options=income_options, value="ALL", clearable=False)]),
    html.Div([html.Label("Country", className="flab"),
              dcc.Dropdown(id="country-filter", options=country_options, value="ALL", clearable=False)])
])

app.layout = html.Div([
    html.Div(className="header", children=[
        html.H1("WORLD BANK GROUP | ITS — DIGITAL DEVELOPMENT INSIGHTS"),
        html.P("Technical Assessment Dashboard • Country-level digital readiness and IT performance • 2018–2023")
    ]),
    html.Div(className="wrap", children=[
        dcc.Tabs(id="main-tabs", value="overview", children=[
            dcc.Tab(label="Page 1 — Executive Overview", value="overview"),
            dcc.Tab(label="Page 2 — Analysis & Written Answers", value="analysis")
        ]),
        html.Div(className="tab-content", children=[
            html.H2("DATA EXPLORER — FILTERS", className="section-title"),
            shared_filters,
            html.Div(id="overview-page", children=render_page("overview")),
            html.Div(id="analysis-page", children=render_page("analysis"), style={"display":"none"})
        ])
    ])
])

@app.callback(
    Output("overview-page", "style"),
    Output("analysis-page", "style"),
    Input("main-tabs", "value")
)
def switch_page(page):
    if page == "analysis":
        return {"display":"none"}, {"display":"block"}
    return {"display":"block"}, {"display":"none"}

@app.callback(
    Output("kpi-internet","children"), Output("kpi-readiness","children"),
    Output("kpi-project","children"), Output("kpi-yoy","children"),
    Output("yoy-note","children"), Output("map-chart","figure"),
    Output("readiness-chart","figure"), Output("trend-chart","figure"),
    Output("skills-chart","figure"),
    Input("year-filter","value"), Input("region-filter","value"),
    Input("income-filter","value"), Input("country-filter","value"),
    prevent_initial_call=False
)
def update_overview(year, region, income, country):
    d = filtered(year, region, income, country)
    internet = d["Internet_Users_Pct"].mean()
    readiness = d["Digital_Readiness_Score"].mean()
    project = d["IT_Project_Completion_Rate_Pct"].mean()

    selected_year = int(year) if year != "ALL" else int(df.Year.max())
    prev = filtered(selected_year-1, region, income, country)
    curr = filtered(selected_year, region, income, country)
    yoy = curr["Internet_Users_Pct"].mean() - prev["Internet_Users_Pct"].mean()
    yoy_text = (
        f"Year-over-year Internet Users change: {yoy:+.2f} percentage points "
        f"({selected_year} average minus {selected_year-1} average). "
        "Positive change is favorable; negative change indicates a decline. "
        "If a country/region filter is applied, both years use that same filter."
        if pd.notna(yoy) else
        f"YoY change is unavailable for {selected_year}: one of the two years has no usable Internet Users observations under the current filters."
    )

    if d["Internet_Users_Pct"].notna().any():
        m = d.dropna(subset=["Internet_Users_Pct"])
        map_fig = px.choropleth(
            m, locations="Country_Code", locationmode="ISO-3",
            color="Internet_Users_Pct", hover_name="Country_Name",
            hover_data={"Income_Group":True,"Region":True,"Internet_Users_Pct":":.1f","Country_Code":False},
            color_continuous_scale="Blues", range_color=(0,100)
        )
        map_fig.update_layout(geo={"showframe":False,"showcoastlines":True,"projection_type":"natural earth"},
                              coloraxis_colorbar={"title":"Internet users %"})
    else:
        map_fig = go.Figure()
        map_fig.add_annotation(text="No Internet Users data for this selection",showarrow=False,x=.5,y=.5)
    chart_layout(map_fig,"World map — Internet users by country")

    by_country = d.groupby("Country_Name",as_index=False)["Digital_Readiness_Score"].mean().dropna()
    by_country = by_country.sort_values("Digital_Readiness_Score").tail(12)
    readiness_fig = px.bar(by_country,x="Digital_Readiness_Score",y="Country_Name",orientation="h",
                           color="Digital_Readiness_Score",color_continuous_scale="Teal")
    chart_layout(readiness_fig,"Countries — average digital readiness score")

    trend = d.groupby("Year",as_index=False)["Internet_Users_Pct"].mean()
    trend_fig = px.line(trend,x="Year",y="Internet_Users_Pct",markers=True)
    trend_fig.update_traces(line={"color":TEAL,"width":3},marker={"size":8})
    chart_layout(trend_fig,"Internet use trend over time")

    skills = d.groupby("Region",as_index=False)["Digital_Skills_Workforce_Pct"].mean().sort_values("Digital_Skills_Workforce_Pct")
    skills_fig = px.bar(skills,x="Digital_Skills_Workforce_Pct",y="Region",orientation="h",
                        color="Digital_Skills_Workforce_Pct",color_continuous_scale="Teal")
    chart_layout(skills_fig,"Digital skills workforce by region")

    yoy_value = "—" if pd.isna(yoy) else f"{yoy:+.2f} pp"
    return fmt(internet,"%"), fmt(readiness), fmt(project,"%"), yoy_value, yoy_text, map_fig, readiness_fig, trend_fig, skills_fig

@app.callback(
    Output("income-summary","children"), Output("quality-table","children"),
    Output("pivot-table","children"), Output("answer-q1","children"),
    Input("year-filter","value"), Input("region-filter","value"),
    Input("income-filter","value"), Input("country-filter","value")
)
def update_analysis(year, region, income, country):
    # Part A task 1 is specifically 2023, and is not altered by the dashboard year slicer.
    y23 = df[df.Year == 2023]
    inc = y23.groupby("Income_Group")["Internet_Users_Pct"].agg(["mean","count"]).reset_index()
    inc.columns = ["Income group","Average Internet users %","Available records"]
    inc["Average Internet users %"] = inc["Average Internet users %"].round(2)
    inc_table = table_component(
        [{"name":c,"id":c} for c in inc.columns], records(inc), page_size=8
    )

    missing_count = df.isna().sum()
    quality = pd.DataFrame({
        "Column": df.columns,
        "Missing count": [int(missing_count[c]) for c in df.columns],
        "Missing %": [round(float(missing_count[c] / len(df) * 100), 2) for c in df.columns]
    })
    def risk(p):
        if p < 3: return "Clean"
        if p <= 10: return "Review"
        return "High Risk"
    quality["Flag"] = quality["Missing %"].map(risk)
    q_table = table_component([{"name":c,"id":c} for c in quality.columns],records(quality),page_size=16)

    pivot = df.pivot_table(index="Region",columns="Year",values="IT_Project_Completion_Rate_Pct",aggfunc="mean").round(2).reset_index()
    pivot.columns = [str(c) for c in pivot.columns]
    p_table = dash_table.DataTable(
        columns=[{"name":c,"id":c} for c in pivot.columns],data=records(pivot),
        style_table={"overflowX":"auto"},
        style_cell={"fontFamily":"Arial","fontSize":11,"padding":"9px","textAlign":"center"},
        style_header={"backgroundColor":NAVY,"color":"white","fontWeight":"bold"},
        style_data_conditional=[
            {"if":{"filter_query":f"{{{str(y)}}} = {pivot[str(y)].min()}","column_id":str(y)},"backgroundColor":"#F4CCCC","color":"#990000"}
            for y in sorted(df.Year.unique())
        ] + [
            {"if":{"filter_query":f"{{{str(y)}}} = {pivot[str(y)].max()}","column_id":str(y)},"backgroundColor":"#D9EAD3","color":"#274E13"}
            for y in sorted(df.Year.unique())
        ]
    )

    # Evidence-based recommendation uses the full 2018–2023 period for stability.
    reg = df.groupby("Region")["Digital_Skills_Workforce_Pct"].agg(["mean","count"]).sort_values("mean")
    lowest_region = reg.index[0]
    lowest_mean = reg.iloc[0]["mean"]
    second_region = reg.index[1]
    second_mean = reg.iloc[1]["mean"]
    y23_reg = df[df.Year==2023].groupby("Region")["Digital_Skills_Workforce_Pct"].agg(["mean","count"]).sort_values("mean")
    current_sentence = ""
    if lowest_region in y23_reg.index:
        current_sentence = (
            f" In 2023, {lowest_region} averaged {y23_reg.loc[lowest_region,'mean']:.2f}% "
            f"({int(y23_reg.loc[lowest_region,'count'])} available country records)."
        )
    answer = html.P(
        f"Prioritize {lowest_region} for digital-skills investment: it has the lowest average "
        f"Digital_Skills_Workforce_Pct across 2018–2023 at {lowest_mean:.2f}% "
        f"({int(reg.iloc[0]['count'])} available country-year records). "
        f"{second_region} is the next-lowest region at {second_mean:.2f}%, so it should be considered "
        f"the next priority.{current_sentence} The recommendation is based on the indicator averages in "
        "this dataset, not on a causal assessment; regional sample sizes and missing observations should "
        "be considered when allocating resources."
    )
    return inc_table, q_table, p_table, answer

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=8050)
