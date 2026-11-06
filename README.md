# World Bank ICT Development Dataset (1990--2026)

## Overview

The **World Bank ICT Development Dataset** is an expanded country-year
dataset covering **25 countries from 1990 to 2026**. It contains **925
records** and **10 indicators** related to digital connectivity,
economic development, education, digital skills, IT project completion,
and cybersecurity.

The dataset is intended for data exploration, comparative analysis,
visualization, and development of interactive dashboards using tools
such as Python, Pandas, Plotly, and Dash.

## Dataset at a Glance

  Attribute              Description
  ---------------------- -----------------
  Countries              25
  Years covered          1990--2026
  Country-year records   925
  Indicators             10
  Frequency              Annual
  Main format            Excel (`.xlsx`)

## Indicators

The dataset includes the following indicators:

1.  Internet users (%)
2.  Mobile subscriptions per 100 people
3.  Fixed broadband subscriptions per 100 people
4.  GDP per capita (USD)
5.  GDP growth rate (%)
6.  Secondary school enrollment (%)
7.  Government expenditure on education (% of GDP)
8.  Digital skills workforce (%)
9.  IT project completion rate (%)
10. Cybersecurity index score

## Workbook Structure

The Excel workbook contains three worksheets:

-   **`WB_ICT_Development_Dataset`** --- the expanded dataset with
    annual country-year records from 1990 through 2026.
-   **`Data_Notes`** --- documentation describing the coverage, data
    status, modeling approach, and limitations.
-   **`Original_2018_2023`** --- a reference copy of the originally
    supplied 2018--2023 dataset.

## Important Data Quality and Methodology Note

The workbook combines original supplied data with modeled values. These
categories are identified in the `Data_Status` column:

  Period         Number of records Status
  ------------ ------------------- ------------------------------
  1990--2017                   700 Modeled historical estimates
  2018--2023                   150 Original supplied dataset
  2024--2026                    75 Projections
  **Total**                **925** 

The 1990--2017 values were modeled backward from the supplied country
observations. The 2024--2026 values were projected forward from the
latest supplied observations.

**The modeled historical estimates and projections are not verified
official World Bank observations.** They are included to support
exploratory analysis and dashboard development. Do not present them as
official historical statistics or use them as a substitute for verified
source data in policy, academic, or formal reporting.

Missing values in the original 2018--2023 records have been retained
rather than silently replaced. Consult the `Data_Notes` worksheet and
the `Data_Status` and `Data_Method` columns before interpreting results.

## Getting Started with Python

Install the packages:

``` bash
pip install pandas openpyxl matplotlib plotly dash
```

Load the dataset with Pandas:

``` python
import pandas as pd

file_path = "WB_ICT_Development_Dataset_1990_2026.xlsx"

df = pd.read_excel(
    file_path,
    sheet_name="WB_ICT_Development_Dataset"
)

print(df.head())
print(df.shape)
print(df["Year"].min(), df["Year"].max())
```

For analysis using only the originally supplied records:

``` python
original_df = df[df["Data_Status"] == "Original supplied dataset"]
```

To inspect modeled historical estimates and projections separately:

``` python
historical_estimates = df[
    df["Data_Status"] == "Modeled historical estimate"
]

projections = df[
    df["Data_Status"] == "Projection"
]
```

## Potential Uses

-   Compare ICT and digital connectivity indicators across countries.
-   Explore changes in internet use, mobile subscriptions, and broadband
    access.
-   Examine relationships between economic, education, and digital
    development indicators.
-   Compare digital skills, IT project completion, and cybersecurity
    measures.
-   Build interactive dashboards and data visualizations.
-   Demonstrate data preparation, exploratory analysis, and reporting
    workflows.

## Repository Contents

The repository is intended to contain the Excel dataset and its
supporting documentation. Refer to the repository files for the current
downloadable workbook and any additional analysis or dashboard
resources.

## Disclaimer

This is an analytical dataset assembled from the supplied workbook and
modeled extensions. The extended years are estimates or projections, not
independently verified World Bank data. Verify indicators against their
official source series before using them for research conclusions,
policy decisions, or formal publications.

## Author

**Edmond Dondon**

-   GitHub: [dondonedmond82](https://github.com/dondonedmond82)

------------------------------------------------------------------------
