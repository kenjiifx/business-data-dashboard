from __future__ import annotations

from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "data" / "business_data.csv"
OUTPUT_DIR = BASE_DIR / "output"
EXCEL_OUTPUT = OUTPUT_DIR / "business_data_dashboard.xlsx"
INSIGHTS_OUTPUT = OUTPUT_DIR / "actionable_insights.txt"


def clean_and_structure_data(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    df.columns = [column.strip().lower() for column in df.columns]
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["region"] = df["region"].astype(str).str.strip().str.title()
    df["product"] = df["product"].astype(str).str.strip().str.title()
    df["channel"] = df["channel"].astype(str).str.strip().str.title()

    numeric_columns = ["revenue", "cost", "units_sold", "customer_satisfaction"]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(subset=["date", "region", "product", "channel", "revenue", "cost", "units_sold"])
    df["customer_satisfaction"] = df["customer_satisfaction"].fillna(df["customer_satisfaction"].median())

    df["profit"] = df["revenue"] - df["cost"]
    df["profit_margin_pct"] = (df["profit"] / df["revenue"] * 100).round(2)
    df["year_month"] = df["date"].dt.to_period("M").astype(str)

    return df.sort_values("date").reset_index(drop=True)


def build_performance_metrics(df: pd.DataFrame) -> dict[str, float]:
    metrics = {
        "total_revenue": float(df["revenue"].sum()),
        "total_cost": float(df["cost"].sum()),
        "total_profit": float(df["profit"].sum()),
        "avg_profit_margin_pct": float(df["profit_margin_pct"].mean()),
        "avg_customer_satisfaction": float(df["customer_satisfaction"].mean()),
        "total_units_sold": float(df["units_sold"].sum()),
    }
    return metrics


def generate_pivot_tables(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    pivot_region = (
        pd.pivot_table(
            df,
            index="region",
            values=["revenue", "profit", "units_sold", "customer_satisfaction"],
            aggfunc={
                "revenue": "sum",
                "profit": "sum",
                "units_sold": "sum",
                "customer_satisfaction": "mean",
            },
        )
        .round(2)
        .reset_index()
    )

    pivot_product = (
        pd.pivot_table(
            df,
            index="product",
            values=["revenue", "profit", "units_sold", "customer_satisfaction"],
            aggfunc={
                "revenue": "sum",
                "profit": "sum",
                "units_sold": "sum",
                "customer_satisfaction": "mean",
            },
        )
        .round(2)
        .reset_index()
    )

    monthly_trends = (
        pd.pivot_table(
            df,
            index="year_month",
            values=["revenue", "profit", "units_sold"],
            aggfunc="sum",
        )
        .round(2)
        .reset_index()
    )

    return pivot_region, pivot_product, monthly_trends


def generate_actionable_insights(
    metrics: dict[str, float], pivot_region: pd.DataFrame, pivot_product: pd.DataFrame
) -> str:
    best_region = pivot_region.sort_values("profit", ascending=False).iloc[0]
    weakest_region = pivot_region.sort_values("profit", ascending=True).iloc[0]
    best_product = pivot_product.sort_values("profit", ascending=False).iloc[0]

    insights = [
        "Business Data Analysis Dashboard | Python, Excel | Mar 2026",
        "",
        f"1) Total revenue is ${metrics['total_revenue']:,.0f} with total profit of ${metrics['total_profit']:,.0f}, indicating a healthy average margin of {metrics['avg_profit_margin_pct']:.2f}%.",
        f"2) {best_region['region']} is the top-performing region by profit (${best_region['profit']:,.0f}); replicate its commercial strategy in lower-performing regions.",
        f"3) {weakest_region['region']} has the lowest regional profit (${weakest_region['profit']:,.0f}); prioritize pricing, discount governance, and cost controls there.",
        f"4) {best_product['product']} is the strongest product by profit (${best_product['profit']:,.0f}); increase inventory and marketing allocation for this line.",
        f"5) Average customer satisfaction is {metrics['avg_customer_satisfaction']:.2f}/5.00; maintain product quality while targeting service uplift in underperforming segments.",
    ]

    return "\n".join(insights)


def build_excel_dashboard(
    df: pd.DataFrame,
    metrics: dict[str, float],
    pivot_region: pd.DataFrame,
    pivot_product: pd.DataFrame,
    monthly_trends: pd.DataFrame,
    output_path: Path,
) -> None:
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        pivot_region.to_excel(writer, sheet_name="Pivot_Region", index=False)
        pivot_product.to_excel(writer, sheet_name="Pivot_Product", index=False)
        monthly_trends.to_excel(writer, sheet_name="Monthly_Trends", index=False)

        workbook = writer.book
        dashboard_sheet = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dashboard_sheet

        header_format = workbook.add_format({"bold": True, "font_size": 12})
        currency_format = workbook.add_format({"num_format": "$#,##0"})
        percent_format = workbook.add_format({"num_format": "0.00%"})
        decimal_format = workbook.add_format({"num_format": "0.00"})

        dashboard_sheet.write("A1", "Business Data Analysis Dashboard", header_format)
        dashboard_sheet.write("A3", "Total Revenue", header_format)
        dashboard_sheet.write("B3", metrics["total_revenue"], currency_format)
        dashboard_sheet.write("A4", "Total Profit", header_format)
        dashboard_sheet.write("B4", metrics["total_profit"], currency_format)
        dashboard_sheet.write("A5", "Average Profit Margin", header_format)
        dashboard_sheet.write("B5", metrics["avg_profit_margin_pct"] / 100, percent_format)
        dashboard_sheet.write("A6", "Average Customer Satisfaction", header_format)
        dashboard_sheet.write("B6", metrics["avg_customer_satisfaction"], decimal_format)

        trend_sheet = writer.sheets["Monthly_Trends"]
        region_sheet = writer.sheets["Pivot_Region"]

        trend_chart = workbook.add_chart({"type": "line"})
        trend_chart.add_series(
            {
                "name": "Revenue Trend",
                "categories": ["Monthly_Trends", 1, 0, len(monthly_trends), 0],
                "values": ["Monthly_Trends", 1, 1, len(monthly_trends), 1],
            }
        )
        trend_chart.set_title({"name": "Monthly Revenue Trend"})
        trend_chart.set_y_axis({"major_gridlines": {"visible": False}})

        region_chart = workbook.add_chart({"type": "column"})
        region_chart.add_series(
            {
                "name": "Profit by Region",
                "categories": ["Pivot_Region", 1, 0, len(pivot_region), 0],
                "values": ["Pivot_Region", 1, 2, len(pivot_region), 2],
            }
        )
        region_chart.set_title({"name": "Regional Profit Comparison"})
        region_chart.set_y_axis({"major_gridlines": {"visible": False}})

        dashboard_sheet.insert_chart("D3", trend_chart, {"x_scale": 1.2, "y_scale": 1.2})
        dashboard_sheet.insert_chart("D20", region_chart, {"x_scale": 1.2, "y_scale": 1.2})

        dashboard_sheet.set_column("A:A", 30)
        dashboard_sheet.set_column("B:B", 18)

        trend_sheet.set_column("A:D", 18)
        region_sheet.set_column("A:E", 18)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    cleaned_df = clean_and_structure_data(DATA_PATH)
    metrics = build_performance_metrics(cleaned_df)
    pivot_region, pivot_product, monthly_trends = generate_pivot_tables(cleaned_df)

    build_excel_dashboard(
        df=cleaned_df,
        metrics=metrics,
        pivot_region=pivot_region,
        pivot_product=pivot_product,
        monthly_trends=monthly_trends,
        output_path=EXCEL_OUTPUT,
    )

    insights = generate_actionable_insights(metrics, pivot_region, pivot_product)
    INSIGHTS_OUTPUT.write_text(insights, encoding="utf-8")

    print(f"Dashboard created: {EXCEL_OUTPUT}")
    print(f"Insights created: {INSIGHTS_OUTPUT}")


if __name__ == "__main__":
    main()
