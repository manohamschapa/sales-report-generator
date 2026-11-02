import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Sales Report Generator", page_icon="📊")
st.title("📊 Sales Report Generator")
st.write("Upload an Excel sales file and download a formatted four-sheet report.")
st.info("Expected first five columns, in order: Date, Employee, Product, Quantity, Unit Price.")

uploaded_file = st.file_uploader("Choose your input Excel file", type=["xlsx"])


def build_report(raw_df):
    if raw_df.empty or len(raw_df.columns) < 5:
        raise ValueError("The workbook must contain sales records and at least five columns.")

    df = raw_df.iloc[:, :5].copy()
    df.columns = ["Date", "Employee", "Product", "Quantity", "Unit Price"]
    df = df.dropna(how="all")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Employee"] = df["Employee"].astype("string").str.strip()
    df["Product"] = df["Product"].astype("string").str.strip()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce")

    valid = (
        df["Date"].notna() & df["Employee"].notna() & df["Employee"].ne("")
        & df["Product"].notna() & df["Product"].ne("")
        & df["Quantity"].notna() & df["Unit Price"].notna()
    )
    ignored = int((~valid).sum())
    df = df.loc[valid].copy()
    if df.empty:
        raise ValueError("No valid transactions found. Check the five expected columns.")
    if (df["Quantity"] < 0).any() or (df["Unit Price"] < 0).any():
        raise ValueError("Quantity and Unit Price cannot be negative.")

    df["Total Sales"] = df["Quantity"] * df["Unit Price"]
    revenue = float(df["Total Sales"].sum())
    count = len(df)
    units = float(df["Quantity"].sum())
    average = revenue / count if count else 0

    employee = df.groupby("Employee", as_index=False).agg(
        **{"Total Quantity": ("Quantity", "sum"), "Total Sales": ("Total Sales", "sum")}
    ).sort_values("Total Sales", ascending=False)
    product = df.groupby("Product", as_index=False).agg(
        **{"Total Quantity": ("Quantity", "sum"), "Total Sales": ("Total Sales", "sum")}
    ).sort_values("Total Sales", ascending=False)

    for table in (employee, product):
        table["Revenue Percentage (%)"] = (
            table["Total Sales"] / revenue * 100 if revenue else 0
        )

    top_employee = employee.iloc[0]["Employee"]
    top_product = product.iloc[0]["Product"]
    start_date, end_date = df["Date"].min(), df["Date"].max()
    period = start_date.strftime("%B %Y") if start_date.to_period("M") == end_date.to_period("M") else f"{start_date:%d %b %Y} – {end_date:%d %b %Y}"

    summary = pd.DataFrame({
        "Metric": [
            "Total Revenue", "Total Transactions", "Total Units Sold",
            "Average Transaction Value", "Highest Revenue Product",
            "Highest Revenue Product Sales", "Top Employee", "Top Employee Sales"
        ],
        "Value": [
            revenue, count, units, average, product.iloc[0]["Product"],
            float(product.iloc[0]["Total Sales"]), top_employee,
            float(employee.iloc[0]["Total Sales"])
        ]
    })
    transactions = df[["Date", "Employee", "Product", "Quantity", "Unit Price", "Total Sales"]]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary Report", index=False, startrow=2)
        employee.to_excel(writer, sheet_name="Employee Report", index=False, startrow=1)
        product.to_excel(writer, sheet_name="Product Sales", index=False, startrow=1)
        transactions.to_excel(writer, sheet_name="Transaction Details", index=False)

        wb = writer.book
        ws = wb["Summary Report"]
        ws["A1"] = "MONTHLY SALES REPORT"
        ws["A2"] = "Reporting Period"
        ws["B2"] = period

        for ws in wb.worksheets:
            header_row = 3 if ws.title == "Summary Report" else 2 if ws.title in ("Employee Report", "Product Sales") else 1
            ws.freeze_panes = f"A{header_row + 1}"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[header_row]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4472C4")
                cell.alignment = Alignment(horizontal="center")
            for cells in ws.columns:
                letter = get_column_letter(cells[0].column)
                width = max((len(str(c.value)) for c in cells if c.value is not None), default=10)
                ws.column_dimensions[letter].width = min(max(width + 3, 12), 38)

        ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="17365D")

        # Currency cells: Summary Value rows 4, 7, 9, 11; grouped reports col C; transactions E:F.
        for row in (4, 7, 9, 11):
            wb["Summary Report"].cell(row=row, column=2).number_format = '"₹"#,##0.00'
        for sheet in ("Employee Report", "Product Sales"):
            for row in range(3, wb[sheet].max_row + 1):
                wb[sheet].cell(row, 3).number_format = '"₹"#,##0.00'
                wb[sheet].cell(row, 4).number_format = '0.00"%"'
        tx = wb["Transaction Details"]
        for row in range(2, tx.max_row + 1):
            tx.cell(row, 1).number_format = "yyyy-mm-dd"
            tx.cell(row, 5).number_format = '"₹"#,##0.00'
            tx.cell(row, 6).number_format = '"₹"#,##0.00'

    output.seek(0)
    return output.getvalue(), {
        "revenue": revenue, "count": count, "units": units,
        "average": average, "period": period, "top_employee": top_employee,
        "top_product": top_product, "ignored": ignored
    }


if uploaded_file:
    try:
        uploaded_file.seek(0)
        input_df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.subheader("Preview of uploaded data")
        st.dataframe(input_df.head(10), use_container_width=True)
        report, stats = build_report(input_df)

        st.success("Report generated successfully.")
        if stats["ignored"]:
            st.warning(f"{stats['ignored']} incomplete/invalid row(s) were skipped.")

        a, b = st.columns(2)
        a.metric("Total Revenue", f"₹{stats['revenue']:,.2f}")
        b.metric("Transactions", f"{stats['count']:,}")
        c, d = st.columns(2)
        c.metric("Units Sold", f"{stats['units']:,.0f}")
        d.metric("Average Transaction", f"₹{stats['average']:,.2f}")
        st.write(f"**Reporting period:** {stats['period']}")
        st.write(f"**Top employee:** {stats['top_employee']}")
        st.write(f"**Highest-revenue product:** {stats['top_product']}")

        st.download_button(
            "⬇️ Download monthly_report.xlsx",
            data=report,
            file_name="monthly_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    except Exception as exc:
        st.error(f"Unable to process this workbook: {exc}")
        st.caption("Check that it is a valid .xlsx file with Date, Employee, Product, Quantity, and Unit Price in the first five columns.")
else:
    st.caption("Upload an .xlsx file to generate your report.")
