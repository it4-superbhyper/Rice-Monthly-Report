import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import unicodedata
from datetime import datetime
from io import BytesIO

# PDF ReportLab Imports
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm

# ---------- PDF Parser ----------
def extract_product_data_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines = []
    for page in doc:
        text = page.get_text("text")
        text = unicodedata.normalize("NFKD", text)
        lines.extend(text.splitlines())

    products = []
    i = 0
    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]

    while i < len(lines):
        line = lines[i].strip()
        if re.search(r"\b(?:\d+[Xx])?\d+(?:\.\d+)?(KG|G|GR)\b", line, re.IGNORECASE):
            name = line.strip()
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("2024 Q"):
                i += 1
            i += 1
            q2024 = []
            while i < len(lines) and len(q2024) < 12:
                q2024 += list(map(int, re.findall(r"\d+", lines[i])))
                i += 1
            while i < len(lines) and not lines[i].strip().startswith("V"):
                i += 1
            i += 1
            v2024 = []
            while i < len(lines) and len(v2024) < 12:
                v2024 += list(map(float, re.findall(r"\d+\.\d+", lines[i])))
                i += 1
            while i < len(lines) and not lines[i].strip().startswith("2025 Q"):
                i += 1
            i += 1
            q2025 = []
            while i < len(lines) and len(q2025) < 12:
                q2025 += list(map(int, re.findall(r"\d+", lines[i])))
                i += 1
            while i < len(lines) and not lines[i].strip().startswith("V"):
                i += 1
            i += 1
            v2025 = []
            while i < len(lines) and len(v2025) < 12:
                v2025 += list(map(float, re.findall(r"\d+\.\d+", lines[i])))
                i += 1
            batch_match = re.search(r"(\d+)[Xx](\d+(?:\.\d+)?)(KG|G|GR)?", name)
            if batch_match:
                units, size, unit = batch_match.groups()
                full_batch = f"{name} {size}{unit.upper() if unit else 'KG'}"
                weight_group = f"{size}{unit.upper() if unit else 'KG'}"
            else:
                single_match = re.search(r"(\d+(?:\.\d+)?)(KG|G|GR)", name)
                size = single_match.group(1) if single_match else "UNSPEC"
                unit = single_match.group(2) if single_match else "KG"
                full_batch = f"{name} {size}{unit.upper()}"
                weight_group = f"{size}{unit.upper()}"
            for j in range(12):
                products.append({
                    "Product": name,
                    "Batch": full_batch,
                    "Weight Group": weight_group,
                    "Month": months[j],
                    "Month_Num": j + 1,
                    "Year": 2024,
                    "Quantity": q2024[j] if j < len(q2024) else 0,
                    "Value": v2024[j] if j < len(v2024) else 0.0
                })
                products.append({
                    "Product": name,
                    "Batch": full_batch,
                    "Weight Group": weight_group,
                    "Month": months[j],
                    "Month_Num": j + 1,
                    "Year": 2025,
                    "Quantity": q2025[j] if j < len(q2025) else 0,
                    "Value": v2025[j] if j < len(v2025) else 0.0
                })
        else:
            i += 1

    return pd.DataFrame(products)

# ---------- PDF Export ----------
def export_dynamic_pdf(data, month_names):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()
    story = []

    def add_group(title, group_df):
        story.append(Paragraph(title, styles['Heading2']))
        story.append(Spacer(1, 6))
        grouped = group_df.groupby("Weight Group")

        for group, gdf in grouped:
            story.append(Paragraph(f"Weight Group: {group}", styles['Heading3']))
            table_data = [[
                "Batch",
                f"Q {month_names['ly']} 2024", f"Q {month_names['bl']} 2025", f"Q {month_names['lm']} 2025",
                "DLY", "DCY",
                f"V {month_names['ly']} 2024", f"V {month_names['bl']} 2025", f"V {month_names['lm']} 2025"
            ]]

            for _, row in gdf.iterrows():
                cells = [
                    row["Batch"],
                    int(row[f"Quantity {month_names['ly']} 2024"]),
                    int(row[f"Quantity {month_names['bl']} 2025"]),
                    int(row[f"Quantity {month_names['lm']} 2025"]),
                    int(row["DLY"]), int(row["DCY"]),
                    f"{row[f'Value {month_names['ly']} 2024']:.2f}",
                    f"{row[f'Value {month_names['bl']} 2025']:.2f}",
                    f"{row[f'Value {month_names['lm']} 2025']:.2f}"
                ]
                table_data.append(cells)

            totals = gdf[[
                f"Quantity {month_names['ly']} 2024",
                f"Quantity {month_names['bl']} 2025",
                f"Quantity {month_names['lm']} 2025",
                "DLY", "DCY",
                f"Value {month_names['ly']} 2024",
                f"Value {month_names['bl']} 2025",
                f"Value {month_names['lm']} 2025"
            ]].sum().round(2)

            total_row = [
                "TOTAL",
                int(totals[f"Quantity {month_names['ly']} 2024"]),
                int(totals[f"Quantity {month_names['bl']} 2025"]),
                int(totals[f"Quantity {month_names['lm']} 2025"]),
                int(totals["DLY"]), int(totals["DCY"]),
                f"{totals[f'Value {month_names['ly']} 2024']:.2f}",
                f"{totals[f'Value {month_names['bl']} 2025']:.2f}",
                f"{totals[f'Value {month_names['lm']} 2025']:.2f}"
            ]

            table_data.append(total_row)

            col_widths = [65*mm] + [20*mm]*8
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
                ('TOPPADDING', (0, 1), (-1, -1), 2),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ])

            for i, row in enumerate(gdf.itertuples(), start=1):
                if row.highlight_ly:
                    style.add('TEXTCOLOR', (6, i), (6, i), colors.red)
                if row.highlight_bl:
                    style.add('TEXTCOLOR', (7, i), (7, i), colors.red)
                if row.highlight_lm:
                    style.add('TEXTCOLOR', (8, i), (8, i), colors.red)

            table.setStyle(style)
            story.append(table)
            story.append(Spacer(1, 18))

    non_basmati_df = data[~data["Basmati_Flag"]]
    basmati_df = data[data["Basmati_Flag"]]
    add_group("📦 Non-Basmati Product Comparison", non_basmati_df)
    if not basmati_df.empty:
        add_group("🍚 Basmati Product Comparison", basmati_df)

    # --- Grand Total Section ---
    total_qty = data[f"Quantity {month_names['lm']} 2025"].sum()
    total_val = data[f"Value {month_names['lm']} 2025"].sum()

    story.append(Paragraph(f"🧾 Total Summary for {month_names['lm']} 2025", styles['Heading2']))
    story.append(Spacer(1, 6))

    summary_table = Table([
        ["Total Quantity Sold", int(total_qty)],
        ["Total Value", f"{total_val:.2f}"]
    ], colWidths=[80*mm, 30*mm])

    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))

    story.append(summary_table)
    story.append(Spacer(1, 12))

    doc.build(story)
    return buffer.getvalue()

# ---------- Streamlit App ----------
st.set_page_config(page_title="Dynamic Quantity & Value Comparison", layout="wide")
st.title("📦 Dynamic Quantity & Value Comparison")

uploaded_file = st.file_uploader("Upload Product Sales PDF", type="pdf")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    with st.spinner("Extracting data from PDF..."):
        df = extract_product_data_from_pdf(pdf_bytes)

    if df.empty:
        st.error("No data extracted. Please check the PDF format.")
    else:
        today = datetime.today()
        current_month = today.month
        last_month = current_month - 1 if current_month > 1 else 12
        before_last_month = last_month - 1 if last_month > 1 else 12
        prev_year_month = last_month

        month_name = lambda m: datetime(2000, m, 1).strftime("%B")
        month_names = {
            "ly": month_name(prev_year_month),
            "bl": month_name(before_last_month),
            "lm": month_name(last_month)
        }

        df_this_year = df[df["Year"] == 2025]
        df_last_year = df[df["Year"] == 2024]

        this_year_last_month = df_this_year[df_this_year["Month_Num"] == last_month]
        this_year_before_last = df_this_year[df_this_year["Month_Num"] == before_last_month]
        last_year_prev_month = df_last_year[df_last_year["Month_Num"] == prev_year_month]

        base = this_year_last_month.copy()
        base = base.merge(
            this_year_before_last[["Batch", "Weight Group", "Quantity", "Value"]],
            on=["Batch", "Weight Group"], how="left", suffixes=('', '_BeforeLast')
        ).merge(
            last_year_prev_month[["Batch", "Weight Group", "Quantity", "Value"]],
            on=["Batch", "Weight Group"], how="left", suffixes=('', '_LastYear')
        )

        base = base.rename(columns={
            "Quantity": f"Quantity {month_names['lm']} 2025",
            "Value": f"Value {month_names['lm']} 2025",
            "Quantity_BeforeLast": f"Quantity {month_names['bl']} 2025",
            "Value_BeforeLast": f"Value {month_names['bl']} 2025",
            "Quantity_LastYear": f"Quantity {month_names['ly']} 2024",
            "Value_LastYear": f"Value {month_names['ly']} 2024"
        })

        for col in base.columns:
            if col.startswith("Quantity") or col.startswith("Value"):
                base[col] = base[col].fillna(0)

        base["DLY"] = base[f"Quantity {month_names['lm']} 2025"] - base[f"Quantity {month_names['ly']} 2024"]
        base["DCY"] = base[f"Quantity {month_names['lm']} 2025"] - base[f"Quantity {month_names['bl']} 2025"]

        def get_min_val_flags(row):
            values = [
                row[f"Value {month_names['ly']} 2024"],
                row[f"Value {month_names['bl']} 2025"],
                row[f"Value {month_names['lm']} 2025"]
            ]
            min_val = min(values)
            return {
                "highlight_ly": values[0] == min_val,
                "highlight_bl": values[1] == min_val,
                "highlight_lm": values[2] == min_val
            }

        flags = base.apply(get_min_val_flags, axis=1, result_type="expand")
        base = pd.concat([base, flags], axis=1)

        base["Basmati_Flag"] = base["Batch"].str.contains(r"\bBASMATI\b", case=False)
        base = base.sort_values(by=["Weight Group", "Batch"]).reset_index(drop=True)

        st.subheader(f"📈 Comparison for {month_names['lm']} 2025 vs Previous Periods")
        st.dataframe(base[[
            "Weight Group", "Batch",
            f"Quantity {month_names['ly']} 2024",
            f"Quantity {month_names['bl']} 2025",
            f"Quantity {month_names['lm']} 2025",
            f"Value {month_names['ly']} 2024",
            f"Value {month_names['bl']} 2025",
            f"Value {month_names['lm']} 2025",
            "DLY", "DCY"
        ]], use_container_width=True)

        pdf_bytes = export_dynamic_pdf(base, month_names)
        st.download_button(
            label="📄 Download Quantity & Value Report (PDF)",
            data=pdf_bytes,
            file_name=f"{month_names['lm']}_comparison_report.pdf",
            mime="application/pdf"
        )
else:
    st.info("Please upload a product sales PDF file.")
