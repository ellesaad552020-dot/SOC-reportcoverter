import io
import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData

st.set_page_config(page_title="Weekly Report Generator", layout="centered")
st.title("Weekly Report Generator")

excel_file = st.file_uploader("ارفعي ملف Excel", type=["xlsx"])
ppt_file = st.file_uploader("ارفعي الباوربوينت الريفرنس", type=["pptx"])


def get_sheet_case_insensitive(wb, target_name):
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() == target_name.strip().lower():
            return wb[sheet_name]
    return None


def read_strip_data(excel_bytes):
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = get_sheet_case_insensitive(wb, "Strip")

    if ws is None:
        raise ValueError("لا توجد شيت باسم Strip في ملف Excel.")

    weeks = []
    production = []
    achieved = []
    slag_pct = []
    slag_kg = []
    target_slag_pct = []

    for row in range(2, 7):  # Week 1 to Week 5
        week = ws.cell(row=row, column=1).value
        prod = ws.cell(row=row, column=2).value
        ach = ws.cell(row=row, column=3).value
        slg_pct = ws.cell(row=row, column=4).value
        slg_kg_val = ws.cell(row=row, column=5).value
        target_val = ws.cell(row=row, column=6).value

        weeks.append(str(week) if week else f"Week {row-1}")
        production.append(float(prod) if prod is not None else 0)
        achieved.append(float(ach) if ach is not None else 0)
        slag_pct.append(float(slg_pct) if slg_pct is not None else 0)
        slag_kg.append(float(slg_kg_val) if slg_kg_val is not None else 0)
        target_slag_pct.append(float(target_val) if target_val is not None else 2.8)

    return weeks, production, achieved, slag_pct, slag_kg, target_slag_pct


def replace_single_series_chart(chart, categories, series_name, values):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)
    chart.replace_data(chart_data)


def replace_two_series_chart(chart, categories, series1_name, values1, series2_name, values2):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series1_name, values1)
    chart_data.add_series(series2_name, values2)
    chart.replace_data(chart_data)


def get_chart_shapes(slide):
    return [shape for shape in slide.shapes if getattr(shape, "has_chart", False)]


def get_table_shapes(slide):
    return [shape for shape in slide.shapes if getattr(shape, "has_table", False)]


def set_cell_text(cell, value):
    cell.text = str(value)


def update_first_table_numeric_rows(table, values):
    rows_count = len(table.rows)
    cols_count = len(table.columns)

    if cols_count < 1:
        raise ValueError("الجدول لا يحتوي على أعمدة.")

    max_rows_to_fill = min(5, rows_count - 1)

    for i in range(max_rows_to_fill):
        row_idx = i + 1
        val = int(values[i]) if float(values[i]).is_integer() else values[i]
        set_cell_text(table.cell(row_idx, 0), val)


def update_strip_slides(ppt_bytes, weeks, production, achieved, slag_pct, slag_kg, target_slag_pct):
    prs = Presentation(io.BytesIO(ppt_bytes))

    if len(prs.slides) < 4:
        raise ValueError("ملف الباوربوينت لا يحتوي على عدد السلايدات المتوقع.")

    slide3 = prs.slides[2]
    slide4 = prs.slides[3]

    slide3_charts = get_chart_shapes(slide3)
    slide4_charts = get_chart_shapes(slide4)
    slide4_tables = get_table_shapes(slide4)

    if len(slide3_charts) < 2:
        raise ValueError("لم أجد شارتين في Slide 3.")
    if len(slide4_charts) < 1:
        raise ValueError("لم أجد شارتًا في Slide 4.")
    if len(slide4_tables) < 1:
        raise ValueError("لم أجد جدولًا في Slide 4.")

    slide3_charts.sort(key=lambda s: s.left)
    slide4_charts.sort(key=lambda s: s.left)
    slide4_tables.sort(key=lambda s: s.left)

    # Slide 3
    replace_single_series_chart(slide3_charts[0].chart, weeks, "Production Roll", production)
    replace_single_series_chart(slide3_charts[1].chart, weeks, "Achieved %", achieved)

    # Slide 4
    replace_two_series_chart(
        slide4_charts[0].chart,
        weeks,
        "Slag %",
        slag_pct,
        "Target Slag %",
        target_slag_pct
    )

    # Slide 4 table
    update_first_table_numeric_rows(slide4_tables[0].table, slag_kg)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


if st.button("Generate PowerPoint"):
    if excel_file is None or ppt_file is None:
        st.error("ارفعي ملف Excel وملف PowerPoint الأول.")
    else:
        try:
            weeks, production, achieved, slag_pct, slag_kg, target_slag_pct = read_strip_data(excel_file.getvalue())

            output_ppt = update_strip_slides(
                ppt_file.getvalue(),
                weeks,
                production,
                achieved,
                slag_pct,
                slag_kg,
                target_slag_pct,
            )

            st.success("تم تحديث شارتات Strip وجدول Slide 4 فقط بدون تعديل العناوين أو التنسيق أو KPIs.")

            st.download_button(
                label="Download PowerPoint",
                data=output_ppt,
                file_name="generated_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except Exception as e:
            st.error(f"حصل خطأ: {e}")
