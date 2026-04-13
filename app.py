import io
import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData

st.set_page_config(page_title="Weekly Report Generator", layout="centered")

st.title("Weekly Report Generator")

excel_file = st.file_uploader("ارفعي ملف Excel", type=["xlsx", "xls"])
ppt_file = st.file_uploader("ارفعي الباوربوينت الريفرنس", type=["pptx"])


def read_strip_data(excel_bytes):
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)

    if "Strip" not in wb.sheetnames:
        raise ValueError("لا توجد شيت باسم Strip في ملف Excel.")

    ws = wb["Strip"]

    weeks = []
    production = []
    achieved = []

    for row in range(2, 7):  # Week 1 to Week 5
        week = ws.cell(row=row, column=1).value
        prod = ws.cell(row=row, column=2).value
        ach = ws.cell(row=row, column=3).value

        weeks.append(str(week) if week else f"Week {row-1}")
        production.append(float(prod) if prod is not None else 0)
        achieved.append(float(ach) if ach is not None else 0)

    return weeks, production, achieved


def replace_chart_data(chart, categories, series_name, values):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)
    chart.replace_data(chart_data)


def update_strip_slide_only(ppt_bytes, weeks, production, achieved):
    prs = Presentation(io.BytesIO(ppt_bytes))

    # Slide 4 = index 3
    if len(prs.slides) < 4:
        raise ValueError("ملف الباوربوينت لا يحتوي على سلايد 4.")

    slide = prs.slides[3]

    charts = []
    for shape in slide.shapes:
        if shape.has_chart:
            charts.append(shape.chart)

    if len(charts) < 2:
        raise ValueError("لم أجد شارتين في سلايد Strip.")

    # أول شارت = Production Roll
    replace_chart_data(charts[0], weeks, "Production Roll", production)

    # ثاني شارت = Achieved %
    replace_chart_data(charts[1], weeks, "Achieved %", achieved)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


if st.button("Generate PowerPoint"):
    if excel_file is None or ppt_file is None:
        st.error("ارفعي ملف Excel وملف PowerPoint الأول.")
    else:
        try:
            weeks, production, achieved = read_strip_data(excel_file.getvalue())

            output_ppt = update_strip_slide_only(
                ppt_file.getvalue(),
                weeks,
                production,
                achieved
            )

            st.success("تم تحديث شارتات سلايد Strip فقط بدون تعديل العناوين أو التنسيق.")

            st.download_button(
                label="Download PowerPoint",
                data=output_ppt,
                file_name="generated_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except Exception as e:
            st.error(f"حصل خطأ: {e}")
