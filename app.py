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


def get_chart_shapes(slide):
    return [shape for shape in slide.shapes if getattr(shape, "has_chart", False)]


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


def trailing_zeros_to_none(values):
    result = list(values)

    last_non_zero_index = -1
    for i, v in enumerate(result):
        if v not in (None, 0, 0.0):
            last_non_zero_index = i

    if last_non_zero_index == -1:
        return [None for _ in result]

    for i in range(last_non_zero_index + 1, len(result)):
        result[i] = None

    return result


# ---------------- STRIP ----------------
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


def update_table_first_col_values(table, values):
    max_rows_to_fill = min(len(values), len(table.rows) - 1)
    for i in range(max_rows_to_fill):
        row_idx = i + 1
        val = int(values[i]) if float(values[i]).is_integer() else values[i]
        table.cell(row_idx, 0).text = str(val)


def update_strip_slides(prs, weeks, production, achieved, slag_pct, slag_kg, target_slag_pct):
    slide3 = prs.slides[2]
    slide4 = prs.slides[3]

    slide3_charts = get_chart_shapes(slide3)
    slide4_charts = get_chart_shapes(slide4)
    slide4_tables = [shape for shape in slide4.shapes if getattr(shape, "has_table", False)]

    if len(slide3_charts) < 2:
        raise ValueError("لم أجد شارتين في Slide 3.")
    if len(slide4_charts) < 1:
        raise ValueError("لم أجد شارتًا في Slide 4.")
    if len(slide4_tables) < 1:
        raise ValueError("لم أجد جدولًا في Slide 4.")

    slide3_charts.sort(key=lambda s: s.left)
    slide4_charts.sort(key=lambda s: s.left)
    slide4_tables.sort(key=lambda s: s.left)

    production_plot = trailing_zeros_to_none(production)
    achieved_plot = trailing_zeros_to_none(achieved)
    slag_pct_plot = trailing_zeros_to_none(slag_pct)

    replace_single_series_chart(slide3_charts[0].chart, weeks, "Production Roll", production_plot)
    replace_single_series_chart(slide3_charts[1].chart, weeks, "Achieved %", achieved_plot)

    replace_two_series_chart(
        slide4_charts[0].chart,
        weeks,
        "Slag %",
        slag_pct_plot,
        "Target Slag %",
        target_slag_pct
    )

    update_table_first_col_values(slide4_tables[0].table, slag_kg)


# ---------------- PASTING (Slides 5-8) ----------------
def read_pasting_data(excel_bytes):
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = get_sheet_case_insensitive(wb, "Pasting")

    if ws is None:
        raise ValueError("لا توجد شيت باسم Pasting في ملف Excel.")

    weeks = []
    produced_blocks = []
    achieved_pct = []
    strip_scrap_pct = []
    strip_scrap_target = []
    plate_scrap_pct = []
    plate_scrap_target = []
    rejected_plates_pct = []
    rejected_plates_target = []

    for row in range(2, 7):  # Week 1 to Week 5
        week = ws.cell(row=row, column=1).value
        produced = ws.cell(row=row, column=2).value
        achieved = ws.cell(row=row, column=3).value
        strip_actual = ws.cell(row=row, column=4).value
        strip_target = ws.cell(row=row, column=5).value
        plate_actual = ws.cell(row=row, column=6).value
        plate_target = ws.cell(row=row, column=7).value
        rejected_actual = ws.cell(row=row, column=8).value
        rejected_target = ws.cell(row=row, column=9).value

        weeks.append(str(week) if week else f"Week {row-1}")
        produced_blocks.append(float(produced) if produced is not None else 0)
        achieved_pct.append(float(achieved) if achieved is not None else 0)
        strip_scrap_pct.append(float(strip_actual) if strip_actual is not None else 0)
        strip_scrap_target.append(float(strip_target) if strip_target is not None else 0.3)
        plate_scrap_pct.append(float(plate_actual) if plate_actual is not None else 0)
        plate_scrap_target.append(float(plate_target) if plate_target is not None else 0.3)
        rejected_plates_pct.append(float(rejected_actual) if rejected_actual is not None else 0)
        rejected_plates_target.append(float(rejected_target) if rejected_target is not None else 0.03)

    return (
        weeks,
        produced_blocks,
        achieved_pct,
        strip_scrap_pct,
        strip_scrap_target,
        plate_scrap_pct,
        plate_scrap_target,
        rejected_plates_pct,
        rejected_plates_target,
    )


def update_pasting_slides(
    prs,
    weeks,
    produced_blocks,
    achieved_pct,
    strip_scrap_pct,
    strip_scrap_target,
    plate_scrap_pct,
    plate_scrap_target,
    rejected_plates_pct,
    rejected_plates_target,
):
    slide5 = prs.slides[4]
    slide6 = prs.slides[5]
    slide7 = prs.slides[6]
    slide8 = prs.slides[7]

    slide5_charts = get_chart_shapes(slide5)
    slide6_charts = get_chart_shapes(slide6)
    slide7_charts = get_chart_shapes(slide7)
    slide8_charts = get_chart_shapes(slide8)

    if len(slide5_charts) < 2:
        raise ValueError("لم أجد شارتين في Slide 5.")
    if len(slide6_charts) < 1:
        raise ValueError("لم أجد شارتًا في Slide 6.")
    if len(slide7_charts) < 1:
        raise ValueError("لم أجد شارتًا في Slide 7.")
    if len(slide8_charts) < 1:
        raise ValueError("لم أجد شارتًا في Slide 8.")

    slide5_charts.sort(key=lambda s: s.left)
    slide6_charts.sort(key=lambda s: s.left)
    slide7_charts.sort(key=lambda s: s.left)
    slide8_charts.sort(key=lambda s: s.left)

    produced_blocks_plot = trailing_zeros_to_none(produced_blocks)
    achieved_pct_plot = trailing_zeros_to_none(achieved_pct)
    strip_scrap_pct_plot = trailing_zeros_to_none(strip_scrap_pct)
    plate_scrap_pct_plot = trailing_zeros_to_none(plate_scrap_pct)
    rejected_plates_pct_plot = trailing_zeros_to_none(rejected_plates_pct)

    replace_single_series_chart(slide5_charts[0].chart, weeks, "Produced Blocks", produced_blocks_plot)
    replace_single_series_chart(slide5_charts[1].chart, weeks, "Achieved %", achieved_pct_plot)

    replace_two_series_chart(
        slide6_charts[0].chart,
        weeks,
        "Strip Scrap %",
        strip_scrap_pct_plot,
        "Target Strip Scrap %",
        strip_scrap_target,
    )

    replace_two_series_chart(
        slide7_charts[0].chart,
        weeks,
        "Plate Scrap %",
        plate_scrap_pct_plot,
        "Target Plate Scrap %",
        plate_scrap_target,
    )

    replace_two_series_chart(
        slide8_charts[0].chart,
        weeks,
        "Rejected Plates %",
        rejected_plates_pct_plot,
        "Target Rejected Plates %",
        rejected_plates_target,
    )


# ---------------- MAIN ----------------
if st.button("Generate PowerPoint"):
    if excel_file is None or ppt_file is None:
        st.error("ارفعي ملف Excel وملف PowerPoint الأول.")
    else:
        try:
            prs = Presentation(io.BytesIO(ppt_file.getvalue()))

            strip_values = read_strip_data(excel_file.getvalue())
            update_strip_slides(prs, *strip_values)

            pasting_values = read_pasting_data(excel_file.getvalue())
            update_pasting_slides(prs, *pasting_values)

            output = io.BytesIO()
            prs.save(output)
            output.seek(0)

            st.success("تم تحديث Slides 3-8، والخط يقف عند آخر أسبوع فيه داتا فعلية.")

            st.download_button(
                label="Download PowerPoint",
                data=output,
                file_name="generated_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except Exception as e:
            st.error(f"حصل خطأ: {e}")
