import io
import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation

st.set_page_config(page_title="Weekly Report Generator", layout="centered")

st.title("Weekly Report Generator")

excel_file = st.file_uploader("ارفعي ملف Excel", type=["xlsx", "xls"])
ppt_file = st.file_uploader("ارفعي الباوربوينت الريفرنس", type=["pptx"])


def read_control_values(excel_bytes):
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)

    if "Control" not in wb.sheetnames:
        raise ValueError("لا توجد شيت باسم Control في ملف Excel.")

    ws = wb["Control"]

    report_title = ws["B1"].value
    slide3_title = ws["B2"].value
    slide3_body = ws["B3"].value

    return {
        "report_title": str(report_title).strip() if report_title else "",
        "slide3_title": str(slide3_title).strip() if slide3_title else "",
        "slide3_body": str(slide3_body).strip() if slide3_body else "",
    }


def get_text_shapes(slide):
    text_shapes = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text_shapes.append(shape)
    return text_shapes


def update_presentation(ppt_bytes, values):
    prs = Presentation(io.BytesIO(ppt_bytes))

    # -------- Slide 1 --------
    if len(prs.slides) >= 1:
        slide1 = prs.slides[0]
        text_shapes = get_text_shapes(slide1)

        if text_shapes and values["report_title"]:
            # يغيّر أول textbox نصي في أول سلايد
            text_shapes[0].text = values["report_title"]

    # -------- Slide 3 --------
    # ملاحظة: slide 3 يعني index 2
    if len(prs.slides) >= 3:
        slide3 = prs.slides[2]
        text_shapes = get_text_shapes(slide3)

        if len(text_shapes) >= 1 and values["slide3_title"]:
            text_shapes[0].text = values["slide3_title"]

        if len(text_shapes) >= 2 and values["slide3_body"]:
            text_shapes[1].text = values["slide3_body"]

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


if st.button("Generate PowerPoint"):
    if excel_file is None or ppt_file is None:
        st.error("ارفعي ملف Excel وملف PowerPoint الأول.")
    else:
        try:
            values = read_control_values(excel_file.getvalue())

            output_ppt = update_presentation(
                ppt_file.getvalue(),
                values
            )

            st.success("تم تعديل أول سلايد وسلايد 3 بنجاح")

            st.download_button(
                label="Download PowerPoint",
                data=output_ppt,
                file_name="generated_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except Exception as e:
            st.error(f"حصل خطأ: {e}")
