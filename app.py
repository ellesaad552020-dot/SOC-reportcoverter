import io
import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation

st.set_page_config(page_title="Weekly Report Generator", layout="centered")

st.title("Weekly Report Generator")

excel_file = st.file_uploader("ارفعي ملف Excel", type=["xlsx", "xls"])
ppt_file = st.file_uploader("ارفعي الباوربوينت الريفرنس", type=["pptx"])

manual_title = st.text_input("عنوان احتياطي لو Excel مفيهوش عنوان", "")


def extract_title_from_excel(excel_bytes):
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)

    sheets_to_check = []
    if "Control" in wb.sheetnames:
        sheets_to_check.append(wb["Control"])

    sheets_to_check.append(wb[wb.sheetnames[0]])

    labels = {
        "report title",
        "title",
        "عنوان التقرير",
        "report_title",
    }

    for ws in sheets_to_check:
        for row in range(1, 21):
            for col in range(1, 6):
                cell_value = ws.cell(row=row, column=col).value
                if isinstance(cell_value, str) and cell_value.strip().lower() in labels:
                    right_value = ws.cell(row=row, column=col + 1).value
                    if right_value:
                        return str(right_value).strip()

    for ws in sheets_to_check:
        b1_value = ws["B1"].value
        if b1_value:
            return str(b1_value).strip()

    return None


def update_first_slide_title(ppt_bytes, new_title):
    prs = Presentation(io.BytesIO(ppt_bytes))

    if len(prs.slides) == 0:
        raise ValueError("ملف الباوربوينت لا يحتوي على أي سلايدات.")

    first_slide = prs.slides[0]
    target_shape = None

    keywords = [
        "weekly",
        "production",
        "month",
        "الانتاج",
        "الإنتاج",
        "الاسبوع",
        "الأسبوع",
    ]

    for shape in first_slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = shape.text.strip()
            if text and any(keyword in text.lower() for keyword in keywords):
                target_shape = shape
                break

    if target_shape is None:
        for shape in first_slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = shape.text.strip()
                if text:
                    target_shape = shape
                    break

    if target_shape is None:
        raise ValueError("لم أجد عنوانًا نصيًا في أول سلايد.")

    target_shape.text = new_title

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


if st.button("Generate PowerPoint"):
    if excel_file is None or ppt_file is None:
        st.error("ارفعي ملف Excel وملف PowerPoint الأول.")
    else:
        try:
            title_from_excel = extract_title_from_excel(excel_file.getvalue())
            final_title = title_from_excel or manual_title.strip()

            if not final_title:
                st.error("لم أجد عنوانًا في Excel. اكتبيه في Control!B1 أو اكتبيه يدويًا.")
            else:
                st.success(f"العنوان المقروء: {final_title}")

                output_ppt = update_first_slide_title(
                    ppt_file.getvalue(),
                    final_title
                )

                st.download_button(
                    label="Download PowerPoint",
                    data=output_ppt,
                    file_name="generated_report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

        except Exception as e:
            st.error(f"حصل خطأ: {e}")
