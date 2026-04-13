import io
import streamlit as st
from pptx import Presentation

st.set_page_config(page_title="Weekly Report Generator", layout="centered")

st.title("Weekly Report Generator")

excel_file = st.file_uploader("ارفعي ملف الداتا Excel", type=["xlsx", "xls"])
ppt_file = st.file_uploader("ارفعي الباوربوينت الريفرنس", type=["pptx"])

report_title = st.text_input("عنوان التقرير", "April 2026 – Week 1 & Week 2")

def update_title_only(ppt_bytes, report_title):
    prs = Presentation(io.BytesIO(ppt_bytes))

    if len(prs.slides) > 0:
        first_slide = prs.slides[0]
        for shape in first_slide.shapes:
            if hasattr(shape, "text_frame") and shape.has_text_frame:
                text = shape.text.strip()
                if text:
                    shape.text = f"Weekly production for {report_title}"
                    break

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

if st.button("Generate PowerPoint"):
    if excel_file is None or ppt_file is None:
        st.error("ارفعي ملف الداتا وملف الباوربوينت الريفرنس الأول")
    else:
        try:
            output_ppt = update_title_only(ppt_file.getvalue(), report_title)

            st.success("تم إنشاء نسخة أولية من الباوربوينت")
            st.download_button(
                label="Download PowerPoint",
                data=output_ppt,
                file_name="generated_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as e:
            st.error(f"حصل خطأ: {e}")
