import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_LABEL_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

DARK_BLUE = RGBColor(22, 62, 100)
LIGHT_BLUE = RGBColor(0, 160, 223)
GOLD = RGBColor(255, 192, 0)
GRAY_TEXT = RGBColor(89, 89, 89)

# Streamlit App Configurations
st.set_page_config(page_title="Benchmark Report Generator", layout="centered")
st.title("Benchmark Report Generator")
st.write("Upload a benchmarking PDF report below to automatically generate your presentation.")

uploaded_file = st.file_uploader("Choose a Benchmarking PDF file", type=["pdf"])

if uploaded_file is not None:
    filename_lower = uploaded_file.name.lower()
    
    # 1. Infer Benefit Type from File Name
    if "ltd" in filename_lower:
        benefit_type = "ltd"
    elif "dental" in filename_lower:
        benefit_type = "dental"
    elif "std" in filename_lower:
        benefit_type = "std"
    else:
        benefit_type = "ehc"

    st.info(f"Processing target file as **{benefit_type.upper()}** layout...")

    # 2. Extract Data from PDF Bytes Buffer
    all_pages_data = []
    company_name = "Benchmarking Client"

    with pdfplumber.open(uploaded_file) as pdf:
        first_page_text = pdf.pages[0].extract_text().split("\n")
        for i, line in enumerate(first_page_text):
            if line.startswith("Active as of:") and (i + 1 < len(first_page_text)):
                company_name = first_page_text[i + 1].strip()
                break
        
        if not company_name or company_name.lower() == "unknown":
            company_name = "Benchmarking Client"

        # Double check benefit type inside text content
        for line in first_page_text:
            if "Benefit:" in line:
                b_text = line.split("Benefit:")[-1].lower()
                if "disability" in b_text or "ltd" in b_text:
                    benefit_type = "ltd"
                elif "dental" in b_text:
                    benefit_type = "dental"
                elif "std" in b_text:
                    benefit_type = "std"

        # Parse data tables page by page
        for pn, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            table_rows = []
            is_inside_table = False        
            expense_type = "Unknown Expense"
            attribute = "Unknown Attribute"

            for line in raw_text.split("\n"):
                if "Expense Type:" in line:
                    expense_type = line.split("Expense Type:")[-1].split("|")[0].replace(" / All", "").strip()
                if "Attribute:" in line:
                    attribute = line.split("Attribute:")[-1].split("|")[0].strip()
                if "Attribute Range Value" in line:
                    is_inside_table = True
                    continue
                if "Totals" in line and is_inside_table:
                    cleaned_total = re.sub(r'\s{2,}', '  ', line.strip())
                    table_rows.append(cleaned_total.rsplit(maxsplit=3))
                    break
                if is_inside_table:
                    cleaned_line = line.strip()
                    if not cleaned_line:
                        continue
                    cleaned_line = re.sub(r'\s{2,}', '  ', cleaned_line)
                    parts = cleaned_line.rsplit(maxsplit=3)
                    if len(parts) == 4:
                        table_rows.append(parts)

            if table_rows:
                columns = ["Attribute Range Value", "Number of Contracts", "Number of Plans", "Percentage of Plans"]
                df = pd.DataFrame(table_rows, columns=columns)
                all_pages_data.append({
                    "page_num": pn + 1,
                    "expense_type": expense_type,
                    "attribute": attribute,
                    "df": df
                })

    # 3. Generate Native PPTX File in Memory Buffer
    if all_pages_data:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        for slide_idx, page_data in enumerate(all_pages_data, start=1):
            df = page_data["df"]
            expense_type = page_data["expense_type"]
            attribute = page_data["attribute"]

            df_filtered = df[df["Attribute Range Value"] != "Totals"].copy()
            df_filtered["Percentage of Plans"] = pd.to_numeric(df_filtered["Percentage of Plans"], errors='coerce')
            df_filtered = df_filtered.dropna(subset=["Percentage of Plans"])
            df_filtered["Attribute Range Value"] = df_filtered["Attribute Range Value"].astype(str).str.strip()
            df_filtered = df_filtered[df_filtered["Attribute Range Value"] != ""]

            categories_list = [str(cat).strip() for cat in df_filtered["Attribute Range Value"].tolist()]
            percentages_list = df_filtered["Percentage of Plans"].tolist()

            if not categories_list or not percentages_list:
                continue

            slide = prs.slides.add_slide(blank_layout)

            # Slide Header
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11), Inches(1.5))
            tf = title_box.text_frame
            tf.word_wrap = True

            p_title = tf.paragraphs[0]
            p_title.text = f"{attribute} - {expense_type}"
            p_title.font.name = "Calibri"
            p_title.font.size = Pt(40)
            p_title.font.bold = True
            p_title.font.color.rgb = DARK_BLUE
            p_title.space_after = Pt(20)

            p_sub = tf.add_paragraph()
            p_sub.text = f"% of Plans Offering {expense_type}"
            p_sub.font.name = "Calibri"
            p_sub.font.size = Pt(18)
            p_sub.font.color.rgb = LIGHT_BLUE

            chart_data = CategoryChartData()
            chart_data.categories = categories_list
            chart_data.add_series("Percentage of Plans", percentages_list)

            x, y, cx, cy = Inches(0.8), Inches(2.2), Inches(11.5), Inches(4.5)
            chart_shape = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
            chart = chart_shape.chart
            chart.has_legend = False

            chart.value_axis.minimum_scale = 0
            chart.value_axis.has_major_gridlines = True
            chart.value_axis.tick_labels.font.size = Pt(11)
            chart.value_axis.tick_labels.font.color.rgb = GRAY_TEXT

            chart.category_axis.tick_labels.font.size = Pt(11)
            chart.category_axis.tick_labels.font.color.rgb = GRAY_TEXT
            chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW

            plot = chart.plots[0]
            plot.has_data_labels = True
            plot.data_labels.font.size = Pt(10)
            plot.data_labels.font.color.rgb = GRAY_TEXT
            plot.data_labels.number_format = '0.0"%"'

            series = chart.series[0]
            for idx, point in enumerate(series.points):
                fill = point.format.fill
                fill.solid()
                if str(categories_list[idx]).endswith("*"):
                    fill.fore_color.rgb = GOLD
                else:
                    fill.fore_color.rgb = DARK_BLUE

            # Structural Layout Footers
            footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11), Inches(0.4))
            footer_box.text_frame.paragraphs[0].text = "Source: Sun Life Plan InSite Benchmarking, August 2026"
            footer_box.text_frame.paragraphs[0].font.size = Pt(11)
            footer_box.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE

            page_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.8), Inches(1.0), Inches(0.4))
            page_box.text_frame.paragraphs[0].text = str(slide_idx)
            page_box.text_frame.paragraphs[0].font.size = Pt(11)
            page_box.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE

        # Save active presentation out to bytes stream
        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)

        current_year = datetime.now().year
        ppt_filename = f"{company_name} - {benefit_type.upper()} Benchmarking {current_year}.pptx"

        st.success("Slide generation completed.")
        
        # 4. Interactive File Download Trigger Button
        st.download_button(
            label="Download Generated PowerPoint Deck",
            data=ppt_buffer,
            file_name=ppt_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    else:
        st.error("Could not parse data out of this document format structure.")