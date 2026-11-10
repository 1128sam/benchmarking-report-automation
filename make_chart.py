import os
import re
import pandas as pd
from datetime import datetime
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_LABEL_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

base_testing_dir = r"C:/automation testing"

latest_folder_file = os.path.join(
    base_testing_dir,
    "latest_output_folder.txt"
)

if not os.path.exists(latest_folder_file):
    raise FileNotFoundError(
        "latest_output_folder.txt not found."
    )

with open(latest_folder_file, "r") as f:
    input_folder = f.read().strip()

if not os.path.exists(input_folder):
    raise FileNotFoundError(
        f"Output folder not found: {input_folder}"
    )

print(f"Reading target datasets from: {input_folder}")

current_year = datetime.now().year
company_name = "Benchmarking Client"

def get_unique_path(path):
    """Return a path that doesn't exist yet, adding (1), (2), ... before the extension."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(f"{base} ({counter}){ext}"):
        counter += 1
    return f"{base} ({counter}){ext}"

meta_path = os.path.join(input_folder, "company_meta.txt")
if os.path.exists(meta_path):
    with open(meta_path, "r") as meta_f:
        company_name = meta_f.read().strip()
        if not company_name or company_name.lower() == "unknown":
            company_name = "Benchmarking Client"

csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]
csv_files.sort(key=lambda f: int(re.search(r'\d+', f).group()) if re.search(r'\d+', f) else 0)
print(f"Building deck for '{company_name}' from {len(csv_files)} CSV files.")

# 2. Initialize PowerPoint Presentation (Widescreen 16:9)
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]

# Color Palette Definitions
DARK_BLUE = RGBColor(22, 62, 100)
LIGHT_BLUE = RGBColor(0, 160, 223)
GOLD = RGBColor(255, 192, 0)
GRAY_TEXT = RGBColor(89, 89, 89)

for slide_idx, csv_filename in enumerate(csv_files, start=1):
    file_path = os.path.join(input_folder, csv_filename)

    with open(file_path, "r") as f:
        expense_type = f.readline().strip().replace("# Expense Type: ", "")
        attribute = f.readline().strip().replace("# Attribute: ", "")
        # Preserve category labels such as "None"; pandas otherwise treats them as NA.
        df = pd.read_csv(f, comment="#", keep_default_na=False)

    df_filtered = df[df["Attribute Range Value"] != "Totals"].copy()
    
    df_filtered["Percentage of Plans"] = pd.to_numeric(df_filtered["Percentage of Plans"], errors='coerce')
    df_filtered = df_filtered.dropna(subset=["Percentage of Plans"])
    
    # prevents Excel from crashing over single numbers, empty strings, or symbols interpreted as NaN
    df_filtered["Attribute Range Value"] = df_filtered["Attribute Range Value"].astype(str).str.strip()
    df_filtered = df_filtered[df_filtered["Attribute Range Value"] != ""]

    categories_list = [str(cat).strip() for cat in df_filtered["Attribute Range Value"].tolist()]
    percentages_list = df_filtered["Percentage of Plans"].tolist()

    if not categories_list or not percentages_list:
        print(f"Skipping slide generation for {csv_filename}: Data processing left no valid chart items.")
        continue

    slide = prs.slides.add_slide(blank_layout)

    # 3. Slide Title and Subtitle Accent
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

    # 4. Prepare Native Chart Data
    chart_data = CategoryChartData()
    chart_data.categories = categories_list
    chart_data.add_series("Percentage of Plans", percentages_list)

    # 5. Position and Insert Vertical Column Chart
    x, y, cx, cy = Inches(0.8), Inches(2.2), Inches(11.5), Inches(4.5)
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    )

    chart = chart_shape.chart
    chart.has_legend = False

    # Format the Y-Axis (Percentage Scales)
    value_axis = chart.value_axis
    value_axis.minimum_scale = 0
    value_axis.has_major_gridlines = True
    value_axis.tick_labels.font.size = Pt(11)
    value_axis.tick_labels.font.color.rgb = GRAY_TEXT

    # Format the X-Axis (Categories text)
    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = Pt(11)
    category_axis.tick_labels.font.color.rgb = GRAY_TEXT
    category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW

    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.size = Pt(10)
    data_labels.font.color.rgb = GRAY_TEXT
    data_labels.number_format = '0.0"%"'

    # 6. Apply Custom Color Formatting to Data Bars
    series = chart.series[0]
    if len(categories_list) < 2:
        point = series.points[0]
        fill = point.format.fill
        fill.solid()
        fill.fore_color.rgb = DARK_BLUE
    else:
        for idx, point in enumerate(series.points):
            fill = point.format.fill
            fill.solid()

            category_name = str(categories_list[idx])
            if category_name.endswith("*"):
                fill.fore_color.rgb = GOLD
            else:
                fill.fore_color.rgb = DARK_BLUE
    # 7. Add Source and Page Metadata Footer
    footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11), Inches(0.4))
    ftf = footer_box.text_frame
    fp = ftf.paragraphs[0]
    fp.text = "Source: Sun Life Plan InSite Benchmarking, August 2026"
    fp.font.name = "Calibri"
    fp.font.size = Pt(11)
    fp.font.bold = True
    fp.font.color.rgb = DARK_BLUE

    page_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.8), Inches(1.0), Inches(0.4))
    ptf = page_box.text_frame
    pp = ptf.paragraphs[0]
    pp.text = str(slide_idx)
    pp.font.name = "Calibri"
    pp.font.size = Pt(11)
    pp.font.bold = True
    pp.font.color.rgb = DARK_BLUE

folder_name = os.path.basename(input_folder)

benefit_type = folder_name.split("_output_files_1")[0]

ppt_filename = (f"{company_name} - {benefit_type.upper()} "f"Benchmarking {current_year}.pptx")

output_ppt_path = get_unique_path(os.path.join(r"C:\automation testing", ppt_filename))

if len(prs.slides) == 0:
    raise RuntimeError("No slides were created.")
prs.save(output_ppt_path)

# Let the watcher script know exactly which file was created
with open(os.path.join(r"C:\automation testing", "last_output.txt"), "w") as f:
    f.write(output_ppt_path)