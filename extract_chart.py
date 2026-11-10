import pdfplumber
import os
import sys

def find_standalone_title(words, title):
    title_words = title.lower().split()

    for word in words:
        if word["text"].strip(",:()").lower() != title_words[0]:
            continue

        line_words = sorted(
            (
                candidate
                for candidate in words
                if abs(candidate["top"] - word["top"]) <= 3
            ),
            key=lambda candidate: candidate["x0"],
        )
        line_text = [
            candidate["text"].strip(",:()").lower()
            for candidate in line_words
        ]
        if line_text == title_words:
            return min(candidate["top"] for candidate in line_words)

    return None

pdf_path = sys.argv[1] if len(sys.argv) > 1 else "a2a testing.pdf"
if not os.path.exists(pdf_path):
    raise FileNotFoundError(f"PDF document not found: {pdf_path}")
print(f"Scanning target document: {pdf_path}")

filename_lower = os.path.basename(pdf_path).lower()
if "ltd" in filename_lower:
    benefit_type = "ltd"
elif "dental" in filename_lower:
    benefit_type = "dental"
elif "std" in filename_lower:
    benefit_type = "std"
else:
    benefit_type = "ehc"

output_folder = f"C:/automation testing/{benefit_type}_extracted_charts"
os.makedirs(output_folder, exist_ok=True)

target_charts = {
    "Plans by Province": f"{benefit_type}_plans_by_province.png",
    "Plans by Industry": f"{benefit_type}_plans_by_industry.png",
    "Benefits Provided": f"{benefit_type}_benefits_provided.png"
}

with pdfplumber.open(pdf_path) as pdf:
    print("Beginning dynamic layout spatial search across pages...")

    extracted_charts = set()
    for pn in range(len(pdf.pages)):
        page = pdf.pages[pn]
        words = page.extract_words()

        for title_phrase, output_filename in target_charts.items():
            if title_phrase in extracted_charts:
                continue

            header_top = find_standalone_title(words, title_phrase)
            if header_top is None:
                continue

            print(f"-> Found '{title_phrase}' chart on Page {pn + 1}!")

            if title_phrase == "Plans by Province":
                left, right, top_padding, height = 0.51, 0.98, 10, 205
            elif title_phrase == "Plans by Industry":
                left, right, top_padding, height = 0.07, 0.78, 10, 275
            else:
                left, right, top_padding, height = 0, 1, 5, 270

            crop_box = (
                float(page.width) * left,
                max(0, header_top - top_padding),
                float(page.width) * right,
                min(float(page.height), header_top + height),
            )

            cropped_page = page.within_bbox(crop_box)
            image_object = cropped_page.to_image(resolution=300)
            file_save_path = os.path.join(output_folder, output_filename)
            image_object.save(file_save_path, format="PNG")
            extracted_charts.add(title_phrase)
            print(f"   Saved chart image to: '{output_filename}'\n")

print("Chart snapshot asset compilation finalized successfully.")