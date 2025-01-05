import os
import re
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import numpy as np
from PIL import Image, ImageDraw, ImageFont

# Path to Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# Directory containing PDFs
pdf_dir = r"C:\\Users\\Work\\Documents\\jobHunt\\Projects\\costcoTester\\PDFs"
output_dir = r"C:\\Users\\Work\\Documents\\jobHunt\\Projects\\costcoTester\\ExcelFiles"
annotated_dir = r"C:\\Users\\Work\\Documents\\jobHunt\\Projects\\costcoTester\\AnnotatedImages"

# Ensure the output and annotated directories exist
os.makedirs(output_dir, exist_ok=True)
os.makedirs(annotated_dir, exist_ok=True)

# Function to fix merged text (e.g., "7.99N" -> "7.99 N")
def fix_merged_text(text):
    return re.sub(r'(\d+\.?\d*)([A-Za-z])', r'\1 \2', text)

# Function to preprocess image (convert to grayscale)
def preprocess_image(image):
    # Convert the image to grayscale for better OCR accuracy
    gray_image = image.convert("L")
    
    # Save the processed image for debugging
    return gray_image

# Function to parse a line of text based on conditions
def parse_text(line):
    pattern_membership = r'^(E)\s+(\d+)\s+(.+?)\s+(\d+\.\d{2})\s+([A-Za-z])$'
    pattern_regular = r'^(\d+)\s+(.+?)\s+(\d+\.\d{2})\s+([A-Za-z])$'
    pattern_discount = r'^(\d+)\s+(.+?)\s+(\d+\.\d{2}-)$'

    if match := re.match(pattern_membership, line):
        return match.group(1), match.group(2), match.group(3), match.group(4), match.group(5)
    elif match := re.match(pattern_regular, line):
        return None, match.group(1), match.group(2), match.group(3), match.group(4)
    elif match := re.match(pattern_discount, line):
        return None, match.group(1), match.group(2), f"-{match.group(3)[:-1]}", None
    else:
        return None, None, None, None, None

# Ensure consistent data types in the Price column
def clean_price(value):
    try:
        return float(value.replace("-", "")) * (-1 if "-" in value else 1)
    except (ValueError, AttributeError):
        return np.nan

# Process each PDF in the directory
for pdf_file in os.listdir(pdf_dir):
    if pdf_file.endswith(".pdf"):
        pdf_path = os.path.join(pdf_dir, pdf_file)
        output_excel = os.path.join(output_dir, os.path.splitext(pdf_file)[0] + ".xlsx")
        annotated_pdf_dir = os.path.join(annotated_dir, os.path.splitext(pdf_file)[0])
        os.makedirs(annotated_pdf_dir, exist_ok=True)

        # Convert PDF to images
        try:
            images = convert_from_path(pdf_path, dpi=200)
        except Exception as e:
            print(f"An error occurred while converting {pdf_file} to images: {e}")
            continue

        # Extract text data
        store_data = []  # To store data before Line 8 on Page 1
        data = []  # To store table data
        purchase_data = []  # To store data after "SUBTOTAL"
        date_data_list = []  # To store lines starting with a date
        stop_processing = False  # Flag to stop adding rows to the table

        last_sku = None
        last_description = None

        for page_number, image in enumerate(images, start=1):
            try:
                processed_image = preprocess_image(image)  # Preprocess image (grayscale)
                draw = ImageDraw.Draw(processed_image)
                font = ImageFont.load_default()  # Default font for annotation

                lines = pytesseract.image_to_string(processed_image, config="--oem 1 --psm 6").splitlines()
                
                for line_number, line in enumerate(lines, start=1):
                    if line.strip():
                        cleaned_line = fix_merged_text(line.strip())
                        print(f"Page {page_number}, Line {line_number}: {cleaned_line}")  # Log OCR output

                        # Annotate the image with the OCR text
                        draw.text((10, 10 + 15 * line_number), f"{line_number}: {cleaned_line}", fill="blue", font=font)

                        if page_number == 1 and line_number <= 8: #will need to revise section
                            store_data.append([page_number, line_number, cleaned_line])
                            continue

                        if re.match(r'\d{2}/\d{2}/\d{4}', cleaned_line):
                            date_data_list.append([page_number, line_number, cleaned_line])

                        if stop_processing:
                            purchase_data.append([page_number, line_number, cleaned_line])
                        elif cleaned_line.startswith("SUBTOTAL"):
                            stop_processing = True
                            purchase_data.append([page_number, line_number, cleaned_line])
                        else:
                            membership, sku, description, price, tax = parse_text(cleaned_line)
                            
                            if price and float(price) < 0:
                                sku = last_sku
                                description = last_description
                                tax = "Discount"
                            

                            if price and float(price) >= 0:
                                last_sku = sku
                                last_description = description

                            data.append([page_number, line_number, cleaned_line, membership, sku, description, price, tax])

                # Save annotated image
                annotated_image_path = os.path.join(annotated_pdf_dir, f"Page_{page_number}.png")
                processed_image.save(annotated_image_path)
                print(f"Annotated page saved as {annotated_image_path}")
            except Exception as e:
                print(f"An error occurred while processing page {page_number} of {pdf_file}: {e}")

        # Create DataFrames and Excel output as previously specified
        store_lines = [3, 4, 5]
        store_table = pd.DataFrame(
            [[1, 3, ", ".join([row[2] for row in store_data if row[1] in store_lines])]],
            columns=["Page", "Line", "Text"]
        )

        columns = ["Page", "Line", "Text", "Membership", "SKU#", "Description", "Price", "Tax Exemption"]
        df = pd.DataFrame(data, columns=columns)

        # Filter purchase data
        purchase_keywords = ["SUBTOTAL", "TAX", "AMOUNT:", "INSTANT SAVINGS", "TOTAL NUMBER OF ITEMS SOLD"]
        purchase_table = pd.DataFrame(
            [row for row in purchase_data if any(row[2].startswith(keyword) for keyword in purchase_keywords)],
            columns=["Page", "Line", "Text"]
        )

        # Extract category and amount from purchase data
        purchase_table[["Category", "Amount"]] = purchase_table["Text"].str.extract(r'(.*?)(\$?[\d\.,-]+)')

        # Create Date Data DataFrame
        date_data = pd.DataFrame(date_data_list, columns=["Page", "Line", "Text"])
        date_data["Date"] = date_data["Text"].str.extract(r'(\d{2}/\d{2}/\d{4})')
        date_data["Time"] = date_data["Text"].str.extract(r'(\d{2}:\d{2})')

        # Ensure consistent data types in the Price column
        df["Price"] = df["Price"].apply(clean_price)

        # Consolidate Store Data into a single cell
        store_table = pd.DataFrame({"Store Information": [store_table.iloc[0, 2]]})
        # Print Store Information
        print("Store Information:")
        print(store_table)

        # Print Itemized Purchase Data
        print("\nItemized Purchase Data:")
        print(df.drop(columns=["Page", "Line", "Text"], inplace=False))

        # Print Purchase Data with Categories
        print("\nPurchase Data with Categories:")
        print(purchase_table.drop(columns=["Page", "Line", "Text"], inplace=False))

        # Print Date Data
        print("\nDate Data:")
        print(date_data.drop(columns=["Page", "Line", "Text"], inplace=False))

        # Save to Excel with multiple sheets and other stuff
    
        try:
            with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
                store_table.to_excel(writer, index=False, sheet_name="Store Data")
                df.drop(columns=["Page", "Line", "Text"], inplace=False).to_excel(
                    writer, index=False, sheet_name="Itemized Purchase Data"
                )
                purchase_table.drop(columns=["Page", "Line", "Text"], inplace=False).to_excel(
                    writer, index=False, sheet_name="Purchase Data with Categories"
                )
                date_data.drop(columns=["Page", "Line", "Text"], inplace=False).to_excel(
                    writer, index=False, sheet_name="Date Data"
                )
            
            print(f"Processed {pdf_file} and saved to {output_excel} with multiple sheets")
        except Exception as e:
            print(f"An error occurred while saving {output_excel}: {e}")
