import fitz  # PyMuPDF
import pandas as pd
import re
import os

def extract_name_and_passkey_improved(page):
    """
    Extract name and passkey from a single PDF page with improved targeting
    """
    # Get text with positioning information
    text_dict = page.get_text("dict")

    name = ""
    passkey = ""

    # Get page dimensions
    page_height = page.rect.height
    page_width = page.rect.width

    all_text_items = []

    # Extract all text with positions
    for block in text_dict["blocks"]:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text_content = span["text"].strip()
                    if text_content:
                        bbox = span["bbox"]
                        all_text_items.append({
                            'text': text_content,
                            'x': bbox[0],
                            'y': bbox[1],
                            'width': bbox[2] - bbox[0],
                            'height': bbox[3] - bbox[1],
                            'font_size': span.get('size', 0)
                        })

    # Sort by y-position (top to bottom)
    all_text_items.sort(key=lambda x: x['y'])

    # Find name: Look for large text in the middle area that matches name pattern
    for i, item in enumerate(all_text_items):
        text = item['text']

        # Check if this could be a name (large font, center area, name pattern)
        if (item['font_size'] > 15 and  # Large font
            item['x'] > page_width * 0.2 and item['x'] < page_width * 0.8 and  # Center horizontally
            item['y'] > page_height * 0.3 and item['y'] < page_height * 0.6 and  # Middle vertically
            re.match(r'^[A-Za-z\s\.]+$', text) and  # Only letters, spaces, dots
            len(text) > 5 and  # Reasonable length
            not any(skip in text.lower() for skip in ['bachelor', 'science', 'hospitality', 'management', 'philippines', 'manila', 'greetings'])):
            name = text
            break

    # Find passkey: Look for alphanumeric code in bottom right area
    passkey_pattern = r'\b[A-Z0-9]{6,12}\b'

    for item in all_text_items:
        # Check if in bottom right area (where QR code typically is)
        if (item['x'] > page_width * 0.7 and  # Right side
            item['y'] > page_height * 0.8):   # Bottom area

            matches = re.findall(passkey_pattern, item['text'])
            if matches:
                passkey = matches[0]
                break

    return name, passkey

def preview_page_text(pdf_path, page_num=0):
    """
    Preview text extraction for a specific page to help with debugging
    """
    doc = fitz.open(pdf_path)
    page = doc[page_num]

    print(f"=== PAGE {page_num + 1} PREVIEW ===")
    print(f"Page dimensions: {page.rect.width} x {page.rect.height}")

    # Get text with positioning
    text_dict = page.get_text("dict")

    all_items = []
    for block in text_dict["blocks"]:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text_content = span["text"].strip()
                    if text_content:
                        bbox = span["bbox"]
                        all_items.append({
                            'text': text_content,
                            'x': bbox[0],
                            'y': bbox[1],
                            'font_size': span.get('size', 0)
                        })

    # Sort by position
    all_items.sort(key=lambda x: x['y'])

    print("\nAll text items (sorted by position):")
    for item in all_items:
        print(f"Y:{item['y']:3.0f} X:{item['x']:3.0f} Size:{item['font_size']:2.0f} | {item['text']}")

    # Test extraction
    name, passkey = extract_name_and_passkey_improved(page)
    print(f"\nExtracted Name: '{name}'")
    print(f"Extracted Passkey: '{passkey}'")

    doc.close()

def scan_pdf_batch(pdf_path, output_excel_path, start_page=0, end_page=None):
    """
    Scan PDF and extract names and passkeys from specified page range
    """
    print(f"Opening PDF: {pdf_path}")

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return None

    total_pages = len(doc)
    if end_page is None:
        end_page = total_pages

    print(f"PDF has {total_pages} pages")
    print(f"Processing pages {start_page + 1} to {min(end_page, total_pages)}")

    # Lists to store extracted data
    results = []

    # Process specified page range
    for page_num in range(start_page, min(end_page, total_pages)):
        print(f"Processing page {page_num + 1}/{total_pages}")

        page = doc[page_num]
        name, passkey = extract_name_and_passkey_improved(page)

        result = {
            'Page': page_num + 1,
            'Name': name,
            'Passkey': passkey
        }
        results.append(result)

        print(f"  Name: {name}")
        print(f"  Passkey: {passkey}")

    doc.close()

    # Create DataFrame
    df = pd.DataFrame(results)

    # Save to Excel
    print(f"\nSaving results to Excel: {output_excel_path}")
    df.to_excel(output_excel_path, index=False)

    print(f"Extraction complete! Processed {len(df)} pages.")

    # Display summary
    valid_names = df[df['Name'] != ''].shape[0]
    valid_passkeys = df[df['Passkey'] != ''].shape[0]
    print(f"Successfully extracted {valid_names} names and {valid_passkeys} passkeys")

    return df

def main():
    """
    Main function with interactive options
    """
    # Configuration - UPDATE THESE PATHS
    pdf_path = "BSHM.pdf"  
    output_excel_path = "extracted_names_passkeys.xlsx"

    print("PDF Name and Passkey Scanner")
    print("=" * 40)

    # Check if PDF file exists
    if not os.path.exists(pdf_path):
        print(f"PDF file not found: {pdf_path}")
        print("\nPlease update the 'pdf_path' variable in the script with your actual PDF file path.")
        print("Example: pdf_path = r'C:\\Users\\ksan\\Documents\\your_certificate_file.pdf'")
        return

    print("Choose an option:")
    print("1. Preview first page (for testing)")
    print("2. Process first 5 pages (small test)")
    print("3. Process all 273 pages")
    print("4. Process custom page range")

    choice = input("Enter your choice (1-4): ").strip()

    if choice == "1":
        preview_page_text(pdf_path, 0)

    elif choice == "2":
        results = scan_pdf_batch(pdf_path, "test_5_pages.xlsx", 0, 5)
        if results is not None:
            print("\nFirst 5 results:")
            print(results)

    elif choice == "3":
        results = scan_pdf_batch(pdf_path, output_excel_path, 0, 273)
        if results is not None:
            print("\nFirst 10 results:")
            print(results.head(10))

    elif choice == "4":
        try:
            start = int(input("Start page (1-based): ")) - 1
            end = int(input("End page (1-based): "))
            custom_output = f"extracted_pages_{start+1}_to_{end}.xlsx"
            results = scan_pdf_batch(pdf_path, custom_output, start, end)
            if results is not None:
                print(f"\nResults for pages {start+1} to {end}:")
                print(results)
        except ValueError:
            print("Invalid page numbers entered.")

    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()
