#!/usr/bin/env python3
"""
OpenAI PDF Table Extractor
==========================
This script uses OpenAI's GPT-4 Vision API to extract table data from PDF files
and save them as CSV files.

Requirements:
    pip install openai pdf2image pandas pillow

System requirement (for pdf2image):
    - Linux: sudo apt-get install poppler-utils
    - macOS: brew install poppler
    - Windows: Download poppler and add to PATH

Usage:
    export OPENAI_API_KEY="your-api-key"
    python openai_pdf_table_extractor.py

The script will:
1. Convert each PDF page to an image
2. Send the image to OpenAI GPT-4 Vision for table extraction
3. Parse the response and save as CSV files
"""

import os
import sys
import json
import base64
import time
from pathlib import Path
from io import BytesIO

import pandas as pd
from openai import OpenAI
from pdf2image import convert_from_path

# Configuration
DOWNLOADS_DIR = Path(__file__).parent / "downloads"
OUTPUT_DIR = Path(__file__).parent / "csv_openai"
MAX_RETRIES = 3
RETRY_DELAY = 2  # seconds


def get_openai_client():
    """Initialize OpenAI client with API key from environment."""
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("ERROR: OPENAI_API_KEY environment variable not set.")
        print("Please set it with: export OPENAI_API_KEY='your-api-key'")
        sys.exit(1)
    return OpenAI(api_key=api_key)


def image_to_base64(image):
    """Convert PIL Image to base64 string."""
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def extract_tables_from_image(client, image_base64, pdf_name, page_num):
    """
    Use OpenAI GPT-4 Vision to extract table data from an image.

    Returns a list of dictionaries, each representing a table with its data.
    """
    prompt = """Analyze this image from a PDF document. Extract ALL tables you find.
The document is in French - translate ALL text to English.

For each table found, provide the data in the following JSON format:
{
    "tables": [
        {
            "table_name": "descriptive name in English based on content",
            "headers": ["column1", "column2", ...],
            "rows": [
                ["value1", "value2", ...],
                ["value1", "value2", ...]
            ]
        }
    ]
}

Important instructions:
1. TRANSLATE all French text to English (headers, product names, categories, etc.)
2. Common French to English translations for this document:
   - Riz local → Local Rice
   - Riz importé → Imported Rice
   - Maïs moulu → Ground Corn
   - Haricot noir/rouge → Black/Red Beans
   - Huile → Cooking Oil
   - Sucre → Sugar
   - Farine de blé → Wheat Flour
   - Produits locaux → Local Products
   - Produits importés → Imported Products
   - Prix → Price
   - Variation → Change
   - Quinzaine → Fortnight/Biweekly
3. If a cell is empty or contains "ND" (No Data), use "ND" as the value
4. Preserve numeric values exactly (including decimals)
5. For merged cells spanning multiple columns, repeat the value or use empty strings
6. Include ALL rows and ALL columns from each table
7. If no tables are found, return {"tables": []}
8. Return ONLY valid JSON, no additional text

Focus on extracting:
- Market price tables (with locations like Cap Haitien, Borgne, Dondon, etc.)
- Price comparison/change tables
- Any other data tables present"""

    for attempt in range(MAX_RETRIES):
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{image_base64}",
                                    "detail": "high"
                                }
                            }
                        ]
                    }
                ],
                max_tokens=4096,
                temperature=0
            )

            content = response.choices[0].message.content

            # Try to parse JSON from the response
            # Handle cases where response might have markdown code blocks
            if "```json" in content:
                content = content.split("```json")[1].split("```")[0]
            elif "```" in content:
                content = content.split("```")[1].split("```")[0]

            result = json.loads(content.strip())
            return result.get("tables", [])

        except json.JSONDecodeError as e:
            print(f"    [WARN] JSON parse error on attempt {attempt + 1}: {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
            continue
        except Exception as e:
            print(f"    [ERROR] API error on attempt {attempt + 1}: {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
            continue

    return []


def table_to_dataframe(table_data):
    """Convert extracted table data to a pandas DataFrame."""
    if not table_data.get("headers") or not table_data.get("rows"):
        return None

    try:
        df = pd.DataFrame(table_data["rows"], columns=table_data["headers"])
        return df
    except Exception as e:
        print(f"    [WARN] Could not create DataFrame: {e}")
        return None


def process_pdf(client, pdf_path):
    """
    Process a single PDF file and extract all tables.

    Returns a list of (table_name, DataFrame) tuples.
    """
    pdf_name = pdf_path.stem
    print(f"\n  Converting PDF to images...")

    try:
        # Convert PDF pages to images
        images = convert_from_path(pdf_path, dpi=150)
        print(f"  Found {len(images)} pages")
    except Exception as e:
        print(f"  [ERROR] Could not convert PDF: {e}")
        return []

    all_tables = []

    for page_num, image in enumerate(images, 1):
        print(f"  Processing page {page_num}/{len(images)}...")

        # Convert image to base64
        image_base64 = image_to_base64(image)

        # Extract tables using OpenAI
        tables = extract_tables_from_image(client, image_base64, pdf_name, page_num)

        if tables:
            print(f"    Found {len(tables)} table(s) on page {page_num}")
            for table in tables:
                df = table_to_dataframe(table)
                if df is not None and len(df) > 0:
                    table_name = table.get("table_name", f"table_page{page_num}")
                    # Clean table name for filename
                    table_name = "".join(c if c.isalnum() or c in "_ " else "_" for c in table_name)
                    table_name = table_name.replace(" ", "_")[:50]
                    all_tables.append((table_name, df, page_num))
        else:
            print(f"    No tables found on page {page_num}")

        # Rate limiting - be nice to the API
        time.sleep(0.5)

    return all_tables


def main():
    print("=" * 70)
    print("OpenAI PDF Table Extractor")
    print("=" * 70)

    # Initialize OpenAI client
    client = get_openai_client()
    print("[OK] OpenAI client initialized")

    # Create output directory
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"[OK] Output directory: {OUTPUT_DIR}")

    # Find all PDFs
    pdf_files = sorted(DOWNLOADS_DIR.glob("*.pdf"))

    if not pdf_files:
        print(f"\n[ERROR] No PDF files found in {DOWNLOADS_DIR}")
        print("Please ensure PDF files are in the 'downloads' folder.")
        sys.exit(1)

    print(f"\nFound {len(pdf_files)} PDF files to process")

    # Process each PDF
    total_tables = 0
    successful_pdfs = 0
    failed_pdfs = 0

    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"\n{'=' * 70}")
        print(f"[{i}/{len(pdf_files)}] Processing: {pdf_path.name}")

        try:
            tables = process_pdf(client, pdf_path)

            if tables:
                # Save each table as a separate CSV
                for table_name, df, page_num in tables:
                    output_filename = f"{pdf_path.stem}_p{page_num}_{table_name}.csv"
                    output_path = OUTPUT_DIR / output_filename

                    df.to_csv(output_path, index=False, encoding="utf-8-sig")
                    print(f"  [OK] Saved: {output_filename} ({len(df)} rows)")
                    total_tables += 1

                successful_pdfs += 1
            else:
                print(f"  [WARN] No tables extracted from {pdf_path.name}")
                failed_pdfs += 1

        except Exception as e:
            print(f"  [ERROR] Failed to process {pdf_path.name}: {e}")
            failed_pdfs += 1

    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"PDFs processed successfully: {successful_pdfs}")
    print(f"PDFs with errors/no tables: {failed_pdfs}")
    print(f"Total tables extracted: {total_tables}")
    print(f"\nOutput directory: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
