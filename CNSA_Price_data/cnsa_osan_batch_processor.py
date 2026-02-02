#!/usr/bin/env python3
"""
CNSA Haiti OSAN Price Bulletin Batch Downloader and Extractor
==============================================================
This script downloads all PDF price bulletins from CNSA Haiti's OSAN 
(Observatoire de la Sécurité Alimentaire du Nord) and extracts the 
food price tables into CSV format.

Requirements:
    pip install requests pdfplumber pandas

Usage:
    python cnsa_osan_batch_processor.py

The script will:
1. Download all 35 PDF bulletins (2018-2020)
2. Extract price tables from each PDF
3. Create individual CSV files for each bulletin
4. Create combined master CSV files with all data

Output directories:
    ./downloads/           - Downloaded PDF files
    ./csv_individual/      - Individual CSV files per bulletin
    ./csv_combined/        - Combined master CSV files
"""

import os
import re
import time
import requests
import pdfplumber
import pandas as pd
from urllib.parse import unquote

# All PDF URLs from https://www.cnsahaiti.org/bulletin-osan/
PDF_SOURCES = [
    # 2020
    {"year": 2020, "month": "July", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2020/Bulletin%20Prix%2C%20Juillet%202020-1_OSAN_C.pdf"},
    {"year": 2020, "month": "July", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2020/Bulletin%20Prix%2C%20Juillet%202020-2_OSAN_C.pdf"},
    {"year": 2020, "month": "June", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2020/Bulletin%20Prix%2C%20Juin%202020-1_OSAN_C.pdf"},
    {"year": 2020, "month": "June", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2020/Bulletin%20Prix%2C%20Juin%202020-2_OSAN_C.pdf"},
    # 2019
    {"year": 2019, "month": "May", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Mai_19/OSAN%20Bulletin%20Prix%20Q1%20Mai2019.pdf"},
    {"year": 2019, "month": "April", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Avril_19/OSAN%20Bulletin%20Prix%20Q1%20AVRIL%202019.pdf"},
    {"year": 2019, "month": "April", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Avril_19/OSAN%20Bulletin%20Prix%20Q2%20AVRIL%202019.pdf"},
    {"year": 2019, "month": "March", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Mars_19/OSAN%20Bulletin%20Prix%20Q1%20Mars%202019.pdf"},
    {"year": 2019, "month": "March", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Mars_19/OSAN%20Bulletin%20Prix%20Q2%20Mars%202019.pdf"},
    {"year": 2019, "month": "February", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Fevrier_19/OSAN%20Bulletin%20Prix%20Q1%2C%20Fevrier%202019.pdf"},
    {"year": 2019, "month": "February", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/Fevrier_19/OSAN%20Bulletin%20Prix%20Q2%2C%20Fevrier%202019.pdf"},
    {"year": 2019, "month": "January", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/janvier_19/OSAN%20Bulletin%20Prix%20Q1%20Janvier%202019.pdf"},
    {"year": 2019, "month": "January", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2019/janvier_19/OSAN%20Bulletin%20Prix%20Q2%20Janvier%202019.pdf"},
    # 2018
    {"year": 2018, "month": "December", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Decembre-18/OSAN%20Bulletin%20Prix%20Q1%20Decembre%202018.pdf"},
    {"year": 2018, "month": "December", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Decembre-18/OSAN%20Bulletin%20Prix%20Q2%20Decembre%202018.pdf"},
    {"year": 2018, "month": "November", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Novembre-18/OSAN%20Bulletin%20Prix%20Q1%20Novembre%20%202018.pdf"},
    {"year": 2018, "month": "November", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Novembre-18/OSAN%20Bulletin%20Prix%20Q2%20Novembre%20%202018.pdf"},
    {"year": 2018, "month": "October", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Octobre-18/OSAN%20Bulletin%20Prix%20Q1%2C%20Octobre%202018.pdf"},
    {"year": 2018, "month": "October", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Octobre-18/OSAN%20Bulletin%20Prix%20Q2%2C%20Octobre%202018.pdf"},
    {"year": 2018, "month": "September", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Septembre-18/OSASE_Bulletin%20Prix%20Q1%20Sept%202018.pdf"},
    {"year": 2018, "month": "September", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Septembre-18/OSASE_Bulletin%20Prix%20Q2%20Sept%202018.pdf"},
    {"year": 2018, "month": "August", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Aout-18/OSASE_Bulletin%20Prix%20Q1%20Aout%202018.pdf"},
    {"year": 2018, "month": "August", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Aout-18/OSASE_Bulletin%20Prix%20Q2%20Aout%202018.pdf"},
    {"year": 2018, "month": "July", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Juillet-18/OSASE_Bulletin%20Prix%20Q1%20Juillet%202018.pdf"},
    {"year": 2018, "month": "June", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Juin-18/OSAN%20Bulletin%20prix%20Q1%20Juin%20%202018.pdf"},
    {"year": 2018, "month": "June", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Juin-18/OSAN%20Bulletin%20prix%20Q2%20Juin%202018.pdf"},
    {"year": 2018, "month": "May", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Mai-18/OSAN%20Bulletin%20prix%20Q1%20Mai%202018.pdf"},
    {"year": 2018, "month": "May", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Mai-18/OSAN%20Bulletin%20prix%20Q2%20Mai%20%202018.pdf"},
    {"year": 2018, "month": "April", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Avril-18/OSAN%20Bulletin%20prix%20Q1%20Avril%20%202018.pdf"},
    {"year": 2018, "month": "April", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Avril-18/OSAN%20Bulletin%20prix%20Q2%20Avril%20%202018.pdf"},
    {"year": 2018, "month": "March", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Mars-18/OSAN_Bulletin%20Prix_Q%201%20Mars%20%202018.pdf"},
    {"year": 2018, "month": "March", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Mars-18/OSAN_Bulletin%20Prix_Q%202%20Mars%20%202018.pdf"},
    {"year": 2018, "month": "February", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Fevrier-18/OSAN_Bulletin%20Prix_Q%201%20Fevrier%20%202018.pdf"},
    {"year": 2018, "month": "February", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Fevrier-18/OSAN_Bulletin%20Prix_Q%202%20Fevrier%20%202018.pdf"},
    {"year": 2018, "month": "January", "period": "Q1", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Janvier-18/OSAN_Bulletin%20Prix_Q%201%20Janvier%20%202018.pdf"},
    {"year": 2018, "month": "January", "period": "Q2", 
     "url": "https://www.cnsahaiti.org/Web/Bulletin_Observatoires/OSAN_Prix/2018/Janvier-18/OSAN_Bulletin%20Prix_Q%202%20Janvier%20%202018.pdf"},
]

# Product translations
PRODUCT_TRANSLATIONS = {
    'Riz local': 'Local Rice', 'Riz': 'Rice',
    'Maïs moulu local': 'Local Ground Corn', 'Maïs moulu': 'Ground Corn',
    'Riz Importé': 'Imported Rice', 'Riz importé': 'Imported Rice',
    'Maïs moulu Importé': 'Imported Ground Corn', 'Maïs moulu importé': 'Imported Ground Corn',
    'Farine de blé': 'Wheat Flour', 'Farine de Blé': 'Wheat Flour',
    'Huille': 'Cooking Oil', 'Huile': 'Cooking Oil',
    'Sucre rouge': 'Brown Sugar', 'Sucre crème': 'Cream Sugar', 'Sucre blanc': 'White Sugar',
    'Petit mil': 'Millet/Sorghum', 'Petit mil ou Sorgho': 'Millet/Sorghum',
    'Haricot noir': 'Black Beans', 'Haricot rouge': 'Red Beans',
    'Haricot importé': 'Imported Beans', 'Haricot Pinto': 'Pinto Beans', 'Haricot': 'Beans',
    'Spaghetti': 'Spaghetti', 'Pistache': 'Peanuts',
    'Pistache(en gousse)': 'Peanuts (in shell)', 'Pois congo': 'Pigeon Peas',
    'Produits locaux': 'Local Products', 'Produits importés': 'Imported Products',
}

def create_directories():
    """Create output directories."""
    for d in ['downloads', 'csv_individual', 'csv_combined']:
        os.makedirs(d, exist_ok=True)

def download_pdf(source):
    """Download a PDF file."""
    filename = f"downloads/OSAN_{source['year']}_{source['month']}_{source['period']}.pdf"
    
    if os.path.exists(filename):
        print(f"  [SKIP] Already exists: {filename}")
        return filename
    
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (compatible; research bot)'}
        response = requests.get(source['url'], headers=headers, timeout=30)
        response.raise_for_status()
        
        with open(filename, 'wb') as f:
            f.write(response.content)
        
        print(f"  [OK] Downloaded: {filename}")
        return filename
    except Exception as e:
        print(f"  [ERROR] Failed to download: {e}")
        return None

def clean_cell(cell):
    """Clean a table cell value."""
    if cell is None:
        return 'ND'
    cell = str(cell).strip()
    cell = re.sub(r'\s+', ' ', cell)
    if cell.lower() == 'nd' or cell == '':
        return 'ND'
    return cell

def translate_product(text):
    """Translate product name from French to English."""
    if not text:
        return text
    text = str(text).strip()
    for fr, en in PRODUCT_TRANSLATIONS.items():
        if fr.lower() in text.lower():
            return en
    return text

def extract_tables_from_pdf(pdf_path, source):
    """Extract tables from a PDF file."""
    results = {'market_prices': None, 'price_changes': None}
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Try multiple extraction strategies
                for strategy in [
                    {'vertical_strategy': 'text', 'horizontal_strategy': 'text'},
                    {'vertical_strategy': 'lines', 'horizontal_strategy': 'lines'},
                ]:
                    tables = page.extract_tables(strategy)
                    
                    for table in tables:
                        if not table or len(table) < 3:
                            continue
                        
                        # Check table content
                        table_text = ' '.join(clean_cell(c) for row in table[:5] for c in row if c)
                        
                        # Market prices table
                        if 'Borgne' in table_text or 'Dondon' in table_text:
                            if results['market_prices'] is None:
                                df = process_market_table(table, source)
                                if df is not None:
                                    results['market_prices'] = df
                        
                        # Price changes table
                        elif 'Quinzaine' in table_text or 'variation' in table_text.lower():
                            if results['price_changes'] is None:
                                df = process_changes_table(table, source)
                                if df is not None:
                                    results['price_changes'] = df
    
    except Exception as e:
        print(f"  [ERROR] Extraction failed: {e}")
    
    return results

def process_market_table(table, source):
    """Process market prices table."""
    try:
        # Find header row
        header_idx = 0
        for i, row in enumerate(table):
            if any('Produit' in str(c) for c in row if c):
                header_idx = i
                break
        
        # Extract data
        rows = []
        for row in table[header_idx+1:]:
            cleaned = [clean_cell(c) for c in row]
            if cleaned[0] and cleaned[0] != 'ND':
                cleaned[0] = translate_product(cleaned[0])
                rows.append(cleaned)
        
        if not rows:
            return None
        
        # Create DataFrame
        columns = ['Product', 'Brand', 'Cap_Haitien', 'Borgne', 'Dondon', 
                   'Ranquitte', 'Bahon', 'Limbe', 'Price_Max', 'Price_Min', 'Price_Median']
        
        df = pd.DataFrame(rows)
        df.columns = columns[:len(df.columns)]
        df['Year'] = source['year']
        df['Month'] = source['month']
        df['Period'] = source['period']
        df['Currency'] = 'HTG'
        
        return df
    
    except Exception as e:
        return None

def process_changes_table(table, source):
    """Process price changes table."""
    try:
        rows = []
        current_category = None
        
        for row in table:
            cleaned = [clean_cell(c) for c in row]
            row_text = ' '.join(cleaned)
            
            if 'locaux' in row_text.lower():
                current_category = 'Local Products'
                continue
            elif 'importé' in row_text.lower():
                current_category = 'Imported Products'
                continue
            
            if cleaned[0] and any(c.replace('.','').replace('-','').isdigit() for c in cleaned if c):
                cleaned[0] = translate_product(cleaned[0])
                rows.append([current_category] + cleaned)
        
        if not rows:
            return None
        
        columns = ['Category', 'Product', 'Brand', 'Unit', 
                   'Price_Previous', 'Price_Current', 'Percent_Change']
        
        df = pd.DataFrame(rows)
        df.columns = columns[:len(df.columns)]
        df['Year'] = source['year']
        df['Month'] = source['month']
        df['Period'] = source['period']
        df['Currency'] = 'HTG'
        
        return df
    
    except Exception as e:
        return None

def main():
    print("=" * 70)
    print("CNSA Haiti OSAN Price Bulletin Batch Processor")
    print("=" * 70)
    print(f"\nTotal bulletins to process: {len(PDF_SOURCES)}")
    print("Date range: January 2018 - July 2020\n")
    
    create_directories()
    
    all_market_data = []
    all_changes_data = []
    
    successful = 0
    failed = 0
    
    for i, source in enumerate(PDF_SOURCES, 1):
        print(f"\n[{i}/{len(PDF_SOURCES)}] Processing {source['year']} {source['month']} {source['period']}")
        
        # Download
        pdf_path = download_pdf(source)
        
        if pdf_path:
            # Extract
            results = extract_tables_from_pdf(pdf_path, source)
            
            # Save individual files
            base_name = f"OSAN_{source['year']}_{source['month']}_{source['period']}"
            
            if results['market_prices'] is not None:
                output_file = f"csv_individual/{base_name}_market_prices.csv"
                results['market_prices'].to_csv(output_file, index=False, encoding='utf-8-sig')
                all_market_data.append(results['market_prices'])
                print(f"  [OK] Saved: {output_file}")
            
            if results['price_changes'] is not None:
                output_file = f"csv_individual/{base_name}_price_changes.csv"
                results['price_changes'].to_csv(output_file, index=False, encoding='utf-8-sig')
                all_changes_data.append(results['price_changes'])
                print(f"  [OK] Saved: {output_file}")
            
            successful += 1
        else:
            failed += 1
        
        # Be nice to the server
        time.sleep(1)
    
    # Create combined files
    print("\n" + "=" * 70)
    print("Creating combined datasets...")
    
    if all_market_data:
        combined = pd.concat(all_market_data, ignore_index=True)
        combined.to_csv('csv_combined/OSAN_all_market_prices.csv', index=False, encoding='utf-8-sig')
        print(f"[OK] csv_combined/OSAN_all_market_prices.csv ({len(combined)} rows)")
    
    if all_changes_data:
        combined = pd.concat(all_changes_data, ignore_index=True)
        combined.to_csv('csv_combined/OSAN_all_price_changes.csv', index=False, encoding='utf-8-sig')
        print(f"[OK] csv_combined/OSAN_all_price_changes.csv ({len(combined)} rows)")
    
    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print("\nOutput directories:")
    print("  ./downloads/       - PDF files")
    print("  ./csv_individual/  - Individual CSV files")
    print("  ./csv_combined/    - Combined master files")

if __name__ == '__main__':
    main()
