import streamlit as st
import pdfplumber
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from io import BytesIO
from PIL import Image
import tempfile
import os
import camelot
from pathlib import Path
from sentence_transformers import SentenceTransformer, util
import numpy as np
import plotly.express as px
import matplotlib.pyplot as plt
import zipfile
from docx import Document
from fpdf import FPDF, HTMLMixin
from docx2pdf import convert as docx2pdf_convert

# --- Constants ---
STREAMLIT_UPLOAD_LIMIT_MB = 200
TEMP_BOQ_FOLDER = "temp_boq_files"

# --- Helper Functions ---

def convert_df_to_csv(df):
    """
    Convert DataFrame to CSV for download.
    """
    return df.to_csv(index=False).encode('utf-8')

class ReportPDF(FPDF, HTMLMixin):
    """
    An improved PDF class for generating well-formatted reports,
    including tables from DataFrames.
    """
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Converted Document', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(4)

    def add_df_to_pdf(self, df, max_rows=100):
        """
        Renders a pandas DataFrame as a properly formatted table in the PDF,
        with automatic column widths and text wrapping.
        """
        if df.empty:
            self.set_font('Arial', '', 11)
            self.multi_cell(0, 10, "No data available for this table.")
            return

        self.set_font('Arial', '', 8)
        page_width = self.w - 2 * self.l_margin
        
        df_copy = df.head(max_rows).copy()
        
        # Sanitize and prepare data
        for col in df_copy.columns:
            df_copy[col] = df_copy[col].astype(str).apply(lambda x: x.encode('latin-1', 'replace').decode('latin-1'))

        # --- Calculate Optimal Column Widths ---
        col_widths = {}
        # Give more space to description-like columns
        for col in df_copy.columns:
            if any(keyword in str(col).lower() for keyword in ['description', 'material', 'item']):
                 col_widths[col] = page_width * 0.4 # 40% of width
            elif any(keyword in str(col).lower() for keyword in ['unit', 'qty', 'quantity']):
                 col_widths[col] = page_width * 0.1 # 10% of width
            else:
                 col_widths[col] = page_width * 0.15 # 15% for others
        
        # Normalize widths to fit the page
        total_width = sum(col_widths.values())
        if total_width > 0:
            for col in col_widths:
                col_widths[col] = page_width * (col_widths[col] / total_width)
        else: # Handle case where all columns have zero width
            for col in df_copy.columns:
                col_widths[col] = page_width / len(df_copy.columns)


        # --- Render Header ---
        self.set_font('Arial', 'B', 8)
        self.set_fill_color(224, 235, 255) # Light blue
        x_start = self.get_x()
        for col in df_copy.columns:
            self.cell(col_widths[col], 7, str(col), 1, 0, 'C', fill=True)
        self.ln()

        # --- Render Data Rows ---
        self.set_font('Arial', '', 7)
        self.set_fill_color(255, 255, 255)
        fill = False
        for _, row in df_copy.iterrows():
            # Calculate max height needed for this row
            max_height = 5 # min height
            for col in df_copy.columns:
                text = str(row[col])
                lines = self.multi_cell(col_widths[col], 5, text, border=0, align='L', split_only=True)
                # Adjust height based on number of lines
                max_height = max(max_height, len(lines) * 5)
            
            # Check for page break before drawing the row
            if self.get_y() + max_height > self.page_break_trigger:
                self.add_page(orientation=self.cur_orientation)
                # Re-draw header on new page
                self.set_font('Arial', 'B', 8)
                self.set_fill_color(224, 235, 255)
                for col_header in df_copy.columns:
                    self.cell(col_widths[col_header], 7, str(col_header), 1, 0, 'C', fill=True)
                self.ln()
                self.set_font('Arial', '', 7)


            # Draw cells with calculated height
            x = x_start
            y = self.get_y()
            for col in df_copy.columns:
                self.set_xy(x, y)
                self.multi_cell(col_widths[col], max_height, str(row[col]), border=1, align='L', fill=fill)
                x += col_widths[col]
            self.ln(max_height)
            fill = not fill

        if len(df) > max_rows:
            self.set_font('Arial', 'I', 8)
            self.cell(0, 6, f"... and {len(df) - max_rows} more rows.", 0, 1)
        self.ln()


def convert_excel_or_csv_to_pdf(file_path, file_ext):
    """
    Converts an Excel or CSV file to a well-formatted PDF.
    """
    pdf_path = os.path.join(TEMP_BOQ_FOLDER, f"{os.path.splitext(os.path.basename(file_path))[0]}.pdf")
    pdf = ReportPDF()
    
    try:
        if file_ext in ['.xlsx', '.xls']:
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                if df.empty:
                    continue
                pdf.add_page(orientation='L')
                pdf.chapter_title(f"Sheet: {sheet_name}")
                pdf.add_df_to_pdf(df)
        elif file_ext == '.csv':
            df = pd.read_csv(file_path)
            if not df.empty:
                pdf.add_page(orientation='L')
                pdf.chapter_title(f"CSV Data: {os.path.basename(file_path)}")
                pdf.add_df_to_pdf(df)
        
        pdf.output(pdf_path)
        return pdf_path
    except Exception as e:
        st.error(f"Error converting {os.path.basename(file_path)} to PDF: {e}")
        return None

def process_uploaded_files(uploaded_files):
    """
    Handles uploaded files, converting them to PDF if necessary, and saves them
    to a temporary directory.
    Returns a list of paths to the processed PDF files and their total size.
    """
    if not os.path.exists(TEMP_BOQ_FOLDER):
        os.makedirs(TEMP_BOQ_FOLDER)

    processed_pdf_paths = []
    total_processed_size = 0

    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.getvalue()
        filename = uploaded_file.name
        file_ext = os.path.splitext(filename)[1].lower()
        
        temp_path = os.path.join(TEMP_BOQ_FOLDER, filename)
        with open(temp_path, "wb") as f:
            f.write(file_bytes)

        if file_ext == '.pdf':
            processed_pdf_paths.append(temp_path)
            total_processed_size += len(file_bytes)
            continue

        pdf_path = None
        try:
            if file_ext == '.zip':
                with zipfile.ZipFile(temp_path, 'r') as zip_ref:
                    for member in zip_ref.infolist():
                        if not member.is_dir() and not member.filename.startswith('__MACOSX/'):
                            extracted_path = zip_ref.extract(member, path=TEMP_BOQ_FOLDER)
                            # Recursively process extracted files
                            with open(extracted_path, 'rb') as f_bytes:
                                mock_uploaded_file = BytesIO(f_bytes.read())
                                mock_uploaded_file.name = os.path.basename(extracted_path)
                                inner_paths, inner_size = process_uploaded_files([mock_uploaded_file])
                                processed_pdf_paths.extend(inner_paths)
                                total_processed_size += inner_size
            elif file_ext == '.docx':
                pdf_path = os.path.join(TEMP_BOQ_FOLDER, f"{os.path.splitext(filename)[0]}.pdf")
                try:
                    docx2pdf_convert(temp_path, pdf_path)
                except Exception:
                    st.warning(f"High-fidelity DOCX conversion failed for '{filename}'. Falling back to basic text extraction.")
                    document = Document(temp_path)
                    pdf = ReportPDF() # Use the improved PDF class
                    pdf.add_page()
                    for para in document.paragraphs:
                        pdf.chapter_body(para.text)
                    pdf.output(pdf_path)

            elif file_ext in ['.xlsx', '.xls', '.csv']:
                pdf_path = convert_excel_or_csv_to_pdf(temp_path, file_ext)

            elif file_ext in ['.jpeg', '.jpg', '.png']:
                pdf_path = os.path.join(TEMP_BOQ_FOLDER, f"{os.path.splitext(filename)[0]}.pdf")
                pdf = ReportPDF()
                pdf.add_page()
                pdf.image(temp_path, x=10, y=20, w=pdf.w - 20)
                pdf.output(pdf_path)
            
            if pdf_path and os.path.exists(pdf_path):
                processed_pdf_paths.append(pdf_path)
                with open(pdf_path, 'rb') as f:
                    total_processed_size += len(f.read())
            
        except Exception as e:
            st.error(f"Failed to process {filename}: {e}")

    return processed_pdf_paths, total_processed_size


# ---------------------------
# --- Targeted PDF Extraction ---
# ---------------------------

# Target materials to focus on
TARGET_MATERIALS = ['concrete', 'steel', 'asphalt']

# Common variations and related terms for target materials
MATERIAL_VARIATIONS = {
    'concrete': ['concrete', 'cement', 'rcc', 'pcc', 'precast', 'cast-in-situ', 'reinforced cement'],
    'steel': ['steel', 'reinforcement', 'rebar', 'iron bar', 'metal', 'ms rod', 'structural steel'],
    'asphalt': ['asphalt', 'bituminous', 'bitumen', 'tack coat', 'prime coat', 'wearing course']
}

# Common units for these materials
MATERIAL_UNITS = {
    'concrete': ['m3', 'm³', 'cum', 'cubic meter', 'cm'],
    'steel': ['kg', 'ton', 'mt', 't', 'tonne'],
    'asphalt': ['m2', 'm²', 'sqm', 'square meter', 'sm', 'm3', 'm³']
}

def extract_country_from_pdf(pdf_path):
    """Extract the country name from the first three pages of the PDF."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:
                text = page.extract_text()
                if text:
                    match = re.search(r'(?i)Country:\s*([A-Za-z ]+)', text)
                    if match:
                        return match.group(1).strip()
    except Exception:
        return "Country Not Found"
    return "Country Not Found"

def extract_text_from_pdf(pdf_path):
    """Extract text from each page of the PDF using pdfplumber."""
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += f'\n {page_text}'
        return text

def extract_text_with_ocr(pdf_path):
    """Use OCR (pytesseract) on PDF pages converted to images."""
    try:
        images = convert_from_path(pdf_path)
        text = ''
        for image in images:
            text += pytesseract.image_to_string(image)
        return text
    except Exception as e:
        st.warning(f"OCR failed for {os.path.basename(pdf_path)}: {e}")
        return ""


def extract_tables_with_pdfplumber(pdf_path):
    """Extract table data from all pages using pdfplumber with multiple strategies."""
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # First try with default settings
            page_tables = page.extract_tables()
            
            # If no tables found or tables seem incomplete, try with different settings
            if not page_tables:
                # Try with looser table settings
                page_tables = page.extract_tables(
                    table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"}
                )
                
                # If still nothing, try line-based table finding
                if not page_tables:
                    page_tables = page.extract_tables(
                        table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"}
                    )
            
            tables.extend(page_tables)
    return tables

def extract_tables_with_camelot(pdf_path):
    """Extract table data from all pages using Camelot."""
    try:
        # Try with lattice mode (for tables with visible borders)
        tables = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
        
        # If that produces poor results, try stream mode
        if len(tables) == 0 or all(table.df.empty for table in tables):
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
            
        return [table.df for table in tables]
    except Exception as e:
        print(f"Camelot extraction error: {e}")
        return []

def parse_quantity(quantity_value):
    """
    Parse quantity values handling various formats including commas and whitespace.
    """
    if quantity_value is None:
        return None
        
    # Convert to string and clean
    qty_str = str(quantity_value).strip()
    
    # Skip empty values
    if not qty_str:
        return None
        
    # Handle numeric values already
    if isinstance(quantity_value, (int, float)):
        return float(quantity_value)
    
    # Remove all non-numeric chars except decimal points
    # First, replace commas with nothing (assuming comma as thousand separator)
    qty_str = qty_str.replace(',', '')
    
    # Then remove any other non-numeric characters
    clean_qty = re.sub(r'[^\d.]', '', qty_str)
    
    try:
        return float(clean_qty) if clean_qty else None
    except ValueError:
        return None

def is_target_material(text):
    """Check if text contains any of our target materials or variations."""
    text = text.lower()
    
    for material in TARGET_MATERIALS:
        for variation in MATERIAL_VARIATIONS[material]:
            if variation in text:
                return True, material
    return False, None

def process_table(table):
    """
    Process table data focusing on concrete, steel, and asphalt materials.
    """
    if not table or len(table) < 2:
        return None  # No valid header + data
        
    header = table[0]
    data_rows = table[1:]
    
    if not header:
        return None
        
    # Clean and sanitize header names
    sanitized_header = []
    for idx, col_name in enumerate(header):
        if col_name is None:
            col_name = f"Unnamed_{idx}"
        else:
            col_name = str(col_name).strip()
        if not col_name:
            col_name = f"Unnamed_{idx}"
        sanitized_header.append(col_name)
    
    # Handle duplicate column names
    unique_header = []
    name_count = {}
    for col_name in sanitized_header:
        if col_name in name_count:
            name_count[col_name] += 1
            unique_header.append(f"{col_name}_{name_count[col_name]}")
        else:
            name_count[col_name] = 1
            unique_header.append(col_name)
        
    # Create the DataFrame
    try:
        df = pd.DataFrame(data_rows, columns=unique_header)
        
        # Try to identify columns based on header text patterns
        renamed_cols = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['quantity', 'qty', 'in figures', 'figure']):
                renamed_cols[col] = 'Quantity'
            elif any(term in col_lower for term in ['unit', 'uom', 'measure']):
                renamed_cols[col] = 'Unit'
            elif any(term in col_lower for term in ['desc', 'material', 'item name', 'particular', 'equipment']):
                renamed_cols[col] = 'Material'
            elif any(term in col_lower for term in ['item', 'no.', 'code', '#', 'sl']):
                renamed_cols[col] = 'Item'
                
        # Rename the columns
        df = df.rename(columns=renamed_cols)
        
        # Filter rows where Material, Quantity, and Unit are not missing or empty
        valid_rows = []
        for idx, row in df.iterrows():
            material = row.get("Material")
            quantity = row.get("Quantity")
            unit = row.get("Unit")
            
            # Check if all required fields are present and non-empty
            if (material is not None and str(material).strip() != "" and
                quantity is not None and str(quantity).strip() != "" and
                unit is not None and str(unit).strip() != ""):
                valid_rows.append(idx)
        
        if valid_rows:
            filtered_df = df.loc[valid_rows].copy()
            
            # Convert quantity values to numeric, handling commas and other formatting
            if 'Quantity' in filtered_df.columns:
                filtered_df['Quantity'] = filtered_df['Quantity'].apply(parse_quantity)
            
            # Filter for rows containing our target materials
            if 'Material' in filtered_df.columns:
                material_rows = []
                material_types = []
                
                for idx, row in filtered_df.iterrows():
                    material_desc = str(row['Material']).lower() if row['Material'] is not None else ""
                    is_target, material_type = is_target_material(material_desc)
                    
                    if is_target:
                        material_rows.append(idx)
                        material_types.append(material_type)
                
                if material_rows:
                    filtered_df = filtered_df.loc[material_rows].copy()
                    filtered_df['Material_Type'] = material_types
                    return filtered_df
                
        return None
    except Exception as e:
        print(f"Error processing table: {e}")
        return None

def extract_material_rows_from_text(text):
    """
    Extract material information from unstructured text,
    focusing on concrete, steel, and asphalt.
    """
    lines = text.split('\n')
    extracted_items = []
    
    # Pattern to match quantities (numbers with potential commas and decimals)
    qty_pattern = r'\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b'
    
    # Pattern to match potential units
    all_units = []
    for material in TARGET_MATERIALS:
        all_units.extend(MATERIAL_UNITS[material])
    unit_pattern = r'\b(?:' + '|'.join(all_units) + r')\b'
    
    # Process each line for material mentions
    for line in lines:
        # Skip lines that are too short
        if len(line.strip()) < 5:
            continue
            
        is_target, material_type = is_target_material(line)
        
        if is_target:
            # Find quantity
            qty_matches = re.findall(qty_pattern, line)
            quantity = None
            if qty_matches:
                # Take the first numeric match and parse it
                quantity = parse_quantity(qty_matches[0])
            
            # Find unit
            unit_matches = re.findall(unit_pattern, line, re.IGNORECASE)
            unit = unit_matches[0] if unit_matches else None
            
            # Only add if we found both a quantity and unit
            if quantity is not None and unit is not None:
                extracted_items.append({
                    'material': line.strip(),
                    'quantity': quantity,
                    'unit': unit,
                    'material_type': material_type
                })
    
    return pd.DataFrame(extracted_items) if extracted_items else pd.DataFrame()

def filter_boq_data(df):
    """
    Filter DataFrame to keep only rows with target materials.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    
    # If we already have material_type column, we're good
    if 'material_type' in df.columns:
        return df
    
    # Otherwise, filter based on material column
    if 'material' in df.columns:
        material_rows = []
        material_types = []
        
        for idx, row in df.iterrows():
            material_desc = str(row['material']).lower() if row['material'] is not None else ""
            is_target, material_type = is_target_material(material_desc)
            
            if is_target:
                material_rows.append(idx)
                material_types.append(material_type)
        
        if material_rows:
            filtered_df = df.loc[material_rows].copy()
            filtered_df['material_type'] = material_types
            return filtered_df
    
    return pd.DataFrame()

def extract_boq_from_pdf(pdf_path, original_filename):
    """
    Extract BOQ data focusing on concrete, steel, and asphalt materials.
    Tries multiple extraction methods to ensure best coverage.
    """
    extracted_data_frames = []
    
    # Try table extraction with pdfplumber first
    tables = extract_tables_with_pdfplumber(pdf_path)
    
    for table in tables:
        if not table:
            continue
            
        processed_df = process_table(table)
        if processed_df is not None and not processed_df.empty:
            extracted_data_frames.append(processed_df)
    
    # If no tables found or extraction yielded poor results, try camelot
    if not extracted_data_frames:
        tables = extract_tables_with_camelot(pdf_path)
        
        for table in tables:
            if isinstance(table, pd.DataFrame):
                table_list = [table.columns.to_list()] + table.values.tolist()
                processed_df = process_table(table_list)
            else:
                processed_df = process_table(table)
                
            if processed_df is not None and not processed_df.empty:
                extracted_data_frames.append(processed_df)
    
    # If still no results, try text-based extraction
    if not extracted_data_frames:
        text = extract_text_from_pdf(pdf_path)
        text_df = extract_material_rows_from_text(text)
        
        if not text_df.empty:
            extracted_data_frames.append(text_df)
    
    # Combine all results
    if extracted_data_frames:
        combined_df = pd.concat(extracted_data_frames, ignore_index=True)
        # Add the source PDF name
        combined_df["Source PDF"] = original_filename
        return combined_df
    
    return pd.DataFrame()

def clean_boq_df(df):
    """
    Standardize column names and ensure required columns exist.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    
    # Standardize column names
    col_mapping = {
        'material': ['material', 'description', 'desc', 'item name', 'particular'],
        'quantity': ['quantity', 'qty', 'in figures', 'figure'],
        'unit': ['unit', 'uom', 'measure'],
        'item': ['item', 'item #', 's/no', 'no.', 'sl'],
        'material_type': ['material_type', 'type']
    }
    
    renamed_columns = {}
    for std_col, variations in col_mapping.items():
        for col in df.columns:
            if col in renamed_columns.values():
                continue
                
            col_lower = str(col).lower().strip()
            if col_lower in variations or any(var in col_lower for var in variations):
                renamed_columns[col] = std_col
                break
    
    # Apply renaming if any matches found
    if renamed_columns:
        df = df.rename(columns=renamed_columns)
    
    # Ensure material_type column exists
    if 'material_type' not in df.columns and 'material' in df.columns:
        material_types = []
        for idx, row in df.iterrows():
            material_desc = str(row['material']).lower() if row['material'] is not None else ""
            is_target, material_type = is_target_material(material_desc)
            material_types.append(material_type if is_target else None)
        
        df['material_type'] = material_types
    
    # Filter out rows where material_type is None
    if 'material_type' in df.columns:
        df = df[df['material_type'].notna()]
    
    # Ensure quantity is numeric
    if 'quantity' in df.columns:
        df['quantity'] = df['quantity'].apply(parse_quantity)
    
    return df

def extract_boq_from_processed_pdfs(pdf_paths):
    """
    Main function to extract BOQ data from a list of processed PDF paths.
    Focuses on concrete, steel, and asphalt materials.
    """
    all_extracted_data = []
    all_countries = set()

    for pdf_path in pdf_paths:
        original_filename = os.path.basename(pdf_path)
        
        # Extract BOQ data focusing on target materials
        extracted_df = extract_boq_from_pdf(pdf_path, original_filename)
        extracted_df = clean_boq_df(extracted_df)
        
        # Get the country from the PDF
        country_name = extract_country_from_pdf(pdf_path)
        if country_name != "Country Not Found":
            all_countries.add(country_name)
        
        # Standardize column names for downstream processing
        for old_col, new_col in [('material', 'Material'), ('quantity', 'Quantity'), 
                                 ('unit', 'Unit'), ('item', 'Item'), 
                                 ('material_type', 'Material_Type')]:
            if old_col in extracted_df.columns and new_col not in extracted_df.columns:
                extracted_df = extracted_df.rename(columns={old_col: new_col})
        
        all_extracted_data.append(extracted_df)

    if not all_extracted_data:
        return pd.DataFrame(), set()

    combined_df = pd.concat(all_extracted_data, ignore_index=True)
    return combined_df, all_countries


# ---------------------------
# --- Emission Factors & GHG Calculation ---
# ---------------------------

# Country-to-region mapping
country_to_region = {
    "Kazakhstan": "Asia", "Kyrgyzstan": "Asia", "Tajikistan": "Asia", "Turkmenistan": "Asia",
    "Uzbekistan": "Asia", "China": "Asia", "Democratic People's Republic of Korea": "Asia",
    "Japan": "Asia", "Mongolia": "Asia", "Republic of Korea": "Asia", "Brunei Darussalam": "Asia",
    "Cambodia": "Asia", "Indonesia": "Asia", "Lao People's Democratic Republic": "Asia", 
    "Malaysia": "Asia", "Myanmar": "Asia", "Philippines": "Asia", "Singapore": "Asia",
    "Thailand": "Asia", "Timor-Leste": "Asia", "Viet Nam": "Asia", "Afghanistan": "Asia",
    "Bangladesh": "Asia", "Bhutan": "Asia", "India": "Asia", "Iran": "Asia", "Maldives": "Asia",
    "Nepal": "Asia", "Pakistan": "Asia", "Sri Lanka": "Asia", "Australia": "Oceania",
    "Christmas Island": "Oceania", "Cocos (Keeling) Islands": "Oceania", "Heard Island and McDonald Islands": "Oceania",
    "New Zealand": "Oceania", "Norfolk Island": "Oceania", "Fiji": "Oceania", "New Caledonia": "Oceania",
    "Papua New Guinea": "Oceania", "Solomon Islands": "Oceania", "Vanuatu": "Oceania", "Guam": "Oceania",
    "Kiribati": "Oceania", "Marshall Islands": "Oceania", "Micronesia": "Oceania", "Nauru": "Oceania",
    "Northern Mariana Islands": "Oceania", "Palau": "Oceania", "United States Minor Outlying Islands": "Oceania",
    "American Samoa": "Oceania", "Cook Islands": "Oceania", "French Polynesia": "Oceania", "Niue": "Oceania",
    "Pitcairn": "Oceania", "Samoa": "Oceania", "Tokelau": "Oceania", "Tonga": "Oceania", "Tuvalu": "Oceania",
    "Wallis and Futuna Islands": "Oceania"
}

def extract_country_from_text(text, country_to_region):
    """
    Extracts a country name from the provided text by comparing against the mapping.
    """
    for country in country_to_region.keys():
        if country.lower() in text.lower():
            return country
    return None

def extract_numeric_value(value):
    if pd.isna(value):
        return None
    match = re.match(r'(\d+\.?\d*)', str(value))
    if match:
        return float(match.group(1))
    return None

def extract_uncertainty(gwp_str):
    """
    Extract uncertainty percentage from a GWP string.
    """
    if not isinstance(gwp_str, str):  # Ensure input is a string
        gwp_str = str(gwp_str)
    match = re.search(r'±\s*(\d+\.?\d*)%', gwp_str)
    if match:
        return float(match.group(1))
    return None

def convert_gwp_to_kg(gwp_str, declared_unit, mass_per_m3=None):
    gwp_match = re.match(r'(\d+\.?\d*)\s*kgCO2e', str(gwp_str))
    if not gwp_match:
        raise ValueError(f"Invalid GWP format: {gwp_str}")
    gwp = float(gwp_match.group(1))
    if declared_unit.lower() in ['1 m3', 'm3']:
        if not mass_per_m3:
            raise ValueError("Mass per 1 m3 is required for m3 units.")
        return gwp / mass_per_m3
    elif declared_unit.lower() in ['1 t', '1 ton', '1 tonne', '1 metric ton', '1 tons']:
        return gwp / 1000
    elif declared_unit.lower().endswith('kg'):
        quantity = float(re.match(r'(\d+\.?\d*)\s*kg', declared_unit, re.IGNORECASE).group(1))
        return gwp / quantity
    else:
        raise ValueError(f"Unsupported declared unit: {declared_unit}")

def convert_boq_to_kg(quantity, unit, material_type=None, density=None):
    """
    Convert BOQ quantities from various units to kg.
    Focuses on concrete, steel, and asphalt materials.
    """
    unit = unit.strip().upper() if unit else ""
    
    # Default densities for target materials (kg/m³)
    densities = {
        'concrete': 2400,  # Reinforced concrete
        'steel': 7850,     # Steel
        'asphalt': 2300    # Asphalt/bituminous material
    }
    
    # If density is provided, use it; otherwise use material-specific density
    if density is None:
        if material_type and material_type.lower() in densities:
            density = densities[material_type.lower()]
        else:
            # Default density if material type unknown
            density = 2000  # Generic construction material density
    
    # Convert based on unit
    if unit in ['T', 'TON', 'TONNE', 'MT']:
        return quantity * 1000  # 1 ton = 1000 kg
    elif unit in ['KG']:
        return quantity  # Already in kg
    elif unit in ['CM', 'M3', 'CUM', 'M³', 'CU M']:
        return quantity * density  # volume * density
    elif unit in ['M2', 'SM', 'SQM', 'M²', 'SQ M']:
        # For area, we need thickness assumption - depends on material
        if material_type == 'asphalt':
            thickness = 0.05  # 5cm typical for asphalt
        elif material_type == 'concrete':
            thickness = 0.15  # 15cm typical for concrete slabs
        else:
            thickness = 0.1   # 10cm general assumption
        return quantity * thickness * density
    elif unit in ['M', 'LM', 'ML']:
        # For linear measurements (mainly for steel rebar)
        if material_type == 'steel':
            # Assuming 16mm diameter rebar
            diameter = 0.016  # meters
            cross_section = 3.14159 * (diameter/2)**2
            return quantity * cross_section * density
        else:
            # General assumption for other materials
            cross_section = 0.1 * 0.1  # 10cm x 10cm
            return quantity * cross_section * density
    else:
        # For unrecognized units, return None
        return None

def process_emission_factors(df, country_name, country_to_region):
    """
    Processes the emission factors DataFrame and selects the appropriate GWP column.
    Focuses on concrete, steel, and asphalt materials.
    """
    required_columns = ['Category', 'Declared Unit', 'Mass per 1 m3', 'Average GWP_Global', 
                        'Average GWP_Asia', 'Average GWP_Oceania', 'Density']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing required columns in emission factors file: {missing_columns}")
        return None, None

    # Use the full dataset without filtering
    filtered_df = df.copy()

    # Select region-specific GWP column
    region = country_to_region.get(country_name, "Global")
    if region == "Asia":
        gwp_column = 'Average GWP_Asia'
    elif region == "Oceania":
        gwp_column = 'Average GWP_Oceania'
    else:
        gwp_column = 'Average GWP_Global'

    # Handle null values in the selected GWP column
    if filtered_df[gwp_column].isnull().any():
        st.warning(f"Some values in {gwp_column} are null. Falling back to Average GWP_Global.")
        gwp_column = 'Average GWP_Global'

    # Convert numeric columns
    filtered_df['Mass per 1 m3'] = filtered_df['Mass per 1 m3'].apply(lambda x: extract_numeric_value(x))
    filtered_df['Density'] = filtered_df['Density'].apply(lambda x: extract_numeric_value(x))

    # Extract GWP uncertainty
    filtered_df['GWP Uncertainty'] = filtered_df[gwp_column].apply(lambda x: extract_uncertainty(x))

    # Convert GWP to kgCO2e/kg
    try:
        filtered_df['GWP (kgCO2e/kg)'] = filtered_df.apply(
            lambda row: convert_gwp_to_kg(row[gwp_column], row['Declared Unit'], row['Mass per 1 m3']), axis=1
        )
    except Exception as e:
        st.error(f"Error converting GWP: {e}")
        return None, None

    # Rename columns for consistency
    filtered_df = filtered_df.rename(columns={'Category': 'EC3 Material Category'})

    return filtered_df, region

# ---------------------------
# --- Matching & Emissions Calculation ---
# ---------------------------

# Initialize SentenceTransformer model for similarity matching
model = SentenceTransformer('all-MiniLM-L6-v2')
pd.set_option('display.float_format', '{:.2f}'.format)

# Updated exact matches focused on concrete, steel, and asphalt
exact_matches = {
    "Steel Reinforcement / Iron Bar (1/2‚Äù Round)": "Reinforcing Bar",
    "Steel Reinforcement": "Reinforcing Bar",
    "M. Steel Bar (G-40 & 60)": "Reinforcing Bar",
    "Reinforcement as per AASHTO M-31 Grade 60": "Reinforcing Bar",
    "Reinforcement (Structural Shapes) as per ASTM-A-36": "Reinforcing Bar",
    "50mm - Expansion Joint - Indigenous Type (Steel Plates)": "Plate Steel",
    "Steel bar D12mm (ASTM)": "Reinforcing Bar",
    "Stainless Steel Tank": "Steel Suspension Assemblies",
    "Recovered steel from existing deck slabs": "Merchant Bar (MBQ)",
    "Maintenance of steel and RCC Railing": "Open Web Steel Joists",
    "Metal Beam Crash Barrier": "Composite and Form Decks",
    "Tubular Steel Railing": "Cold Formed Framing",
    "Supply & erection of MS Galvanized octagonal pole": "Hot-Rolled Sections",
    "Asphaltic base course plant mix (Class A)": "Asphalt",
    "Asphaltic Concrete for wearing course (Class A)": "Asphalt",
    "Cut-back asphalt for bituminous prime coat": "Asphalt",
    "Cut-back asphalt for bituminous tack coat": "Asphalt",
    "Scarifying existing bituminous surface": "Asphalt",
    "Dense Graded Bituminous Macadam (DGBM)": "Asphalt",
    "Bituminous Concrete": "Asphalt",
    "Concrete Class A1 (Elevated)": "Ready Mix",
    "Concrete Class A1 (On ground)": "Ready Mix",
    "Concrete Class A1 (Onground)": "Ready Mix",
    "Concrete Class A3 (Elevated)": "Ready Mix",
    "Concrete Class A3 (On ground)": "Ready Mix",
    "Concrete Class A3 (Underground)": "Civil Precast Concrete",
    "Concrete Class B": "Ready Mix",
    "Lean Concrete": "Flowable Concrete Fill",
    "Precast Concrete Class D2 425kg/Sq.m (6,000 psi)": "Structural Precast Concrete",
    "Precast Concrete Class D2 425kg/Sq.m (6,000 psi) (Inlcuding Additional Admixtures super plasticizer Sikament 520 (ABS) or equivalent 1.25% by weight of Cement w\c ratio must not exceed 0.32-0.35 and slump should not be more than 160 mm the cost which is deemed to be included in the cost of Concrete)": "Structural Precast Concrete",
    "Plum Concrete (Cyclopean/Rubble)": "Civil Precast Concrete",
    "Reinforced Concrete grade 25 Mpa for bottom slab": "Structural Precast Concrete",
    "Concrete ring for slurry pit √ò100cm": "Utility & Underground Precast Concrete",
    "Concrete ring for mixing tank √ò60cmx0.5m": "Utility & Underground Precast Concrete",
    "Reinforced Cement Concrete Crash Barrier": "Structural Precast Concrete",
    "Cast in Situ Cement Concrete M 20 Kerb": "Structural Precast Concrete",
    "Providing and Laying Reinforced Cement Concrete Pipe NP3": "Utility & Underground Precast Concrete",
    "Rust Removal of Exposed Corroded Rebars (Chemrite Descaler A-28 or Equivalent)": "Metal Surface Treatments",
    "Sealing of Exposed Concrete Gaps with Epoxy Mortar (Chmedur 31 or Equivalent)": "Concrete Additives",
    "SBR Latex modified concrete": "Concrete Additives",
    "Cost difference between SR cement and OPC": "Portland Cement",
    "Additional Admixtures Chermite 520 (ABS) or equivalent": "Concrete Additives",
    "Additional Admixtures Silica Fumes 6% by weight of cement": "Concrete Additives",
    "Plum Concrete (Cyclopean/Rubble) (2:1 concrete stone Ratio)": "Civil Precast Concrete",
    "Concrete Class A3 Onground (from concrete mix plant)": "Ready Mix",
    "Concrete Class A3 Elevated (from concrete mix plant)": "Ready Mix",
    "40% cost of salvaged steel from the deck slab, Barrier and Approach Slab": "Merchant Bar (MBQ)",
    "Dismantling and disposal of structures and obstruction": "Not Applicable",
    "50mm - Expansion Joint - Indigonous Type 15MM thick two vertical steel plates welded 10MM thick round steel bars local (Pakistan Make),PSQCA Certified (As Specified in Drawings)": "Plate Steel"
}

# Updated normalization function to force a string
def normalize_string(s):
    s = str(s)
    s = s.lower()
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

exact_matches_normalized = {normalize_string(k): v for k, v in exact_matches.items()}

def match_materials_with_categories(boq_data, emission_factors):
    """
    Match BOQ materials with EC3 emission factor categories.
    Focuses on concrete, steel, and asphalt materials.
    """
    matched_results = []
    
    for material_data in boq_data:
        # Ensure the material value is a string
        material = material_data.get("Material")
        material_str = str(material) if material is not None else ""
        quantity = material_data.get("Quantity")
        unit = material_data.get("Unit")
        source_pdf = material_data.get("Source PDF")
        material_type = material_data.get("Material_Type")  # Use our detected material type

        # Try exact match first
        material_normalized = normalize_string(material_str)
        if material_normalized in exact_matches_normalized:
            category = exact_matches_normalized[material_normalized]
            similarity = 1.0
        else:
            # If no exact match, use the detected material type to guide category selection
            # First filter EC3 categories by material type
            if material_type:
                # Create material type specific patterns
                if material_type.lower() == 'concrete':
                    material_filter = 'concrete|cement|precast|ready mix'
                elif material_type.lower() == 'steel':
                    material_filter = 'steel|reinforc|bar|metal'
                elif material_type.lower() == 'asphalt':
                    material_filter = 'asphalt|bitum'
                else:
                    material_filter = None
                
                # Filter categories if we have a material filter
                if material_filter:
                    filtered_categories = emission_factors[
                        emission_factors["EC3 Material Category"].str.contains(material_filter, case=False, regex=True)
                    ]
                    
                    # If filtering yields results, use these categories for similarity matching
                    if not filtered_categories.empty:
                        material_embedding = model.encode(material_str, convert_to_tensor=True)
                        category_embeddings = model.encode(filtered_categories["EC3 Material Category"].tolist(), convert_to_tensor=True)
                        similarities = util.cos_sim(material_embedding, category_embeddings).cpu().numpy()
                        best_match_index = np.argmax(similarities)
                        category = filtered_categories.iloc[best_match_index]["EC3 Material Category"]
                        similarity = similarities[0][best_match_index].item()
                    else:
                        # Fallback to all categories if filtering yields no results
                        material_embedding = model.encode(material_str, convert_to_tensor=True)
                        category_embeddings = model.encode(emission_factors["EC3 Material Category"].tolist(), convert_to_tensor=True)
                        similarities = util.cos_sim(material_embedding, category_embeddings).cpu().numpy()
                        best_match_index = np.argmax(similarities)
                        category = emission_factors.iloc[best_match_index]["EC3 Material Category"]
                        similarity = similarities[0][best_match_index].item()
                else:
                    # No material filter, use full list
                    material_embedding = model.encode(material_str, convert_to_tensor=True)
                    category_embeddings = model.encode(emission_factors["EC3 Material Category"].tolist(), convert_to_tensor=True)
                    similarities = util.cos_sim(material_embedding, category_embeddings).cpu().numpy()
                    best_match_index = np.argmax(similarities)
                    category = emission_factors.iloc[best_match_index]["EC3 Material Category"]
                    similarity = similarities[0][best_match_index].item()
            else:
                # No material type detected, use full list
                material_embedding = model.encode(material_str, convert_to_tensor=True)
                category_embeddings = model.encode(emission_factors["EC3 Material Category"].tolist(), convert_to_tensor=True)
                similarities = util.cos_sim(material_embedding, category_embeddings).cpu().numpy()
                best_match_index = np.argmax(similarities)
                category = emission_factors.iloc[best_match_index]["EC3 Material Category"]
                similarity = similarities[0][best_match_index].item()
        
        matched_results.append({
            "Material": material_str,
            "Category": category,
            "Similarity": similarity,
            "Quantity": quantity,
            "Unit": unit,
            "Source PDF": source_pdf,
            "Material_Type": material_type
        })
    
    return matched_results

def extract_numeric_gwp(gwp_str):
    match = re.match(r'(\d+\.?\d*)', str(gwp_str))
    if match:
        return float(match.group(1))
    raise ValueError(f"Invalid GWP format: {gwp_str}")

def calculate_ghg_emissions(matched_results, emission_factors, region):
    """
    Calculate GHG emissions based on matched materials and quantities.
    """
    total_emissions = 0
    emissions_data = []
    
    for result in matched_results:
        # Find the emission factor for this category
        emission_rows = emission_factors[emission_factors["EC3 Material Category"] == result["Category"]]
        
        if emission_rows.empty:
            continue  # Skip if no matching emission factor found
            
        emission_row = emission_rows.iloc[0]
        emissions = None
        unsupported_unit_note = None
        
        # Check if the unit is missing or empty
        if pd.isna(result["Unit"]) or result["Unit"] == "":
            unsupported_unit_note = "No unit of measurement found"
        else:
            # Determine which regional GWP value to use
            if region == "Asia" and not pd.isna(emission_row["Average GWP_Asia"]):
                gwp_column = "Average GWP_Asia"
                gwp_region = "Asia"
            elif region == "Oceania" and not pd.isna(emission_row["Average GWP_Oceania"]):
                gwp_column = "Average GWP_Oceania"
                gwp_region = "Oceania"
            else:
                gwp_column = "Average GWP_Global"
                gwp_region = "Global"
            
            try:
                # Convert quantity to kg using material-specific conversion
                quantity_kg = convert_boq_to_kg(
                    result["Quantity"], 
                    result["Unit"], 
                    material_type=result.get("Material_Type"),
                    density=emission_row.get("Density")
                )
                
                if quantity_kg is not None:
                    gwp_value = extract_numeric_gwp(emission_row[gwp_column])
                    emissions = round(quantity_kg * gwp_value, 2)
                    total_emissions += emissions
                else:
                    unsupported_unit_note = f"Unsupported BoQ unit: {result['Unit']}"
            except ValueError as e:
                unsupported_unit_note = f"Unsupported BoQ unit: {result['Unit']}"
            except Exception as e:
                unsupported_unit_note = f"Error: {str(e)}"  # Convert exception to string
        
        # Extract uncertainty information
        gwp_uncertainty = extract_uncertainty(str(emission_row[gwp_column]))  # Ensure value is a string
        if gwp_uncertainty is not None:
            gwp_uncertainty = f"{gwp_uncertainty:.1f}%"
        
        # Format similarity as percentage
        similarity_percentage = f"{result['Similarity'] * 100:.2f}%"
        
        # Add to emissions data list
        emissions_data.append({
            "BoQ Material": result["Material"],
            "EC3 Category": result["Category"],
            "BoQ Quantity (KG)": round(quantity_kg, 2) if quantity_kg is not None else None,
            "Calculated GHG Emissions (kg CO2e)": round(emissions, 2) if emissions is not None else None,
            "EC3 Regional Average GWP (kgCO2e/kg)": gwp_value if quantity_kg is not None else None,
            "Region": gwp_region,
            "Declared Unit": emission_row["Declared Unit"],
            "Similarity": similarity_percentage,
            "GWP Uncertainty": gwp_uncertainty,
            "Source PDF": result["Source PDF"],
            "Material_Type": result.get("Material_Type"),
            "Unsupported Unit": unsupported_unit_note if unsupported_unit_note else "",  # Use empty string instead of None
            "Assumed Thickness (m)": 0.05 if result["Unit"].upper() in ['M2', 'SM', 'SQM', 'M²', 'SQ M'] else None,  # 5cm for area-based materials
            "Assumed Diameter (m)": 0.016 if result["Unit"].upper() in ['M', 'LM', 'ML'] else None  # 16mm for length-based materials
        })
    
    total_emissions = round(total_emissions, 2)
    return emissions_data, total_emissions

# ---------------------------
# --- Dashboard & Interactive Features ---
# ---------------------------

def highlight_high_uncertainty_rows(row):
    gwp_uncertainty = row["GWP Uncertainty"]
    if gwp_uncertainty and isinstance(gwp_uncertainty, str) and gwp_uncertainty.endswith('%'):
        uncertainty_value = float(gwp_uncertainty[:-1])
        if abs(uncertainty_value) > 5:
            return ['background-color: #ffcccc'] * len(row)
    return [''] * len(row)

def color_category_based_on_similarity(row):
    """
    Apply color coding to the EC3 Category column based on similarity.
    """
    similarity = float(row["Similarity"].strip('%')) / 100
    if similarity >= 0.95:
        return [''] + ['color: darkgreen'] + [''] * (len(row) - 2)
    else:
        return [''] + ['color: darkgoldenrod'] + [''] * (len(row) - 2)

def display_dashboard(emissions_df, all_countries):
    """
    Display the emissions dashboard with summary statistics and visualizations.
    """
    # Ensure numeric values are treated correctly
    emissions_df["Calculated GHG Emissions (kg CO2e)"] = pd.to_numeric(
        emissions_df["Calculated GHG Emissions (kg CO2e)"], errors="coerce"
    ).fillna(0)
    
    total_emissions = emissions_df["Calculated GHG Emissions (kg CO2e)"].sum()
    total_emissions_formatted = f"{total_emissions:,.2f}"
    
    # Convert uncertainty values to numeric
    emissions_df["GWP Uncertainty"] = emissions_df["GWP Uncertainty"].apply(
        lambda x: float(str(x).strip('%')) / 100 if isinstance(x, str) and x.endswith('%') else 0
    )
    
    total_uncertainty = (emissions_df["Calculated GHG Emissions (kg CO2e)"] * emissions_df["GWP Uncertainty"]).sum()
    total_uncertainty_percentage = (total_uncertainty / total_emissions) * 100 if total_emissions != 0 else 0
    
    # Create dashboard styling
    st.markdown(
        f"""
        <style>
        .dashboard-title {{
            background-color: #f5f5f5;
            padding: 10px;
            border-radius: 10px;
            color: rgb(0, 37, 105);
            font-size: 28px;
            text-align: center;
            margin-bottom: 20px;
        }}
        .dashboard-label {{
            color: rgb(0, 37, 105);
            font-size: 20px;
            font-weight: bold;
            margin-bottom: 0;
        }}
        .dashboard-value {{
            color: black;
            font-size: 28px;  /* Large font for Total GHG Emissions */
            font-weight: bold;  /* Bold for emphasis */
            margin-top: 0;
        }}
        .dashboard-uncertainty {{
            color: black;
            font-size: 14px;  /* Smaller font for Total Uncertainty */
            font-weight: normal;  /* Normal weight for less emphasis */
            margin-top: 10px;  /* Add some spacing */
        }}
        .dashboard-countries {{
            color: rgb(0, 37, 105);  /* Same color as Total GHG Emissions */
            font-size: 20px;  /* Same font size as Total GHG Emissions header */
            font-weight: bold;  /* Bold for emphasis */
            margin-bottom: 20px;  /* Add some spacing */
        }}
        .dashboard-block {{
            margin-bottom: 60px;
        }}
        </style>
        <div class="dashboard-title">Summary Dashboard</div>
        """,
        unsafe_allow_html=True
    )
    
    # Display detected countries below the Summary Dashboard header
    if all_countries:
        st.markdown(
            f"""
            <p class="dashboard-countries">Detected Countries: {', '.join(all_countries)}</p>
            """,
            unsafe_allow_html=True
        )
    
    # Display summary statistics and material breakdown table side by side
    col1, col2 = st.columns([1, 1.5])
    with col1:
        st.markdown(
            f"""
            <div class="dashboard-block">
                <p class="dashboard-label">Total GHG Emissions</p>
                <p class="dashboard-value">{total_emissions_formatted} kg CO2e</p>
                <p class="dashboard-uncertainty">The total uncertainty is {total_uncertainty:,.2f} kg CO2e, or {total_uncertainty_percentage:.2f}% of total calculated GHG emissions based on EC3's GWP uncertainty.</p>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    with col2:
        # Add "Material Breakdown" header
        st.markdown(
            f"""
            <p class="dashboard-label">Material Breakdown</p>
            """,
            unsafe_allow_html=True
        )
        
        # Group emissions by material type
        if "Material_Type" in emissions_df.columns:
            # Use the correct column name for quantity
            quantity_column = "BoQ Quantity (KG)"
            ghg_column = "Calculated GHG Emissions (kg CO2e)"
            
            material_emissions = emissions_df.groupby("Material_Type").agg(
                Quantity_kg=(quantity_column, "sum"),  # Sum of quantities in kg
                GHG_Emissions_kg=(ghg_column, "sum")  # Sum of GHG emissions
            ).reset_index()
            
            # Filter for concrete, steel, and asphalt
            material_emissions = material_emissions[
                material_emissions["Material_Type"].isin(["concrete", "steel", "asphalt"])
            ]
            
            # Rename columns for display
            material_emissions = material_emissions.rename(columns={
                "Material_Type": "Material Type",
                "Quantity_kg": "Quantity (kg)",
                "GHG_Emissions_kg": "GHG Emissions (kg CO2e)",
         
            })
            
            # Format the numbers: round to 2 decimal places and add thousand separators
            material_emissions["Quantity (kg)"] = material_emissions["Quantity (kg)"].apply(
                lambda x: f"{x:,.2f}" if pd.notnull(x) else ""
            )
            material_emissions["GHG Emissions (kg CO2e)"] = material_emissions["GHG Emissions (kg CO2e)"].apply(
                lambda x: f"{x:,.2f}" if pd.notnull(x) else ""
            )
            
            # Display the data table
            st.dataframe(
                material_emissions,
                column_config={
                    "Material Type": "Material Type",
                    "Quantity (kg)": st.column_config.TextColumn(
                        "Quantity (kg)",
                        help="Total quantity in kilograms"
                    ),
                    "GHG Emissions (kg CO2e)": st.column_config.TextColumn(
                        "GHG Emissions (kg CO2e)",
                        help="Total GHG emissions in kilograms of CO2 equivalent"
                    )
                },
                use_container_width=True
            )
    
    # Display pie chart below the Total GHG Emissions and Material Breakdown table
    # Prepare data for similarity pie chart
    high_similarity_count = emissions_df[emissions_df["Similarity"].apply(lambda x: float(str(x).strip('%'))) >= 95].shape[0]
    low_similarity_count = emissions_df[emissions_df["Similarity"].apply(lambda x: float(str(x).strip('%'))) < 95].shape[0]
    
    pie_data = {
        "Category": ["≥ 95% Similarity", "< 95% Similarity"],
        "Count": [high_similarity_count, low_similarity_count]
    }
    pie_df = pd.DataFrame(pie_data)
    
    # Create pie chart
    fig = px.pie(pie_df, values="Count", names="Category", 
                 color="Category", color_discrete_map={
                     "≥ 95% Similarity": "darkgreen",
                     "< 95% Similarity": "darkgoldenrod"
                 })
    fig.update_layout(
        title={
            "text": "BoQ-EC3 Matching Confidence",
            "font": {"color": "rgb(0, 37, 105)", "size": 20},
            "y": 0.95,
            "x": 0.5,
            "xanchor": "center",
            "yanchor": "top"
        },
        legend={
            "orientation": "h",  # Horizontal legend
            "y": -0.2,  # Position legend below the chart
            "x": 0.5,  # Center the legend
            "xanchor": "center",
            "yanchor": "top"
        },
        margin={"t": 50, "b": 100},  # Add margin for the legend
        height=400
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Display full emissions table
    st.markdown(
        f"""
        <div class="dashboard-title">Initial Machine Calculation</div>
        <h1 style='color: rgb(0, 37, 105); font-size: 28px;'>Full Emissions Table</h1>
        """,
        unsafe_allow_html=True
    )
    
    # Convert uncertainty back to string for display
    emissions_df["GWP Uncertainty"] = emissions_df["GWP Uncertainty"].apply(
        lambda x: f"{x * 100:.1f}%" if isinstance(x, (int, float)) else str(x)
    )
    
    # Ensure "Unsupported Unit" column contains strings
    if "Unsupported Unit" in emissions_df.columns:
        emissions_df["Unsupported Unit"] = emissions_df["Unsupported Unit"].apply(
            lambda x: str(x) if pd.notnull(x) else ""
        )
    
    # Style the table
    styled_df = emissions_df.style.apply(highlight_high_uncertainty_rows, axis=1).apply(color_category_based_on_similarity, axis=1)
    
    # Display the styled table
    st.dataframe(styled_df)
    
    # Download button for the emissions table
    def convert_df_to_csv(df):
        return df.to_csv(index=False).encode('utf-8')
    
    csv = convert_df_to_csv(emissions_df)
    st.download_button(
        label="Download Full Emissions Table as CSV",
        data=csv,
        file_name="full_emissions_table.csv",
        mime="text/csv",
    )
    
def edit_categories_interactive(emissions_df, processed_df):
    """
    Allow interactive editing of EC3 material categories and assumptions.
    """
    # Display the "Manual Override" header with gray background (only once)
    if "manual_override_header_displayed" not in st.session_state:
        st.markdown(
            f"""
            <div class="dashboard-title">Manual Override</div>
            """,
            unsafe_allow_html=True
        )
        st.session_state.manual_override_header_displayed = True
    
    # Add editable columns for thickness and diameter
    emissions_df["User-Provided Thickness (m)"] = emissions_df["Assumed Thickness (m)"]
    emissions_df["User-Provided Diameter (m)"] = emissions_df["Assumed Diameter (m)"]
    
    # Get unique categories from the processed emission factors
    unique_categories = processed_df["EC3 Material Category"].unique().tolist()
    unique_categories.append("None")  # Add an option to deselect
    
    # Select only the relevant columns for the interactive table
    columns_to_show = [
        "BoQ Material",
        "EC3 Category",
        "Similarity",
        "Source PDF",
        "Assumed Thickness (m)",
        "User-Provided Thickness (m)",
        "Assumed Diameter (m)",
        "User-Provided Diameter (m)"
    ]
    interactive_df = emissions_df[columns_to_show]
    
    # Display the "Interactive Table" header without gray background
    st.markdown(
        f"""
        <h1 style='color: rgb(0, 37, 105); font-size: 28px;'>Interactive Table</h1>
        """,
        unsafe_allow_html=True
    )
    
    # Create interactive data editor
    edited_df = st.data_editor(
        interactive_df,
        column_config={
            "EC3 Category": st.column_config.SelectboxColumn(
                "EC3 Material Category",
                help="Select a new category for the material. Choose 'None' to hide this row.",
                options=unique_categories,  # Use the full list of unique categories
                required=True
            ),
            "User-Provided Thickness (m)": st.column_config.NumberColumn(
                "User-Provided Thickness (m)",
                help="Enter the actual thickness for area-based materials.",
                min_value=0.0,
                max_value=1.0,
                step=0.01
            ),
            "User-Provided Diameter (m)": st.column_config.NumberColumn(
                "User-Provided Diameter (m)",
                help="Enter the actual diameter for length-based materials.",
                min_value=0.0,
                max_value=1.0,
                step=0.01
            )
        },
        key="emissions_table"
    )
    
    # Merge the edited columns back into the original emissions_df
    emissions_df.update(edited_df)
    
    # Recalculate quantities and emissions based on user-provided values
    for idx, row in emissions_df.iterrows():
        if pd.notna(row["User-Provided Thickness (m)"]):
            # Recalculate quantity for area-based materials
            density = processed_df.loc[processed_df["EC3 Material Category"] == row["EC3 Category"], "Density"].values[0]
            area = row["BoQ Quantity (KG)"] / (density * row["User-Provided Thickness (m)"])
            emissions_df.at[idx, "BoQ Quantity (KG)"] = area * density * row["User-Provided Thickness (m)"]
        
        if pd.notna(row["User-Provided Diameter (m)"]):
            # Recalculate quantity for length-based materials
            density = processed_df.loc[processed_df["EC3 Material Category"] == row["EC3 Category"], "Density"].values[0]
            length = row["BoQ Quantity (KG)"] / (density * (3.14159 * (row["User-Provided Diameter (m)"] / 2) ** 2))
            emissions_df.at[idx, "BoQ Quantity (KG)"] = length * density * (3.14159 * (row["User-Provided Diameter (m)"] / 2) ** 2)
    
    return emissions_df

# ---------------------------
# --- Main Application ---
# ---------------------------

def main():
    st.markdown(
        "<h1 style='color: rgb(0, 37, 105); font-size: 28px;'>Embodied Emissions Footprint for Construction</h1>", 
        unsafe_allow_html=True
    )
    
    st.sidebar.header("Upload Emission Factors")
    emission_file = st.sidebar.file_uploader("Upload emission factor Excel file", type=["xlsx"])
    
    st.sidebar.header("Upload Bill of Quantities (BoQ) Files")
    boq_files = st.sidebar.file_uploader(
        "Upload BoQ files", 
        type=['pdf', 'docx', 'xlsx', 'csv', 'zip', 'jpg', 'jpeg', 'png'], 
        accept_multiple_files=True
    )

    if boq_files:
        total_uploaded_size = sum(file.size for file in boq_files)
        total_uploaded_size_mb = total_uploaded_size / (1024 * 1024)
        
        st.sidebar.subheader("File Upload Status")
        st.sidebar.progress(total_uploaded_size_mb / STREAMLIT_UPLOAD_LIMIT_MB)
        st.sidebar.info(f"Total Uploaded Size: {total_uploaded_size_mb:.2f} MB / {STREAMLIT_UPLOAD_LIMIT_MB} MB")

    if emission_file and boq_files:
        if st.button("Calculate Emissions", type="primary"):
            with st.spinner("Processing files and calculating emissions..."):
                # FIX: Save the uploaded emission file to a temporary path before reading
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_emission_file:
                    temp_emission_file.write(emission_file.getvalue())
                    temp_emission_path = temp_emission_file.name
                
                try:
                    df_emission = pd.read_excel(temp_emission_path, engine='openpyxl')
                finally:
                    # Clean up the temporary file
                    os.remove(temp_emission_path)
                
                processed_pdf_paths, total_processed_size = process_uploaded_files(boq_files)
                total_processed_size_mb = total_processed_size / (1024 * 1024)

                st.sidebar.subheader("Processed File Size")
                st.sidebar.info(f"Total size after conversion: {total_processed_size_mb:.2f} MB")

                if not processed_pdf_paths:
                    st.error("No processable files found after conversion. Please check your uploaded files.")
                    st.stop()

                combined_boq_df, all_countries = extract_boq_from_processed_pdfs(processed_pdf_paths)
                
                if combined_boq_df.empty:
                    st.warning("No concrete, steel, or asphalt materials detected in the processed files.")
                    st.stop()
                
                st.success(f"Successfully extracted {len(combined_boq_df)} material items focusing on concrete, steel, and asphalt.")
                
                country_name = next(iter(all_countries), None) if all_countries else None
                processed_df, region = process_emission_factors(df_emission, country_name, country_to_region)
                
                if processed_df is not None:
                    st.success("Emission factors processed successfully!")
                    
                    boq_data = combined_boq_df.to_dict('records')
                    matched_results = match_materials_with_categories(boq_data, processed_df)
                    
                    emissions_data, total_emissions = calculate_ghg_emissions(matched_results, processed_df, region)
                    
                    if emissions_data:
                        emissions_df = pd.DataFrame(emissions_data)
                        #print(f"cols: {emissions_df.columns}")
                        emissions_df = emissions_df.rename(columns={"Material": "BoQ Material", "Category": "EC3 Category"})
                        emissions_df["Calculated GHG Emissions (kg CO2e)"] = emissions_df["Calculated GHG Emissions (kg CO2e)"].apply(
                            lambda x: "" if pd.isna(x) else x
                        )
                        emissions_df.index = emissions_df.index + 1
                        
                        st.session_state.emissions_df = emissions_df
                        st.session_state.all_countries = all_countries
                        st.session_state.processed_emission_factors = processed_df

                    else:
                        st.error("No emissions data could be calculated.")
                else:
                    st.error("Could not process emission factors.")

    if 'emissions_df' in st.session_state:
        display_dashboard(st.session_state.emissions_df, st.session_state.all_countries)
        updated_emissions_df = edit_categories_interactive(st.session_state.emissions_df, st.session_state.processed_emission_factors)
        
        st.markdown("<h1 style='color: rgb(0, 37, 105); font-size: 28px;'>Updated Emissions Table</h1>", unsafe_allow_html=True)
        styled_df = updated_emissions_df.style.apply(highlight_high_uncertainty_rows, axis=1).apply(color_category_based_on_similarity, axis=1)
        st.dataframe(styled_df)
        
        csv = convert_df_to_csv(updated_emissions_df)
        st.download_button(
            label="Download Updated Emissions Table as CSV",
            data=csv,
            file_name="updated_emissions_table.csv",
            mime="text/csv",
            key="download_updated_emissions_table"
        )

if __name__ == "__main__":
    main()
