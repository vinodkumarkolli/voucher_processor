#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Direct DOCX Processor

This script manipulates DOCX files at the ZIP/XML level to ensure
reliable placeholder replacement while preserving all formatting,
images, and layout exactly as in the template.
"""

import pandas as pd
import os
import math
import zipfile
import tempfile
import shutil
from xml.etree import ElementTree as ET


def read_csv_data(csv_path):
    """Read and validate CSV data"""
    try:
        df = pd.read_csv(csv_path)
        print(f"Successfully loaded CSV with {len(df)} records")
        print(f"Columns: {list(df.columns)}")
        
        # Validate required columns
        required_columns = ['ID', 'Secret Code']
        for col in required_columns:
            if col not in df.columns:
                print(f"Error: Required column '{col}' not found in CSV")
                return None
        
        return df
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return None


def create_replacement_mapping(df, page_start_index, cells_per_page=14):
    """Create mapping of placeholders to actual values for a page"""
    replacements = {}
    
    # Process records for this page
    page_records = min(cells_per_page, len(df) - page_start_index)
    
    # Replace placeholders for records that exist
    for i in range(page_records):
        record_index = page_start_index + i
        cell_number = i + 1  # 1-based indexing for placeholders
        
        if record_index < len(df):
            row_data = df.iloc[record_index]
            
            # Map placeholders to actual values
            id_value = str(row_data['ID'])
            code_value = str(row_data['Secret Code'])
            
            replacements[f'[sno_{cell_number}]'] = id_value
            replacements[f'[code_{cell_number}]'] = code_value
            
            print(f"    Cell {cell_number}: [sno_{cell_number}] -> {id_value}, [code_{cell_number}] -> {code_value}")
    
    # Clear unused placeholders with empty text
    for i in range(page_records, cells_per_page):
        cell_number = i + 1
        replacements[f'[sno_{cell_number}]'] = ""
        replacements[f'[code_{cell_number}]'] = ""
        print(f"    Cell {cell_number}: [sno_{cell_number}] -> (empty), [code_{cell_number}] -> (empty)")
    
    return replacements


def replace_text_in_xml(xml_content, replacements):
    """Replace placeholders in XML content"""
    replacements_made = 0
    
    for placeholder, value in replacements.items():
        if placeholder in xml_content:
            xml_content = xml_content.replace(placeholder, str(value))
            replacements_made += 1
            print(f"      Replaced '{placeholder}' with '{value}'")
    
    return xml_content, replacements_made


def process_single_page_docx(template_path, output_path, replacements):
    """Process a single DOCX file with replacements"""
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Copy template to temporary location
        temp_template = os.path.join(temp_dir, "temp.docx")
        shutil.copy2(template_path, temp_template)
        
        # Extract DOCX contents
        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir)
        
        with zipfile.ZipFile(temp_template, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        # Process document.xml
        document_xml_path = os.path.join(extract_dir, "word", "document.xml")
        
        if os.path.exists(document_xml_path):
            with open(document_xml_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            
            # Replace placeholders
            modified_xml, replacements_made = replace_text_in_xml(xml_content, replacements)
            
            # Write back the modified XML
            with open(document_xml_path, 'w', encoding='utf-8') as f:
                f.write(modified_xml)
            
            print(f"      Made {replacements_made} replacements in document.xml")
        else:
            print(f"      Warning: document.xml not found")
            replacements_made = 0
        
        # Recreate the DOCX file
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, extract_dir)
                    zip_ref.write(file_path, arc_name)
        
        return replacements_made
    
    finally:
        # Clean up temporary directory
        try:
            shutil.rmtree(temp_dir)
        except:
            pass


def combine_docx_files(file_paths, output_path):
    """Combine multiple DOCX files into one"""
    
    if not file_paths:
        return False
    
    # Create temporary directory for processing
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Extract first document as base
        base_extract_dir = os.path.join(temp_dir, "base")
        os.makedirs(base_extract_dir)
        
        with zipfile.ZipFile(file_paths[0], 'r') as zip_ref:
            zip_ref.extractall(base_extract_dir)
        
        # Read base document.xml
        base_doc_xml_path = os.path.join(base_extract_dir, "word", "document.xml")
        with open(base_doc_xml_path, 'r', encoding='utf-8') as f:
            base_xml_content = f.read()
        
        # Parse base XML
        root = ET.fromstring(base_xml_content)
        body = root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
        
        # Add content from subsequent documents
        for i in range(1, len(file_paths)):
            # Extract current document
            current_extract_dir = os.path.join(temp_dir, f"doc_{i}")
            os.makedirs(current_extract_dir)
            
            with zipfile.ZipFile(file_paths[i], 'r') as zip_ref:
                zip_ref.extractall(current_extract_dir)
            
            # Read current document.xml
            current_doc_xml_path = os.path.join(current_extract_dir, "word", "document.xml")
            with open(current_doc_xml_path, 'r', encoding='utf-8') as f:
                current_xml_content = f.read()
            
            # Parse current XML and get body content
            current_root = ET.fromstring(current_xml_content)
            current_body = current_root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
            
            # Add page break
            page_break = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
            page_break_run = ET.SubElement(page_break, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
            page_break_elem = ET.SubElement(page_break_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')
            page_break_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'page')
            body.append(page_break)
            
            # Add all content from current body (except sectPr)
            for element in current_body:
                if not element.tag.endswith('}sectPr'):
                    body.append(element)
        
        # Write modified document.xml back
        modified_xml = ET.tostring(root, encoding='unicode')
        with open(base_doc_xml_path, 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
            f.write(modified_xml)
        
        # Create final DOCX file
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root_dir, dirs, files in os.walk(base_extract_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, base_extract_dir)
                    zip_ref.write(file_path, arc_name)
        
        return True
    
    except Exception as e:
        print(f"Error combining DOCX files: {e}")
        return False
    
    finally:
        # Clean up temporary directory
        try:
            shutil.rmtree(temp_dir)
        except:
            pass


def process_vouchers(csv_path, template_path, output_path):
    """Main processing function"""
    
    # Read CSV data
    df = read_csv_data(csv_path)
    if df is None:
        return False
    
    # Define cells per page (14 as per the template structure)
    cells_per_page = 14
    
    # Calculate number of pages needed
    total_records = len(df)
    pages_needed = math.ceil(total_records / cells_per_page)
    
    print(f"Total records: {total_records}")
    print(f"Cells per page: {cells_per_page}")
    print(f"Pages needed: {pages_needed}")
    
    # Create temporary directory for page files
    temp_dir = tempfile.mkdtemp()
    page_files = []
    
    try:
        # Process each page
        for page_num in range(pages_needed):
            start_index = page_num * cells_per_page
            
            print(f"\nProcessing page {page_num + 1}/{pages_needed}")
            print(f"  Records {start_index} to {min(start_index + cells_per_page - 1, total_records - 1)}")
            
            # Create replacement mapping for this page
            replacements = create_replacement_mapping(df, start_index, cells_per_page)
            
            # Create temporary file for this page
            page_file_path = os.path.join(temp_dir, f"page_{page_num + 1}.docx")
            
            # Process this page
            replacements_count = process_single_page_docx(template_path, page_file_path, replacements)
            print(f"    Made {replacements_count} replacements on page {page_num + 1}")
            
            page_files.append(page_file_path)
        
        # Combine all pages into final document
        print(f"\nCombining {len(page_files)} pages into final document...")
        success = combine_docx_files(page_files, output_path)
        
        return success
    
    finally:
        # Clean up temporary files
        try:
            shutil.rmtree(temp_dir)
        except:
            pass


def convert_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF using LibreOffice"""
    
    print(f"Converting {docx_path} to {pdf_path}")
    
    try:
        import subprocess
        result = subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', os.path.dirname(pdf_path),
            docx_path
        ], check=True, capture_output=True, text=True)
        
        print(f"✓ PDF generated successfully: {pdf_path}")
        return True
    except Exception as e:
        print(f"LibreOffice conversion error: {e}")
        return False


def main():
    """Main function"""
    
    print("=== Direct DOCX Processor ===\n")
    
    # Create output directory
    os.makedirs('fileOut', exist_ok=True)
    print("Output directory ready")
    
    # File paths
    csv_path = "filesIn/Gift Voucher.csv"
    template_path = "filesIn/template.docx"
    output_path = "fileOut/template_output.docx"
    
    # Verify input files exist
    if not os.path.exists(csv_path):
        print(f"Error: CSV file not found: {csv_path}")
        return False
    
    if not os.path.exists(template_path):
        print(f"Error: Template file not found: {template_path}")
        return False
    
    print(f"Input files verified:")
    print(f"  CSV: {csv_path}")
    print(f"  Template: {template_path}")
    print(f"  Output: {output_path}\n")
    
    # Process vouchers
    success = process_vouchers(csv_path, template_path, output_path)
    
    if success:
        print(f"\n=== Processing Completed Successfully! ===")
        print(f"Output file: {output_path}")
        
        # Try to convert to PDF
        pdf_path = "fileOut/template_output.pdf"
        print(f"\nAttempting PDF conversion...")
        convert_to_pdf(output_path, pdf_path)
    else:
        print("\n=== Processing Failed ===")
    
    return success


if __name__ == "__main__":
    main()