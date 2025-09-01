#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Enhanced DOCX Processor

This script processes DOCX files using lxml-based XML manipulation
for reliable placeholder replacement while preserving as much
formatting as possible, including images and layout.
"""

import pandas as pd
import os
import math
import zipfile
import tempfile
import shutil
import lxml.etree as ET
from docx import Document
from docx.shared import Inches


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


def create_replacement_mapping(df, page_start_index, cells_per_page):
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
def process_single_page_docx(template_path, output_path, replacements):
    """Process a single DOCX file with replacements using lxml-based XML manipulation"""
    
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
            
            # Define namespaces for lxml
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'o': 'urn:schemas-microsoft-com:office:office',
                'v': 'urn:schemas-microsoft-com:vml',
                'w10': 'urn:schemas-microsoft-com:office:word',
                'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
                'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
            }
            
            parser = ET.XMLParser(remove_blank_text=True)
            root = ET.fromstring(xml_content.encode('utf-8'), parser=parser)
            
            replacements_made = 0
            
            # Find all text elements and replace placeholders while preserving formatting
            # Focus on w:t (regular text) and a:t (drawing text) elements
            text_elements = root.xpath('.//w:t', namespaces=namespaces) + root.xpath('.//a:t', namespaces=namespaces)
            
            for elem in text_elements:
                if elem.text:
                    for placeholder, value in replacements.items():
                        if placeholder in elem.text:
                            elem.text = elem.text.replace(placeholder, str(value))
                            replacements_made += 1
                            print(f"      Replaced '{placeholder}' with '{value}'")
            
            # Also check text within v:textbox elements (legacy textboxes)
            v_textbox_elements = root.xpath('.//v:textbox//w:t', namespaces=namespaces)
            for elem in v_textbox_elements:
                if elem.text:
                    for placeholder, value in replacements.items():
                        if placeholder in elem.text:
                            elem.text = elem.text.replace(placeholder, str(value))
                            replacements_made += 1
                            print(f"      Replaced '{placeholder}' with '{value}' in VML textbox")
            
            # Convert back to string with proper XML declaration
            modified_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True, pretty_print=True).decode('utf-8')
            
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
    
    except Exception as e:
        print(f"Error in fallback processing: {e}")
        import traceback
        traceback.print_exc()
        return 0
    
    finally:
        # Clean up temporary directory
        try:
            shutil.rmtree(temp_dir)
        except:
            pass


def combine_docx_files(file_paths, output_path):
    """Combine multiple DOCX files into one while preserving all formatting using lxml"""
    
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
        
        # Parse base XML with namespace registration
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'o': 'urn:schemas-microsoft-com:office:office',
            'v': 'urn:schemas-microsoft-com:vml',
            'w10': 'urn:schemas-microsoft-com:office:word',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        }
        
        parser = ET.XMLParser(remove_blank_text=True)
        root = ET.fromstring(base_xml_content.encode('utf-8'), parser=parser)
        body = root.xpath('.//w:body', namespaces=namespaces)[0]
        
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
            current_root = ET.fromstring(current_xml_content.encode('utf-8'), parser=parser)
            current_body = current_root.xpath('.//w:body', namespaces=namespaces)[0]
            
            # Add page break
            page_break = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
            page_break_run = ET.SubElement(page_break, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
            page_break_elem = ET.SubElement(page_break_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')
            page_break_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'page')
            body.append(page_break)
            
            # Add all content from current body (except sectPr which contains section properties)
            for element in current_body:
                # Skip section properties as they define page layout and should only come from the last section
                if not element.tag.endswith('}sectPr'):
                    # Append the element directly
                    body.append(element)
        
        # Write modified document.xml back with proper formatting
        modified_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True, pretty_print=True).decode('utf-8')
        
        with open(base_doc_xml_path, 'w', encoding='utf-8') as f:
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
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        # Clean up temporary directory
        try:
            shutil.rmtree(temp_dir)
        except:
            pass


def count_placeholders_in_template(template_path):
    """Count the number of unique placeholders in the template"""
    try:
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Extract document.xml from template
            extract_dir = os.path.join(temp_dir, "extracted")
            os.makedirs(extract_dir)
            
            with zipfile.ZipFile(template_path, 'r') as zip_ref:
                zip_ref.extract('word/document.xml', extract_dir)
            
            # Read document.xml
            document_xml_path = os.path.join(extract_dir, "word", "document.xml")
            with open(document_xml_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            
            # Count unique placeholders using regex to catch all occurrences
            import re
            code_placeholders = set(re.findall(r'\[code_\d+\]', xml_content))
            sno_placeholders = set(re.findall(r'\[sno_\d+\]', xml_content))
            unique_placeholders = code_placeholders.union(sno_placeholders)
            
            # Since each cell has two placeholders ([code_n] and [sno_n]),
            # divide by 2 to get the number of cells
            cells_per_page = len(unique_placeholders) // 2
            
            # If no placeholders found with regex, try a more comprehensive search
            if cells_per_page == 0:
                # Look for any text that looks like a placeholder
                all_placeholders = set(re.findall(r'\[[^\]]*\d+[^\]]*\]', xml_content))
                code_like = set([p for p in all_placeholders if 'code' in p])
                sno_like = set([p for p in all_placeholders if 'sno' in p or 'serial' in p.lower()])
                cells_per_page = max(len(code_like), len(sno_like))
            
            print(f"Detected {cells_per_page} cells per page in template")
            return cells_per_page
            
        finally:
            # Clean up temporary directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
                
    except Exception as e:
        print(f"Error counting placeholders in template: {e}")
        return 0


def process_vouchers(csv_path, template_path, output_path):
    """Main processing function"""
    
    # Read CSV data
    df = read_csv_data(csv_path)
    if df is None:
        return False
    
    # Dynamically determine cells per page by counting placeholders in template
    cells_per_page = count_placeholders_in_template(template_path)
    if cells_per_page <= 0:
        print("Error: Could not determine number of cells per page from template")
        return False
    
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