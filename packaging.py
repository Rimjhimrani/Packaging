import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.drawing.image import Image
import io
import base64
from PIL import Image as PILImage
import zipfile
import os
import tempfile
import streamlit as st

class ExactPackagingTemplateManager:
    def __init__(self):
        self.template_fields = {
            # Header Information
            'Revision No.': '',
            'Date': '',
            'QC': '',
            'MM': '',
            'VP': '',
            
            # Vendor Information
            'Vendor Code': '',
            'Vendor Name': '',
            'Vendor Location': '',
            
            # Part Information
            'Part No.': '',
            'Part Description': '',
            'Part Unit Weight': '',
            'Part Weight Unit': '',
            
            # Primary Packaging
            'Primary Packaging Type': '',
            'Primary L-mm': '',
            'Primary W-mm': '',
            'Primary H-mm': '',
            'Primary Qty/Pack': '',
            'Primary Empty Weight': '',
            'Primary Pack Weight': '',
            
            # Secondary Packaging
            'Secondary Packaging Type': '',
            'Secondary L-mm': '',
            'Secondary W-mm': '',
            'Secondary H-mm': '',
            'Secondary Qty/Pack': '',
            'Secondary Empty Weight': '',
            'Secondary Pack Weight': '',
            
            # Packaging Procedures (10 steps)
            'Procedure Step 1': '',
            'Procedure Step 2': '',
            'Procedure Step 3': '',
            'Procedure Step 4': '',
            'Procedure Step 5': '',
            'Procedure Step 6': '',
            'Procedure Step 7': '',
            'Procedure Step 8': '',
            'Procedure Step 9': '',
            'Procedure Step 10': '',
            
            # Approval
            'Issued By': '',
            'Reviewed By': '',
            'Approved By': ''
        }
    
    def create_exact_template_excel(self):
        """Create the exact Excel template matching the image"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Packaging Instruction"
        
        # Define styles
        blue_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        light_blue_fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True, size=12)
        black_font = Font(color="000000", bold=True, size=14)
        regular_font = Font(color="000000", size=12)
        title_font = Font(bold=True, size=12)
        header_font = Font(bold=True)
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Set column widths to match the image exactly
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 11
        ws.column_dimensions['G'].width = 10
        ws.column_dimensions['H'].width = 10
        ws.column_dimensions['I'].width = 10
        ws.column_dimensions['J'].width = 10
        ws.column_dimensions['K'].width = 10
        ws.column_dimensions['L'].width = 18

        # Set row heights
        for row in range(1, 55):
            ws.row_dimensions[row].height = 20

        # Header Row - "Packaging Instruction"
        ws.merge_cells('A1:K1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = blue_fill
        ws['A1'].font = white_font
        ws['A1'].border = border
        ws['A1'].alignment = center_alignment

        # Current Packaging header (right side)
        ws['L1'] = "CURRENT PACKAGING"
        ws['L1'].fill = blue_fill
        ws['L1'].font = white_font
        ws['L1'].border = border
        ws['L1'].alignment = center_alignment

        # Revision information row
        ws['A2'] = "Revision No."
        ws['A2'].border = border
        ws['A2'].alignment = left_alignment
        ws['A2'].font = regular_font

        ws.merge_cells('B2:C2')
        ws['B2'] = ""
        ws['B2'].border = border
        ws['C2'].border = border

        # D2 and E2 are left blank (styled but no value)
        ws['D2'] = ""
        ws['D2'].border = border
        ws['E2'] = ""
        ws['E2'].border = border

        # F2 contains "Date" label
        ws['F2'] = "Date"
        ws['F2'].border = border
        ws['F2'].alignment = left_alignment
        ws['F2'].font = regular_font

        # G2-H2 merged for date value
        ws.merge_cells('G2:J2')
        ws['G2'] = ""
        ws['G2'].border = border
        ws['H2'].border = border

        ws.merge_cells('B3:C3')
        ws['B3'] = ""
        ws['B3'].border = border
        ws['C3'].border = border

        # Empty cells in row 3
        for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}3'] = ""
            ws[f'{col}3'].border = border

        # Empty row 4 with borders (spacing before vendor info)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}4'] = ""
            ws[f'{col}4'].border = border
            
        # Vendor Information Title (Row 5)
        ws.merge_cells('A5:D5')
        ws['A5'] = "Vendor Information"
        ws['A5'].font = title_font
        ws['A5'].alignment = center_alignment
        ws['A5'].border = border

        # Empty cell E5
        ws['E5'] = ""
        ws['E5'].border = border

        # Part Information Title (Row 5)
        ws.merge_cells('F5:I5')
        ws['F5'] = "Part Information"
        ws['F5'].font = title_font
        ws['F5'].alignment = center_alignment
        ws['F5'].border = border

        # Empty cells J5, K5, L5
        for col in ['J', 'K', 'L']:
            ws[f'{col}5'] = ""
            ws[f'{col}5'].border = border

        # Vendor Code Row (Row 6)
        ws['A6'] = "Code"
        ws['A6'].font = header_font
        ws['A6'].alignment = left_alignment
        ws['A6'].border = border

        ws.merge_cells('B6:D6')
        ws['B6'] = ""
        ws['B6'].border = border
        for col in ['C', 'D']:
            ws[f'{col}6'].border = border
        
        # Empty cell E6
        ws['E6'] = ""
        ws['E6'].border = border
        
        # Part fields (Row 6)
        ws['F6'] = "Part No."
        ws['F6'].border = border
        ws['F6'].alignment = left_alignment
        ws['F6'].font = regular_font

        ws.merge_cells('G6:K6')
        ws['G6'] = ""
        ws['G6'].border = border
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}6'].border = border
        
        # L6
        ws['L6'] = ""
        ws['L6'].border = border

        # Vendor Name Row (Row 7)
        ws['A7'] = "Name"
        ws['A7'].font = header_font
        ws['A7'].alignment = left_alignment
        ws['A7'].border = border

        ws.merge_cells('B7:D7')
        ws['B7'] = ""
        ws['B7'].border = border
        for col in ['C', 'D']:
            ws[f'{col}7'].border = border

        # Empty cell E7
        ws['E7'] = ""
        ws['E7'].border = border

        # Part Description (Row 7)
        ws['F7'] = "Description"
        ws['F7'].border = border
        ws['F7'].alignment = left_alignment
        ws['F7'].font = regular_font

        ws.merge_cells('G7:K7')
        ws['G7'] = ""
        ws['G7'].border = border
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}7'].border = border
        
        # L7
        ws['L7'] = ""
        ws['L7'].border = border

        # Vendor Location Row (Row 8)
        ws['A8'] = "Location"
        ws['A8'].font = header_font
        ws['A8'].alignment = left_alignment
        ws['A8'].border = border

        ws.merge_cells('B8:D8')
        ws['B8'] = ""
        ws['B8'].border = border
        for col in ['C', 'D']:
            ws[f'{col}8'].border = border

        # Empty cell E8
        ws['E8'] = ""
        ws['E8'].border = border

        # Part Unit Weight (Row 8)
        ws['F8'] = "Unit Weight"
        ws['F8'].border = border
        ws['F8'].alignment = left_alignment
        ws['F8'].font = regular_font

        ws.merge_cells('G8:K8')
        ws['G8'] = ""
        ws['G8'].border = border
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}8'].border = border

        # L8
        ws['L8'] = ""
        ws['L8'].border = border

        # Additional row for L, W, H (Row 9)
        ws['A9'] = ""
        ws['A9'].border = border
        
        ws['B9'] = ""
        ws['B9'].border = border
        
        ws['C9'] = ""
        ws['C9'].border = border
        
        ws['D9'] = ""
        ws['D9'].border = border
        
        ws['E9'] = ""
        ws['E9'].border = border

        ws['F9'] = "L"
        ws['F9'].border = border
        ws['F9'].alignment = center_alignment
        ws['F9'].font = regular_font

        ws['G9'] = ""
        ws['G9'].border = border

        ws['H9'] = "W"
        ws['H9'].border = border
        ws['H9'].alignment = center_alignment
        ws['H9'].font = regular_font

        ws['I9'] = ""
        ws['I9'].border = border

        ws['J9'] = "H"
        ws['J9'].border = border
        ws['J9'].alignment = center_alignment
        ws['J9'].font = regular_font

        ws['K9'] = ""
        ws['K9'].border = border

        ws['L9'] = ""
        ws['L9'].border = border

        # Empty row 10 (spacing before primary packaging)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}10'] = ""
            ws[f'{col}10'].border = border

        # Primary Packaging Title (Row 11)
        ws.merge_cells('A11:K11')
        ws['A11'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A11'].fill = blue_fill
        ws['A11'].font = white_font
        ws['A11'].border = border
        ws['A11'].alignment = center_alignment

        ws['L11'] = "CURRENT PACKAGING"
        ws['L11'].fill = blue_fill
        ws['L11'].font = white_font
        ws['L11'].border = border
        ws['L11'].alignment = center_alignment

        # Primary packaging headers (Row 12)
        headers = ["Packaging Type", "L-mm", "W-mm", "H-mm", "Qty/Pack", "Empty Weight", "Pack Weight"]
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}12'] = header
            ws[f'{col}12'].border = border
            ws[f'{col}12'].alignment = center_alignment
            ws[f'{col}12'].font = regular_font
        
        # Empty cells for remaining columns in row 12
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}12'] = ""
            ws[f'{col}12'].border = border
        ws['L12'] = ""
        ws['L12'].border = border

        # Primary packaging data rows (13-15)
        for row in range(13, 16):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border
        
        # TOTAL row (Row 15)
        ws['D15'] = "TOTAL"
        ws['D15'].border = border
        ws['D15'].font = black_font
        ws['D15'].alignment = center_alignment

        # Empty row 16 (spacing before secondary packaging)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}16'] = ""
            ws[f'{col}16'].border = border

        # Secondary Packaging Instruction header (Row 17)
        ws.merge_cells('A17:K17')
        ws['A17'] = "Secondary Packaging Instruction (Outer / External)"
        ws['A17'].fill = blue_fill
        ws['A17'].font = white_font
        ws['A17'].border = border
        ws['A17'].alignment = center_alignment

        ws['K17'] = ""
        ws['K17'].border = border

        ws['L12'] = "PROBLEM IF ANY:"
        ws['L12'].border = border
        ws['L12'].font = black_font
        ws['L12'].alignment = left_alignment

        # Secondary packaging headers (Row 18)
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}18'] = header
            ws[f'{col}18'].border = border
            ws[f'{col}18'].alignment = center_alignment
            ws[f'{col}18'].font = regular_font
        
        # Empty cells for remaining columns in row 18
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}18'] = ""
            ws[f'{col}18'].border = border

        ws['L13'] = "CAUTION: REVISED DESIGN"
        ws['L13'].fill = red_fill
        ws['L13'].font = white_font
        ws['L13'].border = border
        ws['L13'].alignment = center_alignment

        # Secondary packaging data rows (19-21)
        for row in range(19, 22):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border
        
        # TOTAL row for secondary (Row 21)
        ws['D21'] = "TOTAL"
        ws['D21'].border = border
        ws['D21'].font = black_font
        ws['D21'].alignment = center_alignment

        # Empty row 22 (spacing before packaging procedure)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}22'] = ""
            ws[f'{col}22'].border = border

        # Packaging Procedure section (Row 23)
        ws.merge_cells('A23:K23')
        ws['A23'] = "Packaging Procedure"
        ws['A23'].fill = blue_fill
        ws['A23'].font = white_font
        ws['A23'].border = border
        ws['A23'].alignment = center_alignment

        ws['L23'] = ""
        ws['L23'].border = border

        # Packaging procedure steps (rows 24-33) - WITH MERGED CELLS
        for i in range(1, 11):
            row = 23 + i
            ws[f'A{row}'] = str(i)
            ws[f'A{row}'].border = border
            ws[f'A{row}'].alignment = center_alignment
            ws[f'A{row}'].font = regular_font
    
            # MERGE CELLS B to J for each procedure step
            ws.merge_cells(f'B{row}:J{row}')
            ws[f'B{row}'] = ""
            ws[f'B{row}'].border = border
            ws[f'B{row}'].alignment = left_alignment
    
            # Add borders to all merged cells
            for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'].border = border
    
            # K and L columns
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Empty row 34 (spacing before reference images)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}34'] = ""
            ws[f'{col}34'].border = border

        # Reference Images/Pictures section (Row 35)
        ws.merge_cells('A35:K35')
        ws['A35'] = "Reference Images/Pictures"
        ws['A35'].fill = blue_fill
        ws['A35'].font = white_font
        ws['A35'].border = border
        ws['A35'].alignment = center_alignment

        ws['L35'] = ""
        ws['L35'].border = border

        # Image section headers (Row 36)
        ws.merge_cells('A36:C36')
        ws['A36'] = "Primary Packaging"
        ws['A36'].border = border
        ws['A36'].alignment = center_alignment
        ws['A36'].font = regular_font

        ws.merge_cells('D36:F36')
        ws['D36'] = "Secondary Packaging"
        ws['D36'].border = border
        ws['D36'].alignment = center_alignment
        ws['D36'].font = regular_font

        ws.merge_cells('G36:J36')
        ws['G36'] = "Label"
        ws['G36'].border = border
        ws['G36'].alignment = center_alignment
        ws['G36'].font = regular_font

        ws['K36'] = ""
        ws['K36'].border = border
        ws['L36'] = ""
        ws['L36'].border = border

        # Image placeholder areas (rows 37-42)
        ws.merge_cells('A37:C42')
        ws['A37'] = "Primary\nPackaging"
        ws['A37'].border = border
        ws['A37'].alignment = center_alignment
        ws['A37'].font = regular_font

        # Arrow 1 (Row 39-40)
        ws.merge_cells('D39:D40')
        ws['D39'] = "→"
        ws['D39'].border = border
        ws['D39'].alignment = center_alignment
        ws['D39'].font = Font(size=20, bold=True)

        # Secondary Packaging image area
        ws.merge_cells('E37:F42')
        ws['E37'] = "SECONDARY\nPACKAGING"
        ws['E37'].border = border
        ws['E37'].alignment = center_alignment
        ws['E37'].font = regular_font
        ws['E37'].fill = light_blue_fill

        # Arrow 2 (Row 39-40)
        ws.merge_cells('G39:G40')
        ws['G39'] = "→"
        ws['G39'].border = border
        ws['G39'].alignment = center_alignment
        ws['G39'].font = Font(size=20, bold=True)

        # Label image area
        ws.merge_cells('H37:K42')
        ws['H37'] = "LABEL"
        ws['H37'].border = border
        ws['H37'].alignment = center_alignment
        ws['H37'].font = regular_font

        # Add borders to all cells in image area
        for row in range(37, 43):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'].border = border

        # Empty row 43 (spacing before approval)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}43'] = ""
            ws[f'{col}43'].border = border

        # Approval Section (Row 44)
        ws.merge_cells('A44:C44')
        ws['A44'] = "Issued By"
        ws['A44'].border = border
        ws['A44'].alignment = center_alignment
        ws['A44'].font = regular_font

        ws.merge_cells('D44:G44')
        ws['D44'] = "Reviewed By"
        ws['D44'].border = border
        ws['D44'].alignment = center_alignment
        ws['D44'].font = regular_font

        ws.merge_cells('H44:K44')
        ws['H44'] = "Approved By"
        ws['H44'].border = border
        ws['H44'].alignment = center_alignment
        ws['H44'].font = regular_font

        ws['L44'] = ""
        ws['L44'].border = border

        # Signature boxes (rows 45-48)
        ws.merge_cells('A45:C48')
        ws['A45'] = ""
        ws['A45'].border = border

        ws.merge_cells('D45:G48')
        ws['D45'] = ""
        ws['D45'].border = border

        ws.merge_cells('H45:K48')
        ws['H45'] = ""
        ws['H45'].border = border

        # Apply borders for signature section
        for row in range(45, 49):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'].border = border

        # Second Approval Section (Row 49)
        ws.merge_cells('A49:C49')
        ws['A49'] = "Issued By"
        ws['A49'].border = border
        ws['A49'].alignment = center_alignment
        ws['A49'].font = regular_font

        ws.merge_cells('D49:G49')
        ws['D49'] = "Reviewed By"
        ws['D49'].border = border
        ws['D49'].alignment = center_alignment
        ws['D49'].font = regular_font

        ws.merge_cells('H49:K49')
        ws['H49'] = "Approved By"
        ws['H49'].border = border
        ws['H49'].alignment = center_alignment
        ws['H49'].font = regular_font

        ws['L49'] = ""
        ws['L49'].border = border

        # Second signature boxes (rows 50-53)
        ws.merge_cells('A50:C53')
        ws['A50'] = ""
        ws['A50'].border = border

        ws.merge_cells('D50:G53')
        ws['D50'] = ""
        ws['D50'].border = border

        ws.merge_cells('H50:K53')
        ws['H50'] = ""
        ws['H50'].border = border

        # Apply borders for second signature section
        for row in range(50, 54):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'].border = border

        # Return the workbook
        return wb
    
    def extract_data_from_uploaded_file(self, uploaded_file):
        """Extract data from uploaded CSV/Excel file"""
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # Convert dataframe to dictionary
            data_dict = {}
            for col in df.columns:
                col_clean = col.strip()  # Remove any whitespace
                if col_clean in self.template_fields:
                    data_dict[col_clean] = str(df[col].iloc[0]) if not df.empty and pd.notna(df[col].iloc[0]) else ""
            
            return data_dict, df
        except Exception as e:
            raise Exception(f"Error reading file: {str(e)}")
    
    def extract_images_from_file(self, uploaded_file):
        """Extract images from uploaded file if it's an Excel file"""
        images = {}
        try:
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                # Create a temporary file to save the uploaded file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                # Try to extract images from Excel file
                try:
                    wb = load_workbook(tmp_path)
                    ws = wb.active
                    
                    # Check if there are any images in the worksheet
                    if hasattr(ws, '_images'):
                        for i, img in enumerate(ws._images):
                            # Convert image to bytes
                            img_bytes = io.BytesIO()
                            img.image.save(img_bytes, format='PNG')
                            img_bytes.seek(0)
                            images[f'image_{i+1}'] = img_bytes.getvalue()
                    
                    wb.close()
                except Exception as e:
                    st.warning(f"Could not extract images from Excel file: {str(e)}")
                
                # Clean up temporary file
                try:
                    os.unlink(tmp_path)
                except:
                    pass
        except Exception as e:
            st.error(f"Error processing file for images: {str(e)}")
        
        return images
    
    def fill_template_with_data(self, template_wb, data_dict):
        """Fill the template with provided data"""
        ws = template_wb.active
        
        # Corrected mapping of data keys to cell positions
        cell_mapping = {
            'Revision No.': 'B2',
            'Date': 'G2',
            'QC': 'J2',
            'MM': 'L2',
            'VP': 'B3',
            'Vendor Code': 'B6',
            'Vendor Name': 'B7',
            'Vendor Location': 'B8',
            'Part No.': 'G6',
            'Part Description': 'G7',
            'Part Unit Weight': 'G8',
            'Primary Packaging Type': 'A13',
            'Primary L-mm': 'B13',
            'Primary W-mm': 'C13',
            'Primary H-mm': 'D13',
            'Primary Qty/Pack': 'E13',
            'Primary Empty Weight': 'F13',
            'Primary Pack Weight': 'G13',
            'Secondary Packaging Type': 'A19',
            'Secondary L-mm': 'B19',
            'Secondary W-mm': 'C19',
            'Secondary H-mm': 'D19',
            'Secondary Qty/Pack': 'E19',
            'Secondary Empty Weight': 'F19',
            'Secondary Pack Weight': 'G19',
            'Issued By': 'A44',
            'Reviewed By': 'D44',
            'Approved By': 'H44'
        }
        
        # Fill procedure steps in merged cells (B24:B33)
        for i in range(1, 11):
            step_key = f'Procedure Step {i}'
            if step_key in data_dict and data_dict[step_key]:
                row = 23 + i
                ws[f'B{row}'] = data_dict[step_key]
        
        # Fill other mapped fields
        for key, cell in cell_mapping.items():
            if key in data_dict and data_dict[key]:
                ws[cell] = data_dict[key]
        
        return template_wb
    
    def insert_images_into_template(self, template_wb, images):
        """Insert images into the template at appropriate locations"""
        ws = template_wb.active
        
        # Define image placement areas
        image_positions = {
            'primary_packaging': {'anchor': 'A37', 'width': 200, 'height': 150},
            'secondary_packaging': {'anchor': 'E37', 'width': 150, 'height': 150},
            'label': {'anchor': 'H37', 'width': 250, 'height': 150}
        }
        
        try:
            # Insert images if available
            for img_name, img_bytes in images.items():
                if img_bytes:
                    # Create PIL Image from bytes
                    pil_img = PILImage.open(io.BytesIO(img_bytes))
                    
                    # Save to BytesIO as PNG
                    img_io = io.BytesIO()
                    pil_img.save(img_io, format='PNG')
                    img_io.seek(0)
                    
                    # Create openpyxl Image
                    img = Image(img_io)
                    
                    # Determine position based on image name or order
                    if 'primary' in img_name.lower() or img_name == 'image_1':
                        pos = image_positions['primary_packaging']
                    elif 'secondary' in img_name.lower() or img_name == 'image_2':
                        pos = image_positions['secondary_packaging']
                    elif 'label' in img_name.lower() or img_name == 'image_3':
                        pos = image_positions['label']
                    else:
                        # Default to first available position
                        pos = image_positions['primary_packaging']
                    
                    # Resize image
                    img.width = pos['width']
                    img.height = pos['height']
                    
                    # Add image to worksheet
                    ws.add_image(img, pos['anchor'])
        
        except Exception as e:
            st.warning(f"Could not insert images: {str(e)}")
        
        return template_wb
    
    def create_filled_template(self, data_dict=None, images=None):
        """Create a filled template with provided data and images"""
        # Create the base template
        wb = self.create_exact_template_excel()
        
        # Fill with data if provided
        if data_dict:
            wb = self.fill_template_with_data(wb, data_dict)
        
        # Insert images if provided
        if images:
            wb = self.insert_images_into_template(wb, images)
        
        return wb
    
    def save_template_to_bytes(self, workbook):
        """Save workbook to bytes for download"""
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output.getvalue()

# Streamlit Application
def main():
    st.set_page_config(page_title="Exact Packaging Template Manager", layout="wide")
    
    st.title("📦 Exact Packaging Template Manager")
    st.markdown("Generate exact packaging instruction templates with your data")
    
    # Initialize the template manager
    template_manager = ExactPackagingTemplateManager()
    
    # Sidebar for options
    st.sidebar.header("Options")
    mode = st.sidebar.selectbox("Select Mode", ["Generate Empty Template", "Fill Template with Data"])
    
    if mode == "Generate Empty Template":
        st.header("Generate Empty Template")
        st.write("Click the button below to generate an empty packaging instruction template.")
        
        if st.button("Generate Empty Template"):
            try:
                # Create empty template
                wb = template_manager.create_exact_template_excel()
                template_bytes = template_manager.save_template_to_bytes(wb)
                
                st.success("Empty template generated successfully!")
                st.download_button(
                    label="Download Empty Template",
                    data=template_bytes,
                    file_name="Packaging_Instruction_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error generating template: {str(e)}")
    
    elif mode == "Fill Template with Data":
        st.header("Fill Template with Data")
        
        # File upload section
        uploaded_file = st.file_uploader(
            "Upload CSV or Excel file with data",
            type=['csv', 'xlsx', 'xls'],
            help="Upload a file containing the data to fill the template"
        )
        
        if uploaded_file is not None:
            try:
                # Extract data from file
                data_dict, df = template_manager.extract_data_from_uploaded_file(uploaded_file)
                
                # Show data preview
                st.subheader("Data Preview")
                st.dataframe(df)
                
                # Extract images if available
                images = template_manager.extract_images_from_file(uploaded_file)
                if images:
                    st.success(f"Found {len(images)} images in the file")
                
                # Show extracted data
                st.subheader("Extracted Data for Template")
                if data_dict:
                    for key, value in data_dict.items():
                        st.write(f"**{key}:** {value}")
                else:
                    st.warning("No matching data found for template fields")
                
                # Generate filled template
                if st.button("Generate Filled Template"):
                    try:
                        wb = template_manager.create_filled_template(data_dict, images)
                        template_bytes = template_manager.save_template_to_bytes(wb)
                        
                        st.success("Filled template generated successfully!")
                        st.download_button(
                            label="Download Filled Template",
                            data=template_bytes,
                            file_name="Filled_Packaging_Instruction_Template.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"Error generating filled template: {str(e)}")
            
            except Exception as e:
                st.error(f"Error processing uploaded file: {str(e)}")
        
        else:
            st.info("Please upload a CSV or Excel file to fill the template with data")
    
    # Show template field information
    st.sidebar.header("Template Fields")
    with st.sidebar.expander("View All Template Fields"):
        for field in template_manager.template_fields.keys():
            st.write(f"• {field}")

if __name__ == "__main__":
    main()
