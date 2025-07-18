import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.drawing.image import Image
from openpyxl.cell.cell import MergedCell
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
    
    def apply_border_to_range(self, ws, start_cell, end_cell):
        """Apply borders to a range of cells"""
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Parse cell references
        start_col = ord(start_cell[0]) - ord('A')
        start_row = int(start_cell[1:])
        end_col = ord(end_cell[0]) - ord('A')
        end_row = int(end_cell[1:])
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col+1)
                cell.border = border
    
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
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        title_font = Font(bold=True, size=12)
        header_font = Font(bold=True)
        
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
        for row in range(1, 51):
            ws.row_dimensions[row].height = 20

        # Header Row - "Packaging Instruction"
        ws.merge_cells('A1:K1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = blue_fill
        ws['A1'].font = white_font
        ws['A1'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A1', 'K1')

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

        ws.merge_cells('B2:E2')
        ws['B2'] = "Revision 1"
        ws['B2'].border = border
        self.apply_border_to_range(ws, 'B2', 'E2')

        # Date field
        ws['F2'] = "Date"
        ws['F2'].border = border
        ws['F2'].alignment = left_alignment
        ws['F2'].font = regular_font

        # Merge cells for date value
        ws.merge_cells('G2:K2')
        ws['G2'] = ""
        ws['G2'].border = border
        self.apply_border_to_range(ws, 'G2', 'K2')

        ws['L2'] = ""
        ws['L2'].border = border

        # Row 3 - empty with borders
        ws.merge_cells('B3:E3')
        ws['B3'] = ""
        self.apply_border_to_range(ws, 'A3', 'L3')

        # Row 4 - Section headers
        ws.merge_cells('A4:D4')
        ws['A4'] = "Vendor Information"
        ws['A4'].font = title_font
        ws['A4'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A4', 'D4')

        ws['E4'] = ""
        ws['E4'].border = border

        ws.merge_cells('F4:I4')
        ws['F4'] = "Part Information"
        ws['F4'].font = title_font
        ws['F4'].alignment = center_alignment
        self.apply_border_to_range(ws, 'F4', 'I4')

        # Apply borders to remaining cells in row 4
        for col in ['J', 'K', 'L']:
            ws[f'{col}4'] = ""
            ws[f'{col}4'].border = border

        # Vendor Code Row
        ws['A5'] = "Code"
        ws['A5'].font = header_font
        ws['A5'].alignment = left_alignment
        ws['A5'].border = border

        ws.merge_cells('B5:D5')
        ws['B5'] = ""
        self.apply_border_to_range(ws, 'B5', 'D5')

        ws['E5'] = ""
        ws['E5'].border = border
        
        # Part fields
        ws['F5'] = "Part No."
        ws['F5'].border = border
        ws['F5'].alignment = left_alignment
        ws['F5'].font = regular_font

        ws.merge_cells('G5:K5')
        ws['G5'] = ""
        self.apply_border_to_range(ws, 'G5', 'K5')

        ws['L5'] = ""
        ws['L5'].border = border

        # Vendor Name Row
        ws['A6'] = "Name"
        ws['A6'].font = header_font
        ws['A6'].alignment = left_alignment
        ws['A6'].border = border

        ws.merge_cells('B6:D6')
        ws['B6'] = ""
        self.apply_border_to_range(ws, 'B6', 'D6')

        ws['E6'] = ""
        ws['E6'].border = border

        ws['F6'] = "Description"
        ws['F6'].border = border
        ws['F6'].alignment = left_alignment
        ws['F6'].font = regular_font

        ws.merge_cells('G6:K6')
        ws['G6'] = ""
        self.apply_border_to_range(ws, 'G6', 'K6')

        ws['L6'] = ""
        ws['L6'].border = border

        # Vendor Location Row
        ws['A7'] = "Location"
        ws['A7'].font = header_font
        ws['A7'].alignment = left_alignment
        ws['A7'].border = border

        ws.merge_cells('B7:D7')
        ws['B7'] = ""
        self.apply_border_to_range(ws, 'B7', 'D7')

        ws['E7'] = ""
        ws['E7'].border = border

        ws['F7'] = "Unit Weight"
        ws['F7'].border = border
        ws['F7'].alignment = left_alignment
        ws['F7'].font = regular_font

        ws.merge_cells('G7:K7')
        ws['G7'] = ""
        self.apply_border_to_range(ws, 'G7', 'K7')

        ws['L7'] = ""
        ws['L7'].border = border

        # Additional row after Unit Weight (Row 8) for L, W, H
        ws['F8'] = "L"
        ws['F8'].border = border
        ws['F8'].alignment = left_alignment
        ws['F8'].font = regular_font

        ws['G8'] = ""
        ws['G8'].border = border

        ws['H8'] = "W"
        ws['H8'].border = border
        ws['H8'].alignment = center_alignment
        ws['H8'].font = regular_font

        ws['I8'] = ""
        ws['I8'].border = border

        ws['J8'] = "H"
        ws['J8'].border = border
        ws['J8'].alignment = center_alignment
        ws['J8'].font = regular_font

        ws['K8'] = ""
        ws['K8'].border = border

        # Empty cells for A-E and L in row 8
        for col in ['A', 'B', 'C', 'D', 'E', 'L']:
            ws[f'{col}8'] = ""
            ws[f'{col}8'].border = border

        # Title row for Primary Packaging
        ws.merge_cells('A9:K9')
        ws['A9'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A9'].fill = blue_fill
        ws['A9'].font = white_font
        ws['A9'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A9', 'K9')

        ws['L9'] = "CURRENT PACKAGING"
        ws['L9'].fill = blue_fill
        ws['L9'].font = white_font
        ws['L9'].border = border
        ws['L9'].alignment = center_alignment

        # Primary packaging headers
        headers = ["Packaging Type", "L-mm", "W-mm", "H-mm", "Qty/Pack", "Empty Weight", "Pack Weight"]
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}10'] = header
            ws[f'{col}10'].border = border
            ws[f'{col}10'].alignment = center_alignment
            ws[f'{col}10'].font = regular_font

        # Empty cells for remaining columns in row 10
        for col in ['H', 'I', 'J', 'K', 'L']:
            ws[f'{col}10'] = ""
            ws[f'{col}10'].border = border

        # Primary packaging data rows (11-13)
        for row in range(11, 14):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border

        # TOTAL row
        ws['D13'] = "TOTAL"
        ws['D13'].border = border
        ws['D13'].font = black_font
        ws['D13'].alignment = center_alignment

        # Secondary Packaging Instruction header
        ws.merge_cells('A14:J14')
        ws['A14'] = "Secondary Packaging Instruction (Outer / External)"
        ws['A14'].fill = blue_fill
        ws['A14'].font = white_font
        ws['A14'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A14', 'J14')

        ws['K14'] = ""
        ws['K14'].border = border

        ws['L14'] = "PROBLEM IF ANY:"
        ws['L14'].border = border
        ws['L14'].font = black_font
        ws['L14'].alignment = left_alignment

        # Secondary packaging headers
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}15'] = header
            ws[f'{col}15'].border = border
            ws[f'{col}15'].alignment = center_alignment
            ws[f'{col}15'].font = regular_font

        # Empty cells for remaining columns in row 15
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}15'] = ""
            ws[f'{col}15'].border = border

        ws['L15'] = "CAUTION: REVISED DESIGN"
        ws['L15'].fill = red_fill
        ws['L15'].font = white_font
        ws['L15'].border = border
        ws['L15'].alignment = center_alignment

        # Secondary packaging data rows (16-18)
        for row in range(16, 19):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border

        # TOTAL row for secondary
        ws['D18'] = "TOTAL"
        ws['D18'].border = border
        ws['D18'].font = black_font
        ws['D18'].alignment = center_alignment

        # Packaging Procedure section
        ws.merge_cells('A19:K19')
        ws['A19'] = "Packaging Procedure"
        ws['A19'].fill = blue_fill
        ws['A19'].font = white_font
        ws['A19'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A19', 'K19')

        ws['L19'] = ""
        ws['L19'].border = border

        # Packaging procedure steps (rows 20-29) - WITH MERGED CELLS
        for i in range(1, 11):
            row = 19 + i
            ws[f'A{row}'] = str(i)
            ws[f'A{row}'].border = border
            ws[f'A{row}'].alignment = center_alignment
            ws[f'A{row}'].font = regular_font

            # MERGE CELLS B to J for each procedure step
            ws.merge_cells(f'B{row}:J{row}')
            ws[f'B{row}'] = ""
            ws[f'B{row}'].alignment = left_alignment
            self.apply_border_to_range(ws, f'B{row}', f'J{row}')

            # K and L columns
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Reference Images/Pictures section
        ws.merge_cells('A30:K30')
        ws['A30'] = "Reference Images/Pictures"
        ws['A30'].fill = blue_fill
        ws['A30'].font = white_font
        ws['A30'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A30', 'K30')

        ws['L30'] = ""
        ws['L30'].border = border

        # Image section headers
        ws.merge_cells('A31:C31')
        ws['A31'] = "Primary Packaging"
        ws['A31'].alignment = center_alignment
        ws['A31'].font = regular_font
        self.apply_border_to_range(ws, 'A31', 'C31')

        ws.merge_cells('D31:G31')
        ws['D31'] = "Secondary Packaging"
        ws['D31'].alignment = center_alignment
        ws['D31'].font = regular_font
        self.apply_border_to_range(ws, 'D31', 'G31')

        ws.merge_cells('H31:J31')
        ws['H31'] = "Label"
        ws['H31'].alignment = center_alignment
        ws['H31'].font = regular_font
        self.apply_border_to_range(ws, 'H31', 'J31')

        ws['K31'] = ""
        ws['K31'].border = border
        ws['L31'] = ""
        ws['L31'].border = border

        # Image placeholder areas (rows 32-37)
        ws.merge_cells('A32:C37')
        ws['A32'] = "Primary\nPackaging"
        ws['A32'].alignment = center_alignment
        ws['A32'].font = regular_font
        self.apply_border_to_range(ws, 'A32', 'C37')

        # Arrow 1
        ws['D35'] = "→"
        ws['D35'].border = border
        ws['D35'].alignment = center_alignment
        ws['D35'].font = Font(size=20, bold=True)

        # Secondary Packaging image area
        ws.merge_cells('E32:F37')
        ws['E32'] = "SECONDARY\nPACKAGING"
        ws['E32'].alignment = center_alignment
        ws['E32'].font = regular_font
        ws['E32'].fill = light_blue_fill
        self.apply_border_to_range(ws, 'E32', 'F37')

        # Arrow 2
        ws['G35'] = "→"
        ws['G35'].border = border
        ws['G35'].alignment = center_alignment
        ws['G35'].font = Font(size=20, bold=True)

        # Label image area
        ws.merge_cells('H32:K37')
        ws['H32'] = "LABEL"
        ws['H32'].alignment = center_alignment
        ws['H32'].font = regular_font
        self.apply_border_to_range(ws, 'H32', 'K37')

        # Add borders to remaining cells in image section
        for row in range(32, 38):
            for col in ['D', 'G', 'L']:
                if row != 35 or col != 'D':  # Skip D35 and G35 which have arrows
                    if row != 35 or col != 'G':
                        ws[f'{col}{row}'] = ""
                        ws[f'{col}{row}'].border = border

        # Approval Section
        ws.merge_cells('A38:C38')
        ws['A38'] = "Issued By"
        ws['A38'].alignment = center_alignment
        ws['A38'].font = regular_font
        self.apply_border_to_range(ws, 'A38', 'C38')

        ws.merge_cells('D38:G38')
        ws['D38'] = "Reviewed By"
        ws['D38'].alignment = center_alignment
        ws['D38'].font = regular_font
        self.apply_border_to_range(ws, 'D38', 'G38')

        ws.merge_cells('H38:K38')
        ws['H38'] = "Approved By"
        ws['H38'].alignment = center_alignment
        ws['H38'].font = regular_font
        self.apply_border_to_range(ws, 'H38', 'K38')

        ws['L38'] = ""
        ws['L38'].border = border

        # Signature boxes (rows 39-42)
        ws.merge_cells('A39:C42')
        ws['A39'] = ""
        self.apply_border_to_range(ws, 'A39', 'C42')

        ws.merge_cells('D39:G42')
        ws['D39'] = ""
        self.apply_border_to_range(ws, 'D39', 'G42')

        ws.merge_cells('H39:K42')
        ws['H39'] = ""
        self.apply_border_to_range(ws, 'H39', 'K42')

        # Apply borders for L column in signature section
        for row in range(39, 43):
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Second Approval Section
        ws.merge_cells('A43:C43')
        ws['A43'] = "Issued By"
        ws['A43'].alignment = center_alignment
        ws['A43'].font = regular_font
        self.apply_border_to_range(ws, 'A43', 'C43')

        ws.merge_cells('D43:G43')
        ws['D43'] = "Reviewed By"
        ws['D43'].alignment = center_alignment
        ws['D43'].font = regular_font
        self.apply_border_to_range(ws, 'D43', 'G43')

        ws.merge_cells('H43:J43')
        ws['H43'] = "Approved By"
        ws['H43'].alignment = center_alignment
        ws['H43'].font = regular_font
        self.apply_border_to_range(ws, 'H43', 'J43')

        ws['K43'] = ""
        ws['K43'].border = border
        ws['L43'] = ""
        ws['L43'].border = border

        # Second signature boxes (rows 44-47)
        ws.merge_cells('A44:C47')
        ws['A44'] = ""
        self.apply_border_to_range(ws, 'A44', 'C47')

        ws.merge_cells('D44:G47')
        ws['D44'] = ""
        self.apply_border_to_range(ws, 'D44', 'G47')

        ws.merge_cells('H44:J47')
        ws['H44'] = ""
        self.apply_border_to_range(ws, 'H44', 'J47')

        # Apply borders for K and L columns in second signature section
        for row in range(44, 48):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Final rows (48-50) - empty with borders
        for row in range(48, 51):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'] = ""
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
        
        # Mapping of data keys to cell positions (updated for exact template)
        cell_mapping = {
            'Revision No.': 'B2',
            'Date': 'G2',
            'Vendor Code': 'B5',
            'Vendor Name': 'B6',
            'Vendor Location': 'B7',
            'Part No.': 'G5',
            'Part Description': 'G6',
            'Part Unit Weight': 'G7',
            'Primary Packaging Type': 'A11',
            'Primary L-mm': 'B11',
            'Primary W-mm': 'C11',
            'Primary H-mm': 'D11',
            'Primary Qty/Pack': 'E11',
            'Primary Empty Weight': 'F11',
            'Primary Pack Weight': 'G11',
            'Secondary Packaging Type': 'A16',
            'Secondary L-mm': 'B16',
            'Secondary W-mm': 'C16',
            'Secondary H-mm': 'D16',
            'Secondary Qty/Pack': 'E16',
            'Secondary Empty Weight': 'F16',
            'Secondary Pack Weight': 'G16',
        }
        
        # Fill procedure steps in merged cells
        for i in range(1, 11):
            key = f'Procedure Step {i}'
            if key in data_dict and data_dict[key]:
                row = 19 + i
                # Only set value on the top-left cell of merged range
                ws[f'B{row}'].value = data_dict[key]
        
        # Fill other mapped cells
        for key, cell_pos in cell_mapping.items():
            if key in data_dict and data_dict[key]:
                # Only set value if it's not a merged cell or if it's the top-left cell
                cell = ws[cell_pos]
                if not isinstance(cell, MergedCell):
                    cell.value = data_dict[key]
                    
        return template_wb
    
    def add_current_packaging_images(self, template_wb, images_dict):
        """Add current packaging images to the template"""
        ws = template_wb.active
        
        try:
            # Add primary packaging image
            if 'current_primary' in images_dict:
                img_data = images_dict['current_primary']
                img = Image(io.BytesIO(img_data))
                img.width = 150
                img.height = 100
                ws.add_image(img, 'L5')
            
            # Add secondary packaging image  
            if 'current_secondary' in images_dict:
                img_data = images_dict['current_secondary']
                img = Image(io.BytesIO(img_data))
                img.width = 150
                img.height = 100
                ws.add_image(img, 'L10')
            
            # Add label image
            if 'current_label' in images_dict:
                img_data = images_dict['current_label']
                img = Image(io.BytesIO(img_data))
                img.width = 150
                img.height = 100
                ws.add_image(img, 'L16')
                
        except Exception as e:
            st.warning(f"Could not add images to template: {str(e)}")
        
        return template_wb
def main():
    st.set_page_config(page_title="Exact Packaging Instruction Template", layout="wide")
    
    st.title("📦 Exact Packaging Instruction Template Manager")
    st.markdown("Create and fill the exact packaging instruction template with proper merged cells and current packaging image upload.")
    
    # Initialize the template manager
    template_manager = ExactPackagingTemplateManager()
    
    # Sidebar for admin controls
    with st.sidebar:
        st.header("Template Controls")
        
        # Download empty template
        if st.button("📥 Download Empty Template"):
            wb = template_manager.create_exact_template_excel()
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="Download Exact Template",
                data=output.getvalue(),
                file_name="exact_packaging_instruction_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ Empty template ready for download!")
    
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📋 Data Input")
    
    # File upload for data
    uploaded_file = st.file_uploader(
        "Upload CSV/Excel file with packaging data",
        type=['csv', 'xlsx', 'xls'],
        help="Upload a file containing the packaging instruction data"
    )
    # Manual data entry option
    with st.expander("🖊️ Manual Data Entry", expanded=not uploaded_file):
        # Create form for manual entry
        with st.form("manual_data_form"):
            st.markdown("### Header Information")
            revision_no = st.text_input("Revision No.", value="Revision 1")
            date = st.date_input("Date", value=pd.Timestamp.now().date())
                
            st.markdown("### Vendor Information")
            vendor_code = st.text_input("Vendor Code")
            vendor_name = st.text_input("Vendor Name")
            vendor_location = st.text_input("Vendor Location")
                
            st.markdown("### Part Information")
            part_no = st.text_input("Part No.")
            part_description = st.text_input("Part Description")
            part_unit_weight = st.text_input("Part Unit Weight")
            part_weight_unit = st.selectbox("Weight Unit", ["grams", "kg", "lbs", "oz"])
                
            st.markdown("### Primary Packaging")
            primary_type = st.text_input("Primary Packaging Type")
            col1_p, col2_p, col3_p = st.columns(3)
            with col1_p:
                primary_l = st.text_input("Length (mm)")
            with col2_p:
                primary_w = st.text_input("Width (mm)")
            with col3_p:
                primary_h = st.text_input("Height (mm)")
                
            primary_qty = st.text_input("Quantity per Pack")
            primary_empty_weight = st.text_input("Empty Weight")
            primary_pack_weight = st.text_input("Pack Weight")
                
            st.markdown("### Secondary Packaging")
            secondary_type = st.text_input("Secondary Packaging Type")
            col1_s, col2_s, col3_s = st.columns(3)
            with col1_s:
                secondary_l = st.text_input("Length (mm)", key="sec_l")
            with col2_s:
                secondary_w = st.text_input("Width (mm)", key="sec_w")
            with col3_s:
                secondary_h = st.text_input("Height (mm)", key="sec_h")
                
            secondary_qty = st.text_input("Quantity per Pack", key="sec_qty")
            secondary_empty_weight = st.text_input("Empty Weight", key="sec_empty")
            secondary_pack_weight = st.text_input("Pack Weight", key="sec_pack")
                
            st.markdown("### Packaging Procedures")
            procedure_steps = []
            for i in range(1, 11):
                step = st.text_area(f"Step {i}", key=f"step_{i}")
                procedure_steps.append(step)
                
            st.markdown("### Approval")
            issued_by = st.text_input("Issued By")
            reviewed_by = st.text_input("Reviewed By")
            approved_by = st.text_input("Approved By")
                
            submitted = st.form_submit_button("Save Manual Data")
                
            if submitted:
                # Create data dictionary from manual input
                manual_data = {
                    'Revision No.': revision_no,
                    'Date': str(date),
                    'Vendor Code': vendor_code,
                    'Vendor Name': vendor_name,
                    'Vendor Location': vendor_location,
                    'Part No.': part_no,
                    'Part Description': part_description,
                    'Part Unit Weight': f"{part_unit_weight} {part_weight_unit}",
                    'Primary Packaging Type': primary_type,
                    'Primary L-mm': primary_l,
                    'Primary W-mm': primary_w,
                    'Primary H-mm': primary_h,
                    'Primary Qty/Pack': primary_qty,
                    'Primary Empty Weight': primary_empty_weight,
                    'Primary Pack Weight': primary_pack_weight,
                    'Secondary Packaging Type': secondary_type,
                    'Secondary L-mm': secondary_l,
                    'Secondary W-mm': secondary_w,
                    'Secondary H-mm': secondary_h,
                    'Secondary Qty/Pack': secondary_qty,
                    'Secondary Empty Weight': secondary_empty_weight,
                    'Secondary Pack Weight': secondary_pack_weight,
                    'Issued By': issued_by,
                    'Reviewed By': reviewed_by,
                    'Approved By': approved_by
                }
                    
                # Add procedure steps
                for i, step in enumerate(procedure_steps, 1):
                    manual_data[f'Procedure Step {i}'] = step
                    
                st.session_state.manual_data = manual_data
                st.success("✅ Manual data saved successfully!")
    
    with col2:
        st.subheader("🖼️ Current Packaging Images")
        
        # Image upload section
        st.markdown("Upload current packaging images:")
        
        col_img1, col_img2, col_img3 = st.columns(3)
        
        with col_img1:
            st.markdown("**Primary Packaging**")
            primary_img = st.file_uploader(
                "Primary Image",
                type=['png', 'jpg', 'jpeg'],
                key="primary_img"
            )
            if primary_img:
                st.image(primary_img, caption="Primary Packaging", use_column_width=True)
        
        with col_img2:
            st.markdown("**Secondary Packaging**")
            secondary_img = st.file_uploader(
                "Secondary Image",
                type=['png', 'jpg', 'jpeg'],
                key="secondary_img"
            )
            if secondary_img:
                st.image(secondary_img, caption="Secondary Packaging", use_column_width=True)
        
        with col_img3:
            st.markdown("**Label**")
            label_img = st.file_uploader(
                "Label Image",
                type=['png', 'jpg', 'jpeg'],
                key="label_img"
            )
            if label_img:
                st.image(label_img, caption="Label", use_column_width=True)
    
    # Processing section
    st.markdown("---")
    st.subheader("🔧 Generate Filled Template")
    
    if st.button("📄 Generate Filled Template", type="primary"):
        try:
            # Determine data source
            data_dict = {}
            
            if uploaded_file:
                # Extract data from uploaded file
                data_dict, df = template_manager.extract_data_from_uploaded_file(uploaded_file)
                st.success(f"✅ Data extracted from {uploaded_file.name}")
                
                # Display extracted data
                with st.expander("📊 View Extracted Data"):
                    st.dataframe(df)
                    
            elif hasattr(st.session_state, 'manual_data'):
                # Use manual data
                data_dict = st.session_state.manual_data
                st.success("✅ Using manual data entry")
            else:
                st.warning("⚠️ Please upload a file or enter manual data first")
                st.stop()
            
            # Create template
            wb = template_manager.create_exact_template_excel()
            
            # Fill template with data
            filled_wb = template_manager.fill_template_with_data(wb, data_dict)
            
            # Prepare images dictionary
            images_dict = {}
            if primary_img:
                images_dict['current_primary'] = primary_img.getvalue()
            if secondary_img:
                images_dict['current_secondary'] = secondary_img.getvalue()
            if label_img:
                images_dict['current_label'] = label_img.getvalue()
            
            # Add images to template
            if images_dict:
                filled_wb = template_manager.add_current_packaging_images(filled_wb, images_dict)
                st.success(f"✅ Added {len(images_dict)} image(s) to template")
            
            # Save filled template
            output = io.BytesIO()
            filled_wb.save(output)
            output.seek(0)
            
            # Generate filename
            part_no = data_dict.get('Part No.', 'Unknown')
            filename = f"packaging_instruction_{part_no.replace('/', '_')}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            st.download_button(
                label="📥 Download Filled Template",
                data=output.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success("🎉 Template generated successfully!")
            
            # Show summary
            with st.expander("📋 Template Summary"):
                st.write("**Data Fields Filled:**")
                filled_fields = [k for k, v in data_dict.items() if v]
                for field in filled_fields:
                    st.write(f"- {field}: {data_dict[field]}")
                
                st.write(f"\n**Images Added:** {len(images_dict)}")
                st.write(f"**Template Status:** Complete")
                
        except Exception as e:
            st.error(f"❌ Error generating template: {str(e)}")
            st.exception(e)
    
    # Footer
    st.markdown("---")
    st.markdown("### 📝 Instructions")
    st.markdown("""
    1. **Upload Data**: Use CSV/Excel file with packaging data or enter manually
    2. **Add Images**: Upload current packaging images (Primary, Secondary, Label)
    3. **Generate**: Click 'Generate Filled Template' to create the filled Excel file
    4. **Download**: Download the completed packaging instruction template
    
    **File Format Requirements:**
    - CSV/Excel files should contain columns matching the template fields
    - Images should be in PNG, JPG, or JPEG format
    - All fields are optional but recommended for complete documentation
    """)

if __name__ == "__main__":
    main()
