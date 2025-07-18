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
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
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

        # Other columns (G2 to J2) left blank but styled
        for col in ['G2', 'H2', 'I2', 'J2']:
            ws[col] = ""
            ws[col].border = border
        ws['K2'] = ""
        ws['K2'].border = border
        ws['L2'] = ""
        ws['L2'].border = border

        ws.merge_cells('B3:C3')
        ws['B3'] = ""
        ws['B3'].border = border
        ws['C3'].border = border

        # Empty cells in row 3
        for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}3'] = ""
            ws[f'{col}3'].border = border

        # Empty row 4 with borders
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col}4'] = ""
            ws[f'{col}4'].border = border
            
        # Vendor Information Title
        ws.merge_cells('A4:D4')
        ws['A4'] = "Vendor Information"
        ws['A4'].font = title_font  # Define this font earlier (e.g., bold/size)
        ws['A4'].alignment = center_alignment  # Define as centered horizontally
        ws['A4'].border = border

        # Part Information Title
        ws.merge_cells('F4:I4')
        ws['F4'] = "Part Information"
        ws['F4'].font = title_font
        ws['F4'].alignment = center_alignment
        ws['F4'].border = border

        # Vendor Code Row
        ws['A5'] = "Code"
        ws['A5'].font = header_font  # Define header_font if needed
        ws['A5'].alignment = left_alignment
        ws['A5'].border = border

        ws.merge_cells('B5:D5')
        ws['B5'] = ""  # Placeholder for value
        ws['B5'].border = border
        
        # Part fields
        ws['F5'] = "Part No."
        ws['F5'].border = border
        ws['F5'].alignment = left_alignment
        ws['F5'].font = regular_font

        ws.merge_cells('G5:K5')
        ws['G5'] = ""
        ws['G5'].border = border
        for col in ['H', 'I', 'J','K']:
            ws[f'{col}5'].border = border
        # Vendor Name Row
        ws['A6'] = "Name"
        ws['A6'].font = header_font
        ws['A6'].alignment = left_alignment
        ws['A6'].border = border

        ws.merge_cells('B6:D6')
        ws['B6'] = ""  # Placeholder for value
        ws['B6'].border = border

        ws['F6'] = "Description"
        ws['F6'].border = border
        ws['F6'].alignment = left_alignment
        ws['F6'].font = regular_font

        ws.merge_cells('G6:K6')
        ws['G6'] = ""
        ws['G6'].border = border
        for col in ['H', 'I', 'J','K']:
            ws[f'{col}6'].border = border
        # Vendor Location Row
        ws['A7'] = "Location"
        ws['A7'].font = header_font
        ws['A7'].alignment = left_alignment
        ws['A7'].border = border

        ws.merge_cells('B7:D7')
        ws['B7'] = ""  # Placeholder for value
        ws['B7'].border = border

        ws['F7'] = "Unit Weight"
        ws['F7'].border = border
        ws['F7'].alignment = left_alignment
        ws['F7'].font = regular_font

        ws.merge_cells('G7:K7')
        ws['G7'] = ""
        ws['G7'].border = border
        for col in ['H', 'I', 'J','K']:
            ws[f'{col}6'].border = border

        # Current packaging section rows 5-7
        for row in range(5, 8):
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Additional row after Unit Weight (Row 8) for L and W only
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

        ws['K8'] = ""
        ws['K8'].border = border

        # Also apply border to L8 (current packaging)
        ws['L8'] = ""
        ws['L8'].border = border

        # Title row after spacing
        ws.merge_cells('A9:K9')
        ws['A9'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A9'].fill = blue_fill
        ws['A9'].font = white_font
        ws['A9'].border = border
        ws['A9'].alignment = center_alignment

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
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}10'] = ""
            ws[f'{col}10'].border = border
        ws['L10'] = ""
        ws['L10'].border = border

        # Primary packaging data rows (11-13)
        for row in range(11, 14):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border
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
        ws['A14'].border = border
        ws['A14'].alignment = center_alignment

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
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border
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
        ws['A19'].border = border
        ws['A19'].alignment = center_alignment

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
        # Reference Images/Pictures section
        ws.merge_cells('A30:K30')
        ws['A30'] = "Reference Images/Pictures"
        ws['A30'].fill = blue_fill
        ws['A30'].font = white_font
        ws['A30'].border = border
        ws['A30'].alignment = center_alignment

        ws['L30'] = ""
        ws['L30'].border = border

        # Image section headers
        ws.merge_cells('A31:C31')
        ws['A31'] = "Primary Packaging"
        ws['A31'].border = border
        ws['A31'].alignment = center_alignment
        ws['A31'].font = regular_font

        ws.merge_cells('D31:G31')
        ws['D31'] = "Secondary Packaging"
        ws['D31'].border = border
        ws['D31'].alignment = center_alignment
        ws['D31'].font = regular_font

        ws.merge_cells('H31:J31')
        ws['H31'] = "Label"
        ws['H31'].border = border
        ws['H31'].alignment = center_alignment
        ws['H31'].font = regular_font

        ws['K31'] = ""
        ws['K31'].border = border
        ws['L31'] = ""
        ws['L31'].border = border

        # Image placeholder areas (rows 32-37)
        ws.merge_cells('A32:C37')
        ws['A32'] = "Primary\nPackaging"
        ws['A32'].border = border
        ws['A32'].alignment = center_alignment
        ws['A32'].font = regular_font

        # Arrow 1
        ws['D35'] = "→"
        ws['D35'].border = border
        ws['D35'].alignment = center_alignment
        ws['D35'].font = Font(size=20, bold=True)

        # Secondary Packaging image area
        ws.merge_cells('E32:F37')
        ws['E32'] = "SECONDARY\nPACKAGING"
        ws['E32'].border = border
        ws['E32'].alignment = center_alignment
        ws['E32'].font = regular_font
        ws['E32'].fill = light_blue_fill

        # Arrow 2
        ws['G35'] = "→"
        ws['G35'].border = border
        ws['G35'].alignment = center_alignment
        ws['G35'].font = Font(size=20, bold=True)

        # Label image area
        ws.merge_cells('H32:K37')
        ws['H32'] = "LABEL"
        ws['H32'].border = border
        ws['H32'].alignment = center_alignment
        ws['H32'].font = regular_font

        # Add borders to all cells in image area
        for row in range(32, 38):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'].border = border
        # Approval Section
        ws.merge_cells('A38:C38')
        ws['A38'] = "Issued By"
        ws['A38'].border = border
        ws['A38'].alignment = center_alignment
        ws['A38'].font = regular_font

        ws.merge_cells('D38:G38')
        ws['D38'] = "Reviewed By"
        ws['D38'].border = border
        ws['D38'].alignment = center_alignment
        ws['D38'].font = regular_font

        ws.merge_cells('H38:K38')
        ws['H38'] = "Approved By"
        ws['H38'].border = border
        ws['H38'].alignment = center_alignment
        ws['H38'].font = regular_font

        ws['K38'] = ""
        ws['K38'].border = border
        ws['L38'] = ""
        ws['L38'].border = border

        # Signature boxes (rows 39-42)
        ws.merge_cells('A39:C42')
        ws['A39'] = ""
        ws['A39'].border = border

        ws.merge_cells('D39:G42')
        ws['D39'] = ""
        ws['D39'].border = border

        ws.merge_cells('H39:K42')
        ws['H39'] = ""
        ws['H39'].border = border

        # Apply borders for signature section
        for row in range(39, 43):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'].border = border
        # Second Approval Section
        ws.merge_cells('A43:C43')
        ws['A43'] = "Issued By"
        ws['A43'].border = border
        ws['A43'].alignment = center_alignment
        ws['A43'].font = regular_font

        ws.merge_cells('D43:G43')
        ws['D43'] = "Reviewed By"
        ws['D43'].border = border
        ws['D43'].alignment = center_alignment
        ws['D43'].font = regular_font

        ws.merge_cells('H43:J43')
        ws['H43'] = "Approved By"
        ws['H43'].border = border
        ws['H43'].alignment = center_alignment
        ws['H43'].font = regular_font

        ws['K43'] = ""
        ws['K43'].border = border
        ws['L43'] = ""
        ws['L43'].border = border

        # Second signature boxes (rows 44-47)
        ws.merge_cells('A44:C47')
        ws['A44'] = ""
        ws['A44'].border = border

        ws.merge_cells('D44:G47')
        ws['D44'] = ""
        ws['D44'].border = border

        ws.merge_cells('H44:J47')
        ws['H44'] = ""
        ws['H44'].border = border

        # Apply borders for second signature section
        for row in range(44, 48):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'].border = border
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
            'Date': 'E2',
            'QC': 'H2',
            'MM': 'J2',
            'VP': 'B3',
            'Vendor Code': 'B5',
            'Vendor Name': 'B6',
            'Vendor Location': 'B7',
            'Part No.': 'G5',
            'Part Description': 'G6',
            'Part Unit Weight': 'G7',
            'Primary Packaging Type': 'A10',
            'Primary L-mm': 'B10',
            'Primary W-mm': 'C10',
            'Primary H-mm': 'D10',
            'Primary Qty/Pack': 'E10',
            'Primary Empty Weight': 'F10',
            'Primary Pack Weight': 'G10',
            'Secondary Packaging Type': 'A15',
            'Secondary L-mm': 'B15',
            'Secondary W-mm': 'C15',
            'Secondary H-mm': 'D15',
            'Secondary Qty/Pack': 'E15',
            'Secondary Empty Weight': 'F15',
            'Secondary Pack Weight': 'G15',
            'Issued By': 'A38',
            'Reviewed By': 'D38',
            'Approved By': 'H38'
        }
        
        # Fill procedure steps in merged cells
        for i in range(1, 11):
            key = f'Procedure Step {i}'
            if key in data_dict and data_dict[key]:
                ws[f'B{18+i}'] = data_dict[key]
        
        # Fill other mapped cells
        for key, cell_pos in cell_mapping.items():
            if key in data_dict and data_dict[key]:
                ws[cell_pos] = data_dict[key]
        
        return template_wb
    
    def add_current_packaging_images(self, template_wb, images_dict):
        """Add current packaging images to the template"""
        ws = template_wb.active
        
        try:
            # Add primary packaging image
            if 'current_primary' in images_dict:
                img_data = images_dict['current_primary']
                img = Image(io.BytesIO(img_data))
                img.width = 100
                img.height = 80
                ws.add_image(img, 'K30')
            
            # Add secondary packaging image  
            if 'current_secondary' in images_dict:
                img_data = images_dict['current_secondary']
                img = Image(io.BytesIO(img_data))
                img.width = 100
                img.height = 80
                ws.add_image(img, 'K32')
            
            # Add label image
            if 'current_label' in images_dict:
                img_data = images_dict['current_label']
                img = Image(io.BytesIO(img_data))
                img.width = 100
                img.height = 80
                ws.add_image(img, 'K34')
                
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
        st.subheader("📋 Data Input & Image Management")
        
        # Option to upload CSV/Excel with data
        uploaded_file = st.file_uploader(
            "Upload data file (CSV/Excel) with optional images", 
            type=['csv', 'xlsx', 'xls']
        )
        
        if uploaded_file is not None:
            try:
                # Extract data
                data_dict, df = template_manager.extract_data_from_uploaded_file(uploaded_file)
                
                st.success("✅ Data file uploaded successfully!")
                
                # Show extracted data
                with st.expander("📊 View Extracted Data"):
                    st.dataframe(df)
                
                # Extract images if available
                extracted_images = template_manager.extract_images_from_file(uploaded_file)
                if extracted_images:
                    st.success(f"✅ Found {len(extracted_images)} images in the file!")
                    
                    # Show extracted images
                    with st.expander("🖼️ View Extracted Images"):
                        for img_name, img_bytes in extracted_images.items():
                            st.image(img_bytes, caption=img_name, width=200)
                else:
                    st.info("ℹ️ No images found in the uploaded file.")
                
                # Generate filled template
                if st.button("🔄 Generate Filled Template"):
                    wb = template_manager.create_exact_template_excel()
                    filled_wb = template_manager.fill_template_with_data(wb, data_dict)
                    
                    output = io.BytesIO()
                    filled_wb.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="📤 Download Filled Template",
                        data=output.getvalue(),
                        file_name="filled_packaging_instruction_template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("✅ Filled template ready for download!")
            
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
    
    with col2:
        st.subheader("🖼️ Current Packaging Images")
        
        # Current packaging images upload
        st.markdown("Upload current packaging images to add to the template:")
        
        current_primary = st.file_uploader(
            "Current Primary Packaging Image", 
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            key="current_primary"
        )
        
        current_secondary = st.file_uploader(
            "Current Secondary Packaging Image", 
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            key="current_secondary"
        )
        
        current_label = st.file_uploader(
            "Current Label Image", 
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            key="current_label"
        )
        
        # Preview uploaded images
        images_dict = {}
        
        if current_primary:
            st.image(current_primary, caption="Current Primary Packaging", width=200)
            images_dict['current_primary'] = current_primary.getvalue()
        
        if current_secondary:
            st.image(current_secondary, caption="Current Secondary Packaging", width=200)
            images_dict['current_secondary'] = current_secondary.getvalue()
        
        if current_label:
            st.image(current_label, caption="Current Label", width=200)
            images_dict['current_label'] = current_label.getvalue()
        
        # Generate template with images
        if images_dict and st.button("🔄 Generate Template with Images"):
            wb = template_manager.create_exact_template_excel()
            wb_with_images = template_manager.add_current_packaging_images(wb, images_dict)
            
            output = io.BytesIO()
            wb_with_images.save(output)
            output.seek(0)
            
            st.download_button(
                label="📤 Download Template with Images",
                data=output.getvalue(),
                file_name="packaging_instruction_with_images.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ Template with images ready for download!")
    
    # Manual data entry section
    st.subheader("✍️ Manual Data Entry")
    
    with st.expander("📝 Enter Data Manually"):
        # Create tabs for different sections
        tab1, tab2, tab3, tab4 = st.tabs(["Header Info", "Vendor & Part Info", "Packaging Info", "Procedure Steps"])
        
        manual_data = {}
        
        with tab1:
            st.subheader("Header Information")
            col1, col2 = st.columns(2)
            
            with col1:
                manual_data['Revision No.'] = st.text_input("Revision No.")
                manual_data['Date'] = st.text_input("Date")
                manual_data['QC'] = st.text_input("QC")
            
            with col2:
                manual_data['MM'] = st.text_input("MM")
                manual_data['VP'] = st.text_input("VP")
        
        with tab2:
            st.subheader("Vendor Information")
            col1, col2 = st.columns(2)
            
            with col1:
                manual_data['Vendor Code'] = st.text_input("Vendor Code")
                manual_data['Vendor Name'] = st.text_input("Vendor Name")
                manual_data['Vendor Location'] = st.text_input("Vendor Location")
            
            with col2:
                st.subheader("Part Information")
                manual_data['Part No.'] = st.text_input("Part No.")
                manual_data['Part Description'] = st.text_area("Part Description", height=100)
                manual_data['Part Unit Weight'] = st.text_input("Part Unit Weight")
        
        with tab3:
            st.subheader("Primary Packaging")
            col1, col2 = st.columns(2)
            
            with col1:
                manual_data['Primary Packaging Type'] = st.text_input("Primary Packaging Type")
                manual_data['Primary L-mm'] = st.text_input("Primary L-mm")
                manual_data['Primary W-mm'] = st.text_input("Primary W-mm")
                manual_data['Primary H-mm'] = st.text_input("Primary H-mm")
            
            with col2:
                manual_data['Primary Qty/Pack'] = st.text_input("Primary Qty/Pack")
                manual_data['Primary Empty Weight'] = st.text_input("Primary Empty Weight")
                manual_data['Primary Pack Weight'] = st.text_input("Primary Pack Weight")
            
            st.subheader("Secondary Packaging")
            col1, col2 = st.columns(2)
            
            with col1:
                manual_data['Secondary Packaging Type'] = st.text_input("Secondary Packaging Type")
                manual_data['Secondary L-mm'] = st.text_input("Secondary L-mm")
                manual_data['Secondary W-mm'] = st.text_input("Secondary W-mm")
                manual_data['Secondary H-mm'] = st.text_input("Secondary H-mm")
            
            with col2:
                manual_data['Secondary Qty/Pack'] = st.text_input("Secondary Qty/Pack")
                manual_data['Secondary Empty Weight'] = st.text_input("Secondary Empty Weight")
                manual_data['Secondary Pack Weight'] = st.text_input("Secondary Pack Weight")
        
        with tab4:
            st.subheader("Packaging Procedure Steps")
            col1, col2 = st.columns(2)
            
            with col1:
                for i in range(1, 6):
                    manual_data[f'Procedure Step {i}'] = st.text_area(f"Step {i}", height=80, key=f"step_{i}")
            
            with col2:
                for i in range(6, 11):
                    manual_data[f'Procedure Step {i}'] = st.text_area(f"Step {i}", height=80, key=f"step_{i}")
            
            st.subheader("Approval")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                manual_data['Issued By'] = st.text_input("Issued By")
            
            with col2:
                manual_data['Reviewed By'] = st.text_input("Reviewed By")
            
            with col3:
                manual_data['Approved By'] = st.text_input("Approved By")
        
        # Generate template with manual data
        if st.button("🔄 Generate Template with Manual Data"):
            wb = template_manager.create_exact_template_excel()
            filled_wb = template_manager.fill_template_with_data(wb, manual_data)
            
            # Add images if uploaded
            if images_dict:
                filled_wb = template_manager.add_current_packaging_images(filled_wb, images_dict)
            
            output = io.BytesIO()
            filled_wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="📤 Download Template with Manual Data",
                data=output.getvalue(),
                file_name="manual_packaging_instruction_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ Template with manual data ready for download!")
    
    # Footer
    st.markdown("---")
    st.markdown("💡 **Tips:**")
    st.markdown("- Upload a CSV/Excel file with data to auto-fill the template")
    st.markdown("- Upload current packaging images to add them to the template")
    st.markdown("- Use manual data entry for custom entries")
    st.markdown("- The template matches the exact format with proper merged cells")

if __name__ == "__main__":
    main()
