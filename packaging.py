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
        white_font = Font(color="FFFFFF", bold=True, size=10)
        black_font = Font(color="000000", bold=True, size=10)
        regular_font = Font(color="000000", size=9)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        # Set column widths to match the image exactly
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 8
        ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 8
        ws.column_dimensions['G'].width = 8
        ws.column_dimensions['H'].width = 8
        ws.column_dimensions['I'].width = 8
        ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 18
        
        # Set row heights
        for row in range(1, 50):
            ws.row_dimensions[row].height = 20
        
        # Header Row - "Packaging Instruction"
        ws.merge_cells('A1:J1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = blue_fill
        ws['A1'].font = white_font
        ws['A1'].border = border
        ws['A1'].alignment = center_alignment
        
        # Current Packaging header (right side)
        ws['K1'] = "CURRENT PACKAGING"
        ws['K1'].fill = blue_fill
        ws['K1'].font = white_font
        ws['K1'].border = border
        ws['K1'].alignment = center_alignment
        
        # Revision information row
        ws['A2'] = "Revision No."
        ws['A2'].border = border
        ws['A2'].alignment = left_alignment
        ws['A2'].font = regular_font
        
        ws.merge_cells('B2:C2')
        ws['B2'] = ""
        ws['B2'].border = border
        ws['C2'].border = border
        
        ws['D2'] = "Date"
        ws['D2'].border = border
        ws['D2'].alignment = left_alignment
        ws['D2'].font = regular_font
        
        ws.merge_cells('E2:F2')
        ws['E2'] = ""
        ws['E2'].border = border
        ws['F2'].border = border
        
        ws['G2'] = "QC"
        ws['G2'].border = border
        ws['G2'].alignment = left_alignment
        ws['G2'].font = regular_font
        
        ws['H2'] = ""
        ws['H2'].border = border
        
        ws['I2'] = "MM"
        ws['I2'].border = border
        ws['I2'].alignment = left_alignment
        ws['I2'].font = regular_font
        
        ws['J2'] = ""
        ws['J2'].border = border
        
        ws['K2'] = ""
        ws['K2'].border = border
        
        # VP field - separate row
        ws['A3'] = "VP"
        ws['A3'].border = border
        ws['A3'].alignment = left_alignment
        ws['A3'].font = regular_font
        
        ws.merge_cells('B3:C3')
        ws['B3'] = ""
        ws['B3'].border = border
        ws['C3'].border = border
        
        # Empty cells in row 3
        for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
            ws[f'{col}3'] = ""
            ws[f'{col}3'].border = border
        
        # Vendor Information section
        ws.merge_cells('A4:D4')
        ws['A4'] = "Vendor Information"
        ws['A4'].fill = blue_fill
        ws['A4'].font = white_font
        ws['A4'].border = border
        ws['A4'].alignment = center_alignment
        
        # Part Information section
        ws.merge_cells('F4:J4')
        ws['F4'] = "Part Information"
        ws['F4'].fill = blue_fill
        ws['F4'].font = white_font
        ws['F4'].border = border
        ws['F4'].alignment = center_alignment
        
        # Current packaging section continues
        ws['K4'] = ""
        ws['K4'].border = border
        
        # Vendor fields
        ws['A5'] = "Code"
        ws['A5'].border = border
        ws['A5'].alignment = left_alignment
        ws['A5'].font = regular_font
        
        ws.merge_cells('B5:D5')
        ws['B5'] = ""
        ws['B5'].border = border
        for col in ['C', 'D']:
            ws[f'{col}5'].border = border
        
        ws['A6'] = "Name"
        ws['A6'].border = border
        ws['A6'].alignment = left_alignment
        ws['A6'].font = regular_font
        
        ws.merge_cells('B6:D6')
        ws['B6'] = ""
        ws['B6'].border = border
        for col in ['C', 'D']:
            ws[f'{col}6'].border = border
        
        ws['A7'] = "Location"
        ws['A7'].border = border
        ws['A7'].alignment = left_alignment
        ws['A7'].font = regular_font
        
        ws.merge_cells('B7:D7')
        ws['B7'] = ""
        ws['B7'].border = border
        for col in ['C', 'D']:
            ws[f'{col}7'].border = border
        
        # Part fields
        ws['F5'] = "Part No."
        ws['F5'].border = border
        ws['F5'].alignment = left_alignment
        ws['F5'].font = regular_font
        
        ws.merge_cells('G5:J5')
        ws['G5'] = ""
        ws['G5'].border = border
        for col in ['H', 'I', 'J']:
            ws[f'{col}5'].border = border
        
        ws['F6'] = "Description"
        ws['F6'].border = border
        ws['F6'].alignment = left_alignment
        ws['F6'].font = regular_font
        
        ws.merge_cells('G6:J6')
        ws['G6'] = ""
        ws['G6'].border = border
        for col in ['H', 'I', 'J']:
            ws[f'{col}6'].border = border
        
        ws['F7'] = "Unit Weight"
        ws['F7'].border = border
        ws['F7'].alignment = left_alignment
        ws['F7'].font = regular_font
        
        ws['G7'] = ""
        ws['G7'].border = border
        
        ws['H7'] = "W"
        ws['H7'].border = border
        ws['H7'].alignment = center_alignment
        ws['H7'].font = regular_font
        
        ws['I7'] = ""
        ws['I7'].border = border
        
        ws['J7'] = "H"
        ws['J7'].border = border
        ws['J7'].alignment = center_alignment
        ws['J7'].font = regular_font
        
        # Current packaging section rows 5-7
        for row in range(5, 8):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Primary Packaging Instruction header
        ws.merge_cells('A8:J8')
        ws['A8'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A8'].fill = blue_fill
        ws['A8'].font = white_font
        ws['A8'].border = border
        ws['A8'].alignment = center_alignment
        
        # Current packaging label
        ws['K8'] = "CURRENT PACKAGING"
        ws['K8'].fill = blue_fill
        ws['K8'].font = white_font
        ws['K8'].border = border
        ws['K8'].alignment = center_alignment
        
        # Primary packaging headers
        headers = ["Packaging Type", "L-mm", "W-mm", "H-mm", "Qty/Pack", "Empty Weight", "Pack Weight"]
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}9'] = header
            ws[f'{col}9'].border = border
            ws[f'{col}9'].alignment = center_alignment
            ws[f'{col}9'].font = regular_font
        
        # Empty cells for remaining columns in row 9
        for col in ['H', 'I', 'J']:
            ws[f'{col}9'] = ""
            ws[f'{col}9'].border = border
        
        # Primary packaging data rows (10-12)
        for row in range(10, 13):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
        
        # TOTAL row
        ws['H12'] = "TOTAL"
        ws['H12'].border = border
        ws['H12'].font = black_font
        ws['H12'].alignment = center_alignment
        
        # Current packaging section rows 9-12
        for row in range(9, 13):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Secondary Packaging Instruction header
        ws.merge_cells('A13:J13')
        ws['A13'] = "Secondary Packaging Instruction (Outer / External)"
        ws['A13'].fill = blue_fill
        ws['A13'].font = white_font
        ws['A13'].border = border
        ws['A13'].alignment = center_alignment
        
        # Secondary packaging headers
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}14'] = header
            ws[f'{col}14'].border = border
            ws[f'{col}14'].alignment = center_alignment
            ws[f'{col}14'].font = regular_font
        
        # Empty cells for remaining columns in row 14
        for col in ['H', 'I', 'J']:
            ws[f'{col}14'] = ""
            ws[f'{col}14'].border = border
        
        # Secondary packaging data rows (15-17)
        for row in range(15, 18):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
        
        # TOTAL row for secondary
        ws['H17'] = "TOTAL"
        ws['H17'].border = border
        ws['H17'].font = black_font
        ws['H17'].alignment = center_alignment
        
        # Problem section (right side)
        ws['K13'] = "PROBLEM IF ANY:"
        ws['K13'].border = border
        ws['K13'].font = black_font
        ws['K13'].alignment = left_alignment
        
        for row in range(14, 18):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Red caution section
        ws['K18'] = "CAUTION: REVISED DESIGN"
        ws['K18'].fill = red_fill
        ws['K18'].font = white_font
        ws['K18'].border = border
        ws['K18'].alignment = center_alignment
        
        # Packaging Procedure section
        ws.merge_cells('A18:J18')
        ws['A18'] = "Packaging Procedure"
        ws['A18'].fill = blue_fill
        ws['A18'].font = white_font
        ws['A18'].border = border
        ws['A18'].alignment = center_alignment
        
        # Packaging procedure steps (rows 19-28)
        for i in range(1, 11):
            row = 18 + i
            ws[f'A{row}'] = str(i)
            ws[f'A{row}'].border = border
            ws[f'A{row}'].alignment = center_alignment
            ws[f'A{row}'].font = regular_font
            
            for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
        
        # Current packaging section for procedure
        for row in range(19, 29):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Reference Images/Pictures section
        ws.merge_cells('A29:J29')
        ws['A29'] = "Reference Images/Pictures"
        ws['A29'].fill = blue_fill
        ws['A29'].font = white_font
        ws['A29'].border = border
        ws['A29'].alignment = center_alignment
        
        # Image section headers
        ws.merge_cells('A30:C30')
        ws['A30'] = "Primary Packaging"
        ws['A30'].border = border
        ws['A30'].alignment = center_alignment
        ws['A30'].font = regular_font
        
        ws.merge_cells('D30:G30')
        ws['D30'] = "Secondary Packaging"
        ws['D30'].border = border
        ws['D30'].alignment = center_alignment
        ws['D30'].font = regular_font
        
        ws.merge_cells('H30:J30')
        ws['H30'] = "Label"
        ws['H30'].border = border
        ws['H30'].alignment = center_alignment
        ws['H30'].font = regular_font
        
        # Image placeholder areas
        # Primary Packaging image area
        ws.merge_cells('A31:C36')
        ws['A31'] = "Primary\nPackaging"
        ws['A31'].border = border
        ws['A31'].alignment = center_alignment
        ws['A31'].font = regular_font
        
        # Arrow 1
        ws['D34'] = "→"
        ws['D34'].border = border
        ws['D34'].alignment = center_alignment
        ws['D34'].font = Font(size=20, bold=True)
        
        # Secondary Packaging image area
        ws.merge_cells('E31:F36')
        ws['E31'] = "SECONDARY\nPACKAGING"
        ws['E31'].border = border
        ws['E31'].alignment = center_alignment
        ws['E31'].font = regular_font
        ws['E31'].fill = light_blue_fill
        
        # Arrow 2
        ws['G34'] = "→"
        ws['G34'].border = border
        ws['G34'].alignment = center_alignment
        ws['G34'].font = Font(size=20, bold=True)
        
        # Label image area
        ws.merge_cells('H31:J36')
        ws['H31'] = "LABEL"
        ws['H31'].border = border
        ws['H31'].alignment = center_alignment
        ws['H31'].font = regular_font
        
        # Current packaging section for images
        for row in range(29, 37):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Approval section headers
        ws.merge_cells('A37:C37')
        ws['A37'] = "Issued By"
        ws['A37'].border = border
        ws['A37'].alignment = center_alignment
        ws['A37'].font = regular_font
        
        ws.merge_cells('D37:G37')
        ws['D37'] = "Reviewed By"
        ws['D37'].border = border
        ws['D37'].alignment = center_alignment
        ws['D37'].font = regular_font
        
        ws.merge_cells('H37:J37')
        ws['H37'] = "Approved By"
        ws['H37'].border = border
        ws['H37'].alignment = center_alignment
        ws['H37'].font = regular_font
        
        # Signature boxes
        ws.merge_cells('A38:C41')
        ws['A38'] = ""
        ws['A38'].border = border
        
        ws.merge_cells('D38:G41')
        ws['D38'] = ""
        ws['D38'].border = border
        
        ws.merge_cells('H38:J41')
        ws['H38'] = ""
        ws['H38'].border = border
        
        # Add borders to all merged cells
        for row in range(38, 42):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'].border = border
        
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
        
        # Fill procedure steps
        for i in range(1, 11):
            key = f'Procedure Step {i}'
            if key in data_dict and data_dict[key]:
                ws[f'B{18+i}'] = data_dict[key]
        
        # Fill other mapped cells
        for key, cell_pos in cell_mapping.items():
            if key in data_dict and data_dict[key]:
                ws[cell_pos] = data_dict[key]
        
        return template_wb

def main():
    st.set_page_config(page_title="Exact Packaging Instruction Template", layout="wide")
    
    st.title("📦 Exact Packaging Instruction Template Manager")
    st.markdown("Create and fill the exact packaging instruction template with data and image extraction.")
    
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
        st.subheader("📋 Data Input & Image Extraction")
        
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
                        file_name="filled_exact_packaging_instruction.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Template filled successfully!")
                    
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
        
        # Manual image upload section
        st.subheader("🖼️ Manual Image Upload")
        
        col_img1, col_img2, col_img3 = st.columns(3)
        
        with col_img1:
            st.write("**Primary Packaging Image**")
            primary_image = st.file_uploader("Upload Primary", type=['png', 'jpg', 'jpeg'], key="primary")
            if primary_image:
                st.image(primary_image, caption="Primary Packaging", width=150)
        
        with col_img2:
            st.write("**Secondary Packaging Image**")
            secondary_image = st.file_uploader("Upload Secondary", type=['png', 'jpg', 'jpeg'], key="secondary")
            if secondary_image:
                st.image(secondary_image, caption="Secondary Packaging", width=150)
        
        with col_img3:
            st.write("**Label Image**")
            label_image = st.file_uploader("Upload Label", type=['png', 'jpg', 'jpeg'], key="label")
            if label_image:
                st.image(label_image, caption="Label", width=150)
        
        # Manual data entry form
        st.subheader("✏️ Manual Data Entry")
        with st.expander("Enter data manually"):
            manual_data = {}
            
            # Header Information
            st.write("**Header Information**")
            col_a, col_b = st.columns(2)
            with col_a:
                manual_data['Revision No.'] = st.text_input("Revision No.", key="rev_no")
                manual_data['QC'] = st.text_input("QC", key="qc")
                manual_data['VP'] = st.text_input("VP", key="vp")
            with col_b:
                manual_data['Date'] = st.text_input("Date", key="date")
                manual_data['MM'] = st.text_input("MM", key="mm")
            
            # Vendor Information
            st.write("**Vendor Information**")
            manual_data['Vendor Code'] = st.text_input("Vendor Code", key="vendor_code")
            manual_data['Vendor Name'] = st.text_input("Vendor Name", key="vendor_name")
            manual_data['Vendor Location'] = st.text_input("Vendor Location", key="vendor_location")
            
            # Part Information
            st.write("**Part Information**")
            manual_data['Part No.'] = st.text_input("Part No.", key="part_no")
            manual_data['Part Description'] = st.text_input("Part Description", key="part_desc")
            manual_data['Part Unit Weight'] = st.text_input("Part Unit Weight", key="part_weight")
            
            # Primary Packaging
            st.write("**Primary Packaging**")
            col_a, col_b = st.columns(2)
            with col_a:
                manual_data['Primary Packaging Type'] = st.text_input("Primary Packaging Type", key="primary_type")
                manual_data['Primary L-mm'] = st.text_input("Primary L-mm", key="primary_l")
                manual_data['Primary W-mm'] = st.text_input("Primary W-mm", key="primary_w")
            with col_b:
                manual_data['Primary H-mm'] = st.text_input("Primary H-mm", key="primary_h")
                manual_data['Primary Qty/Pack'] = st.text_input("Primary Qty/Pack", key="primary_qty")
                manual_data['Primary Empty Weight'] = st.text_input("Primary Empty Weight", key="primary_empty")
                manual_data['Primary Pack Weight'] = st.text_input("Primary Pack Weight", key="primary_pack")
            
            # Secondary Packaging
            st.write("**Secondary Packaging**")
            col_a, col_b = st.columns(2)
            with col_a:
                manual_data['Secondary Packaging Type'] = st.text_input("Secondary Packaging Type", key="secondary_type")
                manual_data['Secondary L-mm'] = st.text_input("Secondary L-mm", key="secondary_l")
                manual_data['Secondary W-mm'] = st.text_input("Secondary W-mm", key="secondary_w")
            with col_b:
                manual_data['Secondary H-mm'] = st.text_input("Secondary H-mm", key="secondary_h")
                manual_data['Secondary Qty/Pack'] = st.text_input("Secondary Qty/Pack", key="secondary_qty")
                manual_data['Secondary Empty Weight'] = st.text_input("Secondary Empty Weight", key="secondary_empty")
                manual_data['Secondary Pack Weight'] = st.text_input("Secondary Pack Weight", key="secondary_pack")
            
            # Packaging Procedures
            st.write("**Packaging Procedures (10 Steps)**")
            for i in range(1, 11):
                manual_data[f'Procedure Step {i}'] = st.text_area(f"Step {i}", key=f"step_{i}", height=50)
            
            # Approval
            st.write("**Approval**")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                manual_data['Issued By'] = st.text_input("Issued By", key="issued_by")
            with col_b:
                manual_data['Reviewed By'] = st.text_input("Reviewed By", key="reviewed_by")
            with col_c:
                manual_data['Approved By'] = st.text_input("Approved By", key="approved_by")
            
            # Generate template with manual data
            if st.button("🔄 Generate Template with Manual Data"):
                wb = template_manager.create_exact_template_excel()
                filled_wb = template_manager.fill_template_with_data(wb, manual_data)
                
                output = io.BytesIO()
                filled_wb.save(output)
                output.seek(0)
                
                st.download_button(
                    label="📤 Download Manually Filled Template",
                    data=output.getvalue(),
                    file_name="manual_filled_packaging_instruction.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("✅ Template created with manual data!")
    
    with col2:
        st.subheader("📋 Template Preview & Information")
        
        # Show template structure
        st.write("**Template Structure:**")
        st.info("""
        📊 **Exact Packaging Instruction Template includes:**
        
        **Header Section:**
        - Revision No., Date, QC, MM, VP fields
        
        **Information Sections:**
        - Vendor Information (Code, Name, Location)
        - Part Information (Part No., Description, Unit Weight)
        
        **Packaging Sections:**
        - Primary Packaging Instruction (Type, Dimensions, Weight)
        - Secondary Packaging Instruction (Type, Dimensions, Weight)
        
        **Procedure Section:**
        - 10 detailed packaging procedure steps
        
        **Visual Section:**
        - Reference images for Primary, Secondary, and Label
        - Visual flow with arrows
        
        **Approval Section:**
        - Issued By, Reviewed By, Approved By signature boxes
        
        **Additional Features:**
        - Current Packaging comparison column
        - Problem identification section
        - Caution alerts for revised designs
        """)
        
        # Show field mapping
        with st.expander("📋 Field Mapping Guide"):
            st.write("**CSV/Excel Column Headers should match exactly:**")
            
            field_categories = {
                "Header Fields": [
                    "Revision No.", "Date", "QC", "MM", "VP"
                ],
                "Vendor Fields": [
                    "Vendor Code", "Vendor Name", "Vendor Location"
                ],
                "Part Fields": [
                    "Part No.", "Part Description", "Part Unit Weight", "Part Weight Unit"
                ],
                "Primary Packaging Fields": [
                    "Primary Packaging Type", "Primary L-mm", "Primary W-mm", 
                    "Primary H-mm", "Primary Qty/Pack", "Primary Empty Weight", 
                    "Primary Pack Weight"
                ],
                "Secondary Packaging Fields": [
                    "Secondary Packaging Type", "Secondary L-mm", "Secondary W-mm", 
                    "Secondary H-mm", "Secondary Qty/Pack", "Secondary Empty Weight", 
                    "Secondary Pack Weight"
                ],
                "Procedure Fields": [
                    f"Procedure Step {i}" for i in range(1, 11)
                ],
                "Approval Fields": [
                    "Issued By", "Reviewed By", "Approved By"
                ]
            }
            
            for category, fields in field_categories.items():
                st.write(f"**{category}:**")
                for field in fields:
                    st.write(f"  • {field}")
        
        # Usage instructions
        with st.expander("📖 Usage Instructions"):
            st.markdown("""
            **How to use this template manager:**
            
            1. **Download Empty Template**: Click the button in the sidebar to get the base template
            
            2. **Upload Data File**: 
               - Upload CSV or Excel file with your data
               - Column headers must match the field names exactly
               - Images in Excel files will be automatically extracted
            
            3. **Manual Entry**: 
               - Fill in data manually using the form
               - All fields are optional
            
            4. **Generate Template**: 
               - Click generate to create filled template
               - Download the completed file
            
            5. **Image Handling**:
               - Upload images manually or include in Excel file
               - Supports PNG, JPG, JPEG formats
               - Images will be positioned in appropriate sections
            
            **Tips:**
            - Keep field names consistent with the mapping guide
            - Use the preview to understand the template structure
            - Test with sample data first
            """)
        
        # Sample data download
        if st.button("📥 Download Sample Data File"):
            # Create sample data
            sample_data = {
                'Revision No.': ['Rev-001'],
                'Date': ['2024-01-15'],
                'QC': ['John Doe'],
                'MM': ['Jane Smith'],
                'VP': ['Mike Johnson'],
                'Vendor Code': ['VEN-001'],
                'Vendor Name': ['ABC Packaging Solutions'],
                'Vendor Location': ['Mumbai, India'],
                'Part No.': ['PART-12345'],
                'Part Description': ['Electronic Component Housing'],
                'Part Unit Weight': ['150'],
                'Primary Packaging Type': ['Anti-static bag'],
                'Primary L-mm': ['200'],
                'Primary W-mm': ['150'],
                'Primary H-mm': ['50'],
                'Primary Qty/Pack': ['1'],
                'Primary Empty Weight': ['5'],
                'Primary Pack Weight': ['155'],
                'Secondary Packaging Type': ['Cardboard box'],
                'Secondary L-mm': ['250'],
                'Secondary W-mm': ['200'],
                'Secondary H-mm': ['100'],
                'Secondary Qty/Pack': ['10'],
                'Secondary Empty Weight': ['50'],
                'Secondary Pack Weight': ['1600'],
                'Procedure Step 1': ['Remove part from manufacturing line'],
                'Procedure Step 2': ['Inspect part for defects'],
                'Procedure Step 3': ['Place part in anti-static bag'],
                'Procedure Step 4': ['Seal the anti-static bag'],
                'Procedure Step 5': ['Place sealed bag in primary packaging'],
                'Procedure Step 6': ['Add padding if necessary'],
                'Procedure Step 7': ['Close primary packaging'],
                'Procedure Step 8': ['Place multiple primary packages in secondary box'],
                'Procedure Step 9': ['Add secondary padding and protection'],
                'Procedure Step 10': ['Seal secondary packaging and apply labels'],
                'Issued By': ['Production Manager'],
                'Reviewed By': ['Quality Manager'],
                'Approved By': ['Operations Director']
            }
            
            sample_df = pd.DataFrame(sample_data)
            csv_buffer = io.StringIO()
            sample_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            
            st.download_button(
                label="📤 Download Sample CSV",
                data=csv_buffer.getvalue(),
                file_name="sample_packaging_data.csv",
                mime="text/csv"
            )
            
            st.success("✅ Sample data file ready for download!")

if __name__ == "__main__":
    main()
