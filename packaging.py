import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.drawing.image import Image  # ✅ Corrected
import io
import base64
from PIL import Image as PILImage
import zipfile
import os
import tempfile

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
        white_font = Font(color="FFFFFF", bold=True)
        black_font = Font(color="000000", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        # Set column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 12
        ws.column_dimensions['J'].width = 12
        ws.column_dimensions['K'].width = 20
        
        # Header Row - "Packaging Instruction"
        ws.merge_cells('A1:J1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = blue_fill
        ws['A1'].font = white_font
        ws['A1'].border = border
        ws['A1'].alignment = center_alignment
        
        # Revision information row
        ws['A2'] = "Revision No."
        ws['A2'].border = border
        ws['B2'] = ""
        ws['B2'].border = border
        ws['C2'] = "Date"
        ws['C2'].border = border
        ws['D2'] = ""
        ws['D2'].border = border
        ws['E2'] = "QC"
        ws['E2'].border = border
        ws['F2'] = ""
        ws['F2'].border = border
        ws['G2'] = "MM"
        ws['G2'].border = border
        ws['H2'] = ""
        ws['H2'].border = border
        ws['I2'] = "VP"
        ws['I2'].border = border
        ws['J2'] = ""
        ws['J2'].border = border
        
        # Current Packaging section (right side)
        ws.merge_cells('K1:K2')
        ws['K1'] = "CURRENT PACKAGING"
        ws['K1'].fill = blue_fill
        ws['K1'].font = white_font
        ws['K1'].border = border
        ws['K1'].alignment = center_alignment
        
        # Vendor Information section
        ws.merge_cells('A4:D4')
        ws['A4'] = "Vendor Information"
        ws['A4'].fill = blue_fill
        ws['A4'].font = white_font
        ws['A4'].border = border
        ws['A4'].alignment = center_alignment
        
        ws['A5'] = "Code"
        ws['A5'].border = border
        ws['B5'] = ""
        ws['B5'].border = border
        ws['C5'] = ""
        ws['C5'].border = border
        ws['D5'] = ""
        ws['D5'].border = border
        
        ws['A6'] = "Name"
        ws['A6'].border = border
        ws['B6'] = ""
        ws['B6'].border = border
        ws['C6'] = ""
        ws['C6'].border = border
        ws['D6'] = ""
        ws['D6'].border = border
        
        ws['A7'] = "Location"
        ws['A7'].border = border
        ws['B7'] = ""
        ws['B7'].border = border
        ws['C7'] = ""
        ws['C7'].border = border
        ws['D7'] = ""
        ws['D7'].border = border
        
        # Part Information section
        ws.merge_cells('F4:J4')
        ws['F4'] = "Part Information"
        ws['F4'].fill = blue_fill
        ws['F4'].font = white_font
        ws['F4'].border = border
        ws['F4'].alignment = center_alignment
        
        ws['F5'] = "Part No."
        ws['F5'].border = border
        ws['G5'] = ""
        ws['G5'].border = border
        ws['H5'] = ""
        ws['H5'].border = border
        ws['I5'] = ""
        ws['I5'].border = border
        ws['J5'] = ""
        ws['J5'].border = border
        
        ws['F6'] = "Description"
        ws['F6'].border = border
        ws['G6'] = ""
        ws['G6'].border = border
        ws['H6'] = ""
        ws['H6'].border = border
        ws['I6'] = ""
        ws['I6'].border = border
        ws['J6'] = ""
        ws['J6'].border = border
        
        ws['F7'] = "Unit Weight"
        ws['F7'].border = border
        ws['G7'] = ""
        ws['G7'].border = border
        ws['H7'] = "W"
        ws['H7'].border = border
        ws['I7'] = ""
        ws['I7'].border = border
        ws['J7'] = "gms"
        ws['J7'].border = border
        
        # Current Packaging section (right side rows 3-8)
        for row in range(3, 9):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Primary Packaging Instruction
        ws.merge_cells('A9:J9')
        ws['A9'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A9'].fill = blue_fill
        ws['A9'].font = white_font
        ws['A9'].border = border
        ws['A9'].alignment = center_alignment
        
        # Current Packaging section (right side row 9)
        ws['K9'] = "CURRENT PACKAGING"
        ws['K9'].fill = blue_fill
        ws['K9'].font = white_font
        ws['K9'].border = border
        ws['K9'].alignment = center_alignment
        
        # Primary packaging headers
        ws['A10'] = "Packaging Type"
        ws['A10'].border = border
        ws['B10'] = "L-mm"
        ws['B10'].border = border
        ws['C10'] = "W-mm"
        ws['C10'].border = border
        ws['D10'] = "H-mm"
        ws['D10'].border = border
        ws['E10'] = "Qty/Pack"
        ws['E10'].border = border
        ws['F10'] = "Empty Weight"
        ws['F10'].border = border
        ws['G10'] = "Pack Weight"
        ws['G10'].border = border
        ws['H10'] = ""
        ws['H10'].border = border
        ws['I10'] = ""
        ws['I10'].border = border
        ws['J10'] = ""
        ws['J10'].border = border
        
        # Primary packaging data rows
        for row in range(11, 14):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
        
        # TOTAL row
        ws['H13'] = "TOTAL"
        ws['H13'].border = border
        ws['H13'].font = black_font
        
        # Current Packaging section (right side rows 10-13)
        for row in range(10, 14):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Secondary Packaging Instruction
        ws.merge_cells('A15:J15')
        ws['A15'] = "Secondary Packaging Instruction (Outer / External)"
        ws['A15'].fill = blue_fill
        ws['A15'].font = white_font
        ws['A15'].border = border
        ws['A15'].alignment = center_alignment
        
        # Secondary packaging headers
        ws['A16'] = "Packaging Type"
        ws['A16'].border = border
        ws['B16'] = "L-mm"
        ws['B16'].border = border
        ws['C16'] = "W-mm"
        ws['C16'].border = border
        ws['D16'] = "H-mm"
        ws['D16'].border = border
        ws['E16'] = "Qty/Pack"
        ws['E16'].border = border
        ws['F16'] = "Empty Weight"
        ws['F16'].border = border
        ws['G16'] = "Pack Weight"
        ws['G16'].border = border
        ws['H16'] = ""
        ws['H16'].border = border
        ws['I16'] = ""
        ws['I16'].border = border
        ws['J16'] = ""
        ws['J16'].border = border
        
        # Secondary packaging data rows
        for row in range(17, 20):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border
        
        # TOTAL row for secondary
        ws['H19'] = "TOTAL"
        ws['H19'].border = border
        ws['H19'].font = black_font
        
        # Problem section (right side)
        ws['K15'] = "PROBLEM IF ANY:"
        ws['K15'].border = border
        ws['K15'].font = black_font
        
        for row in range(16, 20):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Red section
        ws['K20'] = "CAUTION: REVISED DESIGN"
        ws['K20'].fill = red_fill
        ws['K20'].font = white_font
        ws['K20'].border = border
        ws['K20'].alignment = center_alignment
        
        # Packaging Procedure section
        ws.merge_cells('A21:J21')
        ws['A21'] = "Packaging Procedure"
        ws['A21'].fill = blue_fill
        ws['A21'].font = white_font
        ws['A21'].border = border
        ws['A21'].alignment = center_alignment
        
        # Packaging procedure steps
        for i in range(1, 11):
            ws[f'A{21+i}'] = str(i)
            ws[f'A{21+i}'].border = border
            for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws[f'{col}{21+i}'] = ""
                ws[f'{col}{21+i}'].border = border
        
        # Current Packaging section (right side for procedure)
        for row in range(21, 32):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
        
        # Reference Images/Pictures section
        ws.merge_cells('A33:J33')
        ws['A33'] = "Reference Images/Pictures"
        ws['A33'].fill = blue_fill
        ws['A33'].font = white_font
        ws['A33'].border = border
        ws['A33'].alignment = center_alignment
        
        # Image section headers
        ws.merge_cells('A34:C34')
        ws['A34'] = "Primary Packaging"
        ws['A34'].border = border
        ws['A34'].alignment = center_alignment
        
        ws.merge_cells('D34:G34')
        ws['D34'] = "Secondary Packaging"
        ws['D34'].border = border
        ws['D34'].alignment = center_alignment
        
        ws.merge_cells('H34:J34')
        ws['H34'] = "Label"
        ws['H34'].border = border
        ws['H34'].alignment = center_alignment
        
        # Image placeholders
        # Primary Packaging image area
        ws.merge_cells('A35:C41')
        ws['A35'] = "Primary\nPackaging"
        ws['A35'].border = border
        ws['A35'].alignment = center_alignment
        
        # Arrow 1
        ws['D39'] = "→"
        ws['D39'].border = border
        ws['D39'].alignment = center_alignment
        ws['D39'].font = Font(size=20, bold=True)
        
        # Secondary Packaging image area
        ws.merge_cells('E35:F41')
        ws['E35'] = "SECONDARY\nPACKAGING"
        ws['E35'].border = border
        ws['E35'].alignment = center_alignment
        ws['E35'].fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
        
        # Arrow 2
        ws['G39'] = "→"
        ws['G39'].border = border
        ws['G39'].alignment = center_alignment
        ws['G39'].font = Font(size=20, bold=True)
        
        # Label image area
        ws.merge_cells('H35:J41')
        ws['H35'] = "LABEL"
        ws['H35'].border = border
        ws['H35'].alignment = center_alignment
        
        # Green checkbox
        ws['K41'] = "☑"
        ws['K41'].border = border
        ws['K41'].alignment = center_alignment
        ws['K41'].font = Font(size=20, color="008000")
        
        # Approval section
        ws.merge_cells('A43:C43')
        ws['A43'] = "Issued By"
        ws['A43'].border = border
        ws['A43'].alignment = center_alignment
        
        ws.merge_cells('D43:G43')
        ws['D43'] = "Reviewed By"
        ws['D43'].border = border
        ws['D43'].alignment = center_alignment
        
        ws.merge_cells('H43:J43')
        ws['H43'] = "Approved By"
        ws['H43'].border = border
        ws['H43'].alignment = center_alignment
        
        # Signature boxes
        ws.merge_cells('A44:C47')
        ws['A44'] = ""
        ws['A44'].border = border
        
        ws.merge_cells('D44:G47')
        ws['D44'] = ""
        ws['D44'].border = border
        
        ws.merge_cells('H44:J47')
        ws['H44'] = ""
        ws['H44'].border = border
        
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
        
        # Mapping of data keys to cell positions
        cell_mapping = {
            'Revision No.': 'B2',
            'Date': 'D2',
            'QC': 'F2',
            'MM': 'H2',
            'VP': 'J2',
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
            'Secondary Packaging Type': 'A17',
            'Secondary L-mm': 'B17',
            'Secondary W-mm': 'C17',
            'Secondary H-mm': 'D17',
            'Secondary Qty/Pack': 'E17',
            'Secondary Empty Weight': 'F17',
            'Secondary Pack Weight': 'G17',
            'Issued By': 'A44',
            'Reviewed By': 'D44',
            'Approved By': 'H44'
        }
        
        # Fill procedure steps
        for i in range(1, 11):
            key = f'Procedure Step {i}'
            if key in data_dict and data_dict[key]:
                ws[f'B{21+i}'] = data_dict[key]
        
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
                manual_data['Primary H-mm'] = st.text_input("Primary H-mm", key="primary_h")
            with col_b:
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
                manual_data['Secondary H-mm'] = st.text_input("Secondary H-mm", key="secondary_h")
            with col_b:
                manual_data['Secondary Qty/Pack'] = st.text_input("Secondary Qty/Pack", key="secondary_qty")
                manual_data['Secondary Empty Weight'] = st.text_input("Secondary Empty Weight", key="secondary_empty")
                manual_data['Secondary Pack Weight'] = st.text_input("Secondary Pack Weight", key="secondary_pack")
            
            # Packaging Procedures
            st.write("**Packaging Procedures**")
            for i in range(1, 11):
                manual_data[f'Procedure Step {i}'] = st.text_input(f"Step {i}", key=f"step_{i}")
            
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
                    label="📤 Download Filled Template",
                    data=output.getvalue(),
                    file_name="filled_exact_packaging_instruction_manual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("✅ Template filled with manual data successfully!")
    
    with col2:
        st.subheader("📋 Template Preview & Instructions")
        
        # Show template structure
        st.markdown("""
        **Template Structure:**
        - **Header Section**: Revision info, QC, MM, VP details
        - **Vendor Information**: Code, Name, Location
        - **Part Information**: Part No., Description, Unit Weight
        - **Primary Packaging**: Type, dimensions, quantities, weights
        - **Secondary Packaging**: Type, dimensions, quantities, weights
        - **Packaging Procedures**: 10-step process instructions
        - **Reference Images**: Primary, Secondary, Label sections
        - **Approval Section**: Issued By, Reviewed By, Approved By
        """)
        
        # Show field mapping
        with st.expander("📊 Field Mapping Guide"):
            st.markdown("""
            **CSV/Excel Column Headers should match:**
            - `Revision No.`, `Date`, `QC`, `MM`, `VP`
            - `Vendor Code`, `Vendor Name`, `Vendor Location`
            - `Part No.`, `Part Description`, `Part Unit Weight`
            - `Primary Packaging Type`, `Primary L-mm`, `Primary W-mm`, `Primary H-mm`
            - `Primary Qty/Pack`, `Primary Empty Weight`, `Primary Pack Weight`
            - `Secondary Packaging Type`, `Secondary L-mm`, `Secondary W-mm`, `Secondary H-mm`
            - `Secondary Qty/Pack`, `Secondary Empty Weight`, `Secondary Pack Weight`
            - `Procedure Step 1` through `Procedure Step 10`
            - `Issued By`, `Reviewed By`, `Approved By`
            """)
        
        # Show sample data format
        with st.expander("📋 Sample Data Format"):
            sample_data = {
                'Revision No.': ['Rev-001'],
                'Date': ['2024-01-15'],
                'QC': ['John Doe'],
                'MM': ['Jane Smith'],
                'VP': ['Bob Johnson'],
                'Vendor Code': ['VEN001'],
                'Vendor Name': ['ABC Packaging Ltd'],
                'Vendor Location': ['Mumbai, India'],
                'Part No.': ['PART-12345'],
                'Part Description': ['Electronic Component'],
                'Part Unit Weight': ['50'],
                'Primary Packaging Type': ['Plastic Bag'],
                'Primary L-mm': ['100'],
                'Primary W-mm': ['80'],
                'Primary H-mm': ['20'],
                'Primary Qty/Pack': ['10'],
                'Primary Empty Weight': ['5'],
                'Primary Pack Weight': ['55'],
                'Secondary Packaging Type': ['Cardboard Box'],
                'Secondary L-mm': ['200'],
                'Secondary W-mm': ['150'],
                'Secondary H-mm': ['100'],
                'Secondary Qty/Pack': ['5'],
                'Secondary Empty Weight': ['25'],
                'Secondary Pack Weight': ['300'],
                'Procedure Step 1': ['Remove parts from production line'],
                'Procedure Step 2': ['Inspect for defects'],
                'Procedure Step 3': ['Place in primary packaging'],
                'Issued By': ['Production Manager'],
                'Reviewed By': ['Quality Manager'],
                'Approved By': ['Plant Manager']
            }
            
            sample_df = pd.DataFrame(sample_data)
            st.dataframe(sample_df)
        
        # Instructions
        st.markdown("""
        **Instructions:**
        1. **Download Empty Template**: Use the sidebar to get the base template
        2. **Upload Data File**: Upload CSV/Excel with your data (optional images)
        3. **Manual Entry**: Fill data manually using the form
        4. **Generate Template**: Create filled template with your data
        5. **Add Images**: Upload images manually or include them in Excel file
        
        **Features:**
        - ✅ Exact template format matching requirements
        - ✅ Data extraction from CSV/Excel files
        - ✅ Image extraction from Excel files
        - ✅ Manual data entry form
        - ✅ Template generation with filled data
        - ✅ Professional formatting and styling
        """)
        
        # Show current template fields
        with st.expander("🔍 Available Template Fields"):
            for category, fields in {
                "Header": ['Revision No.', 'Date', 'QC', 'MM', 'VP'],
                "Vendor": ['Vendor Code', 'Vendor Name', 'Vendor Location'],
                "Part": ['Part No.', 'Part Description', 'Part Unit Weight'],
                "Primary Packaging": ['Primary Packaging Type', 'Primary L-mm', 'Primary W-mm', 'Primary H-mm', 'Primary Qty/Pack', 'Primary Empty Weight', 'Primary Pack Weight'],
                "Secondary Packaging": ['Secondary Packaging Type', 'Secondary L-mm', 'Secondary W-mm', 'Secondary H-mm', 'Secondary Qty/Pack', 'Secondary Empty Weight', 'Secondary Pack Weight'],
                "Procedures": [f'Procedure Step {i}' for i in range(1, 11)],
                "Approval": ['Issued By', 'Reviewed By', 'Approved By']
            }.items():
                st.write(f"**{category}:**")
                st.write(", ".join(fields))
    
    # Footer
    st.markdown("---")
    st.markdown("**📦 Exact Packaging Instruction Template Manager** - Created with Streamlit")
    st.markdown("*Upload your data, generate professional packaging instruction templates with ease!*")

if __name__ == "__main__":
    main()
