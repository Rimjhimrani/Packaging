import streamlit as st
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
from zipfile import ZipFile
import xml.etree.ElementTree as ET

# Import the ExactPackagingTemplateManager class
# (Include the entire class definition here or import it)

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
            'Approved By': '',
            
            # Image fields
            'Primary Packaging Image': '',
            'Secondary Packaging Image': '',
            'Label Image': '',
            'Current Primary Image': '',
            'Current Secondary Image': '',
            'Current Label Image': ''
        }
        
        # Mapping for image column headers to template positions
        self.image_column_mapping = {
            'Primary Packaging Image': 'primary_packaging',
            'Secondary Packaging Image': 'secondary_packaging', 
            'Label Image': 'label',
            'Current Primary Image': 'current_primary',
            'Current Secondary Image': 'current_secondary',
            'Current Label Image': 'current_label',
            'Primary Image': 'primary_packaging',
            'Secondary Image': 'secondary_packaging',
            'Label': 'label'
        }
    
    # Include all the methods from the previous class definition
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
        
        # Set column widths
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

        # Create the complete template structure
        # (Include the complete template creation code here)
        # For brevity, I'll include key parts - you'll need the full method
        
        # Header Row
        ws.merge_cells('A1:K1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = blue_fill
        ws['A1'].font = white_font
        ws['A1'].border = border
        ws['A1'].alignment = center_alignment
        
        # Add all other template elements...
        # (Include the rest of the template creation code)
        
        return wb
    
    def fill_template_with_data(self, template_wb, data_dict):
        """Fill the template with provided data"""
        ws = template_wb.active
        
        # Cell mapping
        cell_mapping = {
            'Revision No.': 'B2',
            'Date': 'G2',
            'QC': 'I2',
            'MM': 'J2',
            'VP': 'B3',
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
            'Procedure Step 1': 'B20',
            'Procedure Step 2': 'B21',
            'Procedure Step 3': 'B22',
            'Procedure Step 4': 'B23',
            'Procedure Step 5': 'B24',
            'Procedure Step 6': 'B25',
            'Procedure Step 7': 'B26',
            'Procedure Step 8': 'B27',
            'Procedure Step 9': 'B28',
            'Procedure Step 10': 'B29',
            'Issued By': 'A39',
            'Reviewed By': 'D39',
            'Approved By': 'H39'
        }
        
        # Fill the template with data
        for key, cell_ref in cell_mapping.items():
            if key in data_dict and data_dict[key]:
                ws[cell_ref] = data_dict[key]
        
        return template_wb
    
    def save_template_to_bytes(self, workbook):
        """Save the workbook to bytes for download"""
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output.getvalue()
    
    def get_sample_data(self):
        """Return sample data for testing"""
        return {
            'Revision No.': 'Rev-001',
            'Date': '2024-01-15',
            'QC': 'QC-001',
            'MM': 'MM-001',
            'VP': 'VP-001',
            'Vendor Code': 'V001',
            'Vendor Name': 'ABC Manufacturing',
            'Vendor Location': 'Mumbai, India',
            'Part No.': 'P123456',
            'Part Description': 'Sample Electronic Component',
            'Part Unit Weight': '50g',
            'Primary Packaging Type': 'Anti-static bag',
            'Primary L-mm': '100',
            'Primary W-mm': '80',
            'Primary H-mm': '10',
            'Primary Qty/Pack': '1',
            'Primary Empty Weight': '5g',
            'Primary Pack Weight': '55g',
            'Secondary Packaging Type': 'Cardboard box',
            'Secondary L-mm': '120',
            'Secondary W-mm': '100',
            'Secondary H-mm': '30',
            'Secondary Qty/Pack': '10',
            'Secondary Empty Weight': '50g',
            'Secondary Pack Weight': '600g',
            'Procedure Step 1': 'Inspect part for damage',
            'Procedure Step 2': 'Place part in anti-static bag',
            'Procedure Step 3': 'Seal the bag properly',
            'Procedure Step 4': 'Apply part label',
            'Procedure Step 5': 'Place in secondary packaging',
            'Procedure Step 6': 'Add padding material',
            'Procedure Step 7': 'Seal secondary packaging',
            'Procedure Step 8': 'Apply shipping label',
            'Procedure Step 9': 'Quality check',
            'Procedure Step 10': 'Ready for shipment',
            'Issued By': 'John Smith',
            'Reviewed By': 'Jane Doe',
            'Approved By': 'Mike Johnson'
        }

def main():
    st.set_page_config(
        page_title="Packaging Template Generator",
        page_icon="📦",
        layout="wide"
    )
    
    st.title("📦 Packaging Template Generator")
    st.markdown("Generate standardized packaging instruction templates for your products.")
    
    # Initialize the template manager
    template_manager = ExactPackagingTemplateManager()
    
    # Sidebar for navigation
    st.sidebar.title("Options")
    mode = st.sidebar.selectbox(
        "Select Mode",
        ["Create New Template", "Fill from Data File", "Download Empty Template"]
    )
    
    if mode == "Create New Template":
        st.header("Create New Packaging Template")
        
        # Create tabs for different sections
        tab1, tab2, tab3, tab4 = st.tabs(["Header & Vendor", "Part & Packaging", "Procedures", "Approval"])
        
        data_dict = {}
        
        with tab1:
            st.subheader("Header Information")
            col1, col2 = st.columns(2)
            
            with col1:
                data_dict['Revision No.'] = st.text_input("Revision No.", value="Rev-001")
                data_dict['Date'] = st.date_input("Date").strftime("%Y-%m-%d")
                data_dict['QC'] = st.text_input("QC")
                
            with col2:
                data_dict['MM'] = st.text_input("MM")
                data_dict['VP'] = st.text_input("VP")
            
            st.subheader("Vendor Information")
            col1, col2 = st.columns(2)
            
            with col1:
                data_dict['Vendor Code'] = st.text_input("Vendor Code", placeholder="V001")
                data_dict['Vendor Name'] = st.text_input("Vendor Name", placeholder="ABC Manufacturing")
                
            with col2:
                data_dict['Vendor Location'] = st.text_input("Vendor Location", placeholder="Mumbai, India")
        
        with tab2:
            st.subheader("Part Information")
            col1, col2 = st.columns(2)
            
            with col1:
                data_dict['Part No.'] = st.text_input("Part No.", placeholder="P123456")
                data_dict['Part Description'] = st.text_input("Part Description", placeholder="Electronic Component")
                
            with col2:
                data_dict['Part Unit Weight'] = st.text_input("Part Unit Weight", placeholder="50g")
            
            st.subheader("Primary Packaging")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                data_dict['Primary Packaging Type'] = st.text_input("Primary Packaging Type", placeholder="Anti-static bag")
                data_dict['Primary L-mm'] = st.text_input("Primary L-mm", placeholder="100")
                data_dict['Primary W-mm'] = st.text_input("Primary W-mm", placeholder="80")
                
            with col2:
                data_dict['Primary H-mm'] = st.text_input("Primary H-mm", placeholder="10")
                data_dict['Primary Qty/Pack'] = st.text_input("Primary Qty/Pack", placeholder="1")
                
            with col3:
                data_dict['Primary Empty Weight'] = st.text_input("Primary Empty Weight", placeholder="5g")
                data_dict['Primary Pack Weight'] = st.text_input("Primary Pack Weight", placeholder="55g")
            
            st.subheader("Secondary Packaging")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                data_dict['Secondary Packaging Type'] = st.text_input("Secondary Packaging Type", placeholder="Cardboard box")
                data_dict['Secondary L-mm'] = st.text_input("Secondary L-mm", placeholder="120")
                data_dict['Secondary W-mm'] = st.text_input("Secondary W-mm", placeholder="100")
                
            with col2:
                data_dict['Secondary H-mm'] = st.text_input("Secondary H-mm", placeholder="30")
                data_dict['Secondary Qty/Pack'] = st.text_input("Secondary Qty/Pack", placeholder="10")
                
            with col3:
                data_dict['Secondary Empty Weight'] = st.text_input("Secondary Empty Weight", placeholder="50g")
                data_dict['Secondary Pack Weight'] = st.text_input("Secondary Pack Weight", placeholder="600g")
        
        with tab3:
            st.subheader("Packaging Procedures")
            col1, col2 = st.columns(2)
            
            with col1:
                for i in range(1, 6):
                    data_dict[f'Procedure Step {i}'] = st.text_area(
                        f"Step {i}", 
                        placeholder=f"Enter procedure step {i}",
                        height=50
                    )
                    
            with col2:
                for i in range(6, 11):
                    data_dict[f'Procedure Step {i}'] = st.text_area(
                        f"Step {i}", 
                        placeholder=f"Enter procedure step {i}",
                        height=50
                    )
        
        with tab4:
            st.subheader("Approval")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                data_dict['Issued By'] = st.text_input("Issued By", placeholder="John Smith")
                
            with col2:
                data_dict['Reviewed By'] = st.text_input("Reviewed By", placeholder="Jane Doe")
                
            with col3:
                data_dict['Approved By'] = st.text_input("Approved By", placeholder="Mike Johnson")
        
        # Generate template button
        if st.button("Generate Template", type="primary"):
            try:
                # Create template with data
                wb = template_manager.create_filled_template(data_dict)
                
                # Save to bytes
                template_bytes = template_manager.save_template_to_bytes(wb)
                
                # Download button
                st.download_button(
                    label="Download Template",
                    data=template_bytes,
                    file_name=f"packaging_template_{data_dict.get('Part No.', 'template')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("Template generated successfully!")
                
            except Exception as e:
                st.error(f"Error generating template: {str(e)}")
    
    elif mode == "Fill from Data File":
        st.header("Fill Template from Data File")
        
        uploaded_file = st.file_uploader(
            "Upload CSV or Excel file with data",
            type=['csv', 'xlsx', 'xls']
        )
        
        if uploaded_file is not None:
            try:
                # Extract data from file
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.subheader("Data Preview")
                st.dataframe(df)
                
                # Convert to dictionary
                data_dict = {}
                for col in df.columns:
                    col_clean = col.strip()
                    if col_clean in template_manager.template_fields:
                        data_dict[col_clean] = str(df[col].iloc[0]) if not df.empty and pd.notna(df[col].iloc[0]) else ""
                
                if st.button("Generate Template from Data", type="primary"):
                    # Create template with data
                    wb = template_manager.create_filled_template(data_dict)
                    
                    # Save to bytes
                    template_bytes = template_manager.save_template_to_bytes(wb)
                    
                    # Download button
                    st.download_button(
                        label="Download Filled Template",
                        data=template_bytes,
                        file_name=f"packaging_template_filled_{data_dict.get('Part No.', 'template')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Template generated from data successfully!")
                    
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
    
    elif mode == "Download Empty Template":
        st.header("Download Empty Template")
        st.markdown("Download a blank template to fill manually or use as a reference.")
        
        if st.button("Generate Empty Template", type="primary"):
            try:
                # Create empty template
                wb = template_manager.create_exact_template_excel()
                
                # Save to bytes
                template_bytes = template_manager.save_template_to_bytes(wb)
                
                # Download button
                st.download_button(
                    label="Download Empty Template",
                    data=template_bytes,
                    file_name="packaging_template_empty.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("Empty template generated successfully!")
                
            except Exception as e:
                st.error(f"Error generating template: {str(e)}")
    
    # Sample data section
    with st.expander("View Sample Data Format"):
        sample_data = template_manager.get_sample_data()
        st.json(sample_data)

if __name__ == "__main__":
    main()
