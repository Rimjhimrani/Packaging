import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import io
import base64

class PackagingTemplateManager:
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
    
    def create_template_excel(self):
        """Create the Excel template with exact formatting"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Packaging Instruction"
        
        # Define styles
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Header Row
        ws.merge_cells('A1:F1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = header_fill
        ws['A1'].font = header_font
        ws['A1'].border = border
        
        # Revision info
        ws['A2'] = "Revision No."
        ws['B2'] = ""
        ws['C2'] = "Date"
        ws['D2'] = ""
        ws['E2'] = "QC"
        ws['F2'] = ""
        ws['G2'] = "MM"
        ws['H2'] = ""
        ws['I2'] = "VP"
        ws['J2'] = ""
        
        # Vendor Information
        ws['A4'] = "Vendor Information"
        ws['A5'] = "Code"
        ws['A6'] = "Name"
        ws['A7'] = "Location"
        
        # Part Information
        ws['E4'] = "Part Information"
        ws['E5'] = "Part No."
        ws['E6'] = "Description"
        ws['E7'] = "Unit Weight"
        
        # Primary Packaging
        ws['A9'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A10'] = "Packaging Type"
        ws['B10'] = "L-mm"
        ws['C10'] = "W-mm"
        ws['D10'] = "H-mm"
        ws['E10'] = "Qty/Pack"
        ws['F10'] = "Empty Weight"
        ws['G10'] = "Pack Weight"
        
        # Secondary Packaging
        ws['A13'] = "Secondary Packaging Instruction (Outer / External)"
        ws['A14'] = "Packaging Type"
        ws['B14'] = "L-mm"
        ws['C14'] = "W-mm"
        ws['D14'] = "H-mm"
        ws['E14'] = "Qty/Pack"
        ws['F14'] = "Empty Weight"
        ws['G14'] = "Pack Weight"
        
        # Packaging Procedure
        ws['A17'] = "Packaging Procedure"
        for i in range(1, 11):
            ws[f'A{17+i}'] = str(i)
        
        # Reference Images/Pictures
        ws['A29'] = "Reference Images/Pictures"
        ws['A30'] = "Primary Packaging"
        ws['B30'] = "Secondary Packaging"
        ws['C30'] = "Label"
        
        # Approval section
        ws['A35'] = "Issued By"
        ws['C35'] = "Reviewed By"
        ws['E35'] = "Approved By"
        
        # Apply borders to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
        
        return wb
    
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
            'Part No.': 'F5',
            'Part Description': 'F6',
            'Part Unit Weight': 'F7',
            'Primary Packaging Type': 'A11',
            'Primary L-mm': 'B11',
            'Primary W-mm': 'C11',
            'Primary H-mm': 'D11',
            'Primary Qty/Pack': 'E11',
            'Primary Empty Weight': 'F11',
            'Primary Pack Weight': 'G11',
            'Secondary Packaging Type': 'A15',
            'Secondary L-mm': 'B15',
            'Secondary W-mm': 'C15',
            'Secondary H-mm': 'D15',
            'Secondary Qty/Pack': 'E15',
            'Secondary Empty Weight': 'F15',
            'Secondary Pack Weight': 'G15',
            'Issued By': 'A36',
            'Reviewed By': 'C36',
            'Approved By': 'E36'
        }
        
        # Fill procedure steps
        for i in range(1, 11):
            key = f'Procedure Step {i}'
            if key in data_dict and data_dict[key]:
                ws[f'B{17+i}'] = data_dict[key]
        
        # Fill other mapped cells
        for key, cell_pos in cell_mapping.items():
            if key in data_dict and data_dict[key]:
                ws[cell_pos] = data_dict[key]
        
        return template_wb

def main():
    st.set_page_config(page_title="Packaging Instruction Template Manager", layout="wide")
    
    st.title("📦 Packaging Instruction Template Manager")
    st.markdown("Upload data to fill the packaging instruction template while keeping the original format intact.")
    
    # Initialize the template manager
    template_manager = PackagingTemplateManager()
    
    # Sidebar for admin controls
    with st.sidebar:
        st.header("Admin Controls")
        
        # Download empty template
        if st.button("📥 Download Empty Template"):
            wb = template_manager.create_template_excel()
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="Download Template",
                data=output.getvalue(),
                file_name="packaging_instruction_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📋 Data Input")
        
        # Option to upload CSV/Excel with data
        uploaded_file = st.file_uploader(
            "Upload data file (CSV/Excel)", 
            type=['csv', 'xlsx', 'xls']
        )
        
        if uploaded_file is not None:
            try:
                # Read the uploaded file
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.success("✅ Data file uploaded successfully!")
                st.dataframe(df)
                
                # Convert dataframe to dictionary for template filling
                data_dict = {}
                for col in df.columns:
                    if col in template_manager.template_fields:
                        data_dict[col] = df[col].iloc[0] if not df.empty else ""
                
                # Generate filled template
                if st.button("🔄 Generate Filled Template"):
                    wb = template_manager.create_template_excel()
                    filled_wb = template_manager.fill_template_with_data(wb, data_dict)
                    
                    output = io.BytesIO()
                    filled_wb.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="📤 Download Filled Template",
                        data=output.getvalue(),
                        file_name="filled_packaging_instruction.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Template filled successfully!")
                    
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
        
        # Manual data entry form
        st.subheader("✏️ Manual Data Entry")
        with st.expander("Enter data manually"):
            data_dict = {}
            
            # Header Information
            st.write("**Header Information**")
            col_a, col_b = st.columns(2)
            with col_a:
                data_dict['Revision No.'] = st.text_input("Revision No.")
                data_dict['QC'] = st.text_input("QC")
                data_dict['VP'] = st.text_input("VP")
            with col_b:
                data_dict['Date'] = st.text_input("Date")
                data_dict['MM'] = st.text_input("MM")
            
            # Vendor Information
            st.write("**Vendor Information**")
            data_dict['Vendor Code'] = st.text_input("Vendor Code")
            data_dict['Vendor Name'] = st.text_input("Vendor Name")
            data_dict['Vendor Location'] = st.text_input("Vendor Location")
            
            # Part Information
            st.write("**Part Information**")
            data_dict['Part No.'] = st.text_input("Part No.")
            data_dict['Part Description'] = st.text_input("Part Description")
            data_dict['Part Unit Weight'] = st.text_input("Part Unit Weight")
            
            # Primary Packaging
            st.write("**Primary Packaging**")
            col_a, col_b = st.columns(2)
            with col_a:
                data_dict['Primary Packaging Type'] = st.text_input("Primary Packaging Type")
                data_dict['Primary L-mm'] = st.text_input("Primary L-mm")
                data_dict['Primary W-mm'] = st.text_input("Primary W-mm")
                data_dict['Primary H-mm'] = st.text_input("Primary H-mm")
            with col_b:
                data_dict['Primary Qty/Pack'] = st.text_input("Primary Qty/Pack")
                data_dict['Primary Empty Weight'] = st.text_input("Primary Empty Weight")
                data_dict['Primary Pack Weight'] = st.text_input("Primary Pack Weight")
            
            # Secondary Packaging
            st.write("**Secondary Packaging**")
            col_a, col_b = st.columns(2)
            with col_a:
                data_dict['Secondary Packaging Type'] = st.text_input("Secondary Packaging Type")
                data_dict['Secondary L-mm'] = st.text_input("Secondary L-mm")
                data_dict['Secondary W-mm'] = st.text_input("Secondary W-mm")
                data_dict['Secondary H-mm'] = st.text_input("Secondary H-mm")
            with col_b:
                data_dict['Secondary Qty/Pack'] = st.text_input("Secondary Qty/Pack")
                data_dict['Secondary Empty Weight'] = st.text_input("Secondary Empty Weight")
                data_dict['Secondary Pack Weight'] = st.text_input("Secondary Pack Weight")
            
            # Packaging Procedures
            st.write("**Packaging Procedures**")
            for i in range(1, 11):
                data_dict[f'Procedure Step {i}'] = st.text_input(f"Step {i}")
            
            # Approval
            st.write("**Approval**")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                data_dict['Issued By'] = st.text_input("Issued By")
            with col_b:
                data_dict['Reviewed By'] = st.text_input("Reviewed By")
            with col_c:
                data_dict['Approved By'] = st.text_input("Approved By")
            
            # Generate template with manual data
            if st.button("🔄 Generate Template with Manual Data"):
                wb = template_manager.create_template_excel()
                filled_wb = template_manager.fill_template_with_data(wb, data_dict)
                
                output = io.BytesIO()
                filled_wb.save(output)
                output.seek(0)
                
                st.download_button(
                    label="📤 Download Filled Template",
                    data=output.getvalue(),
                    file_name="filled_packaging_instruction_manual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("✅ Template generated successfully!")
    
    with col2:
        st.subheader("📋 Template Fields Reference")
        st.markdown("These are the available fields that can be filled in the template:")
        
        # Show all available fields
        field_categories = {
            "Header Information": ['Revision No.', 'Date', 'QC', 'MM', 'VP'],
            "Vendor Information": ['Vendor Code', 'Vendor Name', 'Vendor Location'],
            "Part Information": ['Part No.', 'Part Description', 'Part Unit Weight'],
            "Primary Packaging": ['Primary Packaging Type', 'Primary L-mm', 'Primary W-mm', 
                                'Primary H-mm', 'Primary Qty/Pack', 'Primary Empty Weight', 
                                'Primary Pack Weight'],
            "Secondary Packaging": ['Secondary Packaging Type', 'Secondary L-mm', 'Secondary W-mm', 
                                  'Secondary H-mm', 'Secondary Qty/Pack', 'Secondary Empty Weight', 
                                  'Secondary Pack Weight'],
            "Packaging Procedures": [f'Procedure Step {i}' for i in range(1, 11)],
            "Approval": ['Issued By', 'Reviewed By', 'Approved By']
        }
        
        for category, fields in field_categories.items():
            with st.expander(category):
                for field in fields:
                    st.write(f"• {field}")
        
        st.info("💡 **Tip**: You can upload a CSV/Excel file with columns matching these field names, and the system will automatically map and fill the template.")

if __name__ == "__main__":
    main()
