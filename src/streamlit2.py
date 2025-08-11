import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO
from dataclasses import dataclass
import traceback

import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO
from dataclasses import dataclass
import traceback

from presentation_generator import Config, PresentationGenerator


def display_validation_errors(validation_results):
    """Display validation errors in a user-friendly format"""
    
    # Check for missing columns
    if validation_results and isinstance(validation_results[0], str):
        # This means we have missing columns
        st.error("❌ Missing Required Columns")
        st.write("The following columns are missing from your Excel file:")
        for col in validation_results:
            st.write(f"• {col}")
        return True
    
    # Check for invalid data
    has_errors = False
    if validation_results:
        for result_dict in validation_results:
            if result_dict:  # If the dictionary is not empty
                has_errors = True
                st.error("❌ Invalid Data Found")
                
                for column, invalid_entries in result_dict.items():
                    if column == "acceptable_values":
                        continue

                    st.write(f"**Column: {column}**")

                    display_entries = invalid_entries[:5]
                    
                    if display_entries:
                        error_df = pd.DataFrame([
                            {"Row Index": row_id, "Invalid Value": str(value)} 
                            for row_id, value in display_entries
                        ])

                        st.write("**Acceptable Values:**", ", ".join(map(str, result_dict.get("acceptable_values", []))))

                        st.dataframe(
                            error_df, 
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        # Show count of remaining errors if any
                        if len(invalid_entries) > 5:
                            remaining = len(invalid_entries) - 5
                            st.info(f"... and {remaining} more invalid entries in this column")
                    
                    st.write("---")
    
    return has_errors

def main():
    st.set_page_config(
        page_title="DSE Survey PPT Generator",
        page_icon="📊",
    )
    
    st.title("📊 DSE考生問卷調查 PowerPoint Generator")
    st.markdown("Upload your Excel file to generate a PowerPoint presentation for DSE survey analysis.")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload the Excel file containing DSE survey data"
    )
    
    if uploaded_file is not None:
        # Create temporary file for the uploaded data
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_file_path = tmp_file.name
        
        try:
            # Display file info
            st.success(f"✅ File uploaded successfully: {uploaded_file.name}")
            
            # Read and display basic info about the file
            with st.expander("📋 File Information", expanded=False):
                df_preview = pd.read_excel(uploaded_file)
                st.write(f"**Rows:** {len(df_preview)}")
                st.write(f"**Columns:** {len(df_preview.columns)}")
                st.write("**Column Names:**")
                cols_per_row = 3
                for i in range(0, len(df_preview.columns), cols_per_row):
                    cols = df_preview.columns[i:i+cols_per_row].tolist()
                    st.write("• " + " • ".join(cols))
                
                st.write("**Data Preview (First 5 rows):**")
                st.dataframe(df_preview.head(), use_container_width=True)
            
            # Validate data
            st.subheader("🔍 Data Validation")
            with st.spinner("Validating data..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_output:
                        output_path = tmp_output.name

                    config = Config(data_file=temp_file_path, output_path=output_path)
                    presentation_generator = PresentationGenerator(config)
                    validation_results = presentation_generator.validate_data()

                    if not validation_results:
                        st.success("✅ Data validation passed! No issues found.")
                    else:
                        # has_errors = display_validation_errors(validation_results)
                        has_errors = True
                        if has_errors:
                            st.warning("⚠️ Warning: Issues were found in your data. You may still proceed, but the presentation may not generate correctly.")
                    
                except Exception as e:
                    st.error(f"❌ Error during validation: {str(e)}")
                    st.error("Please check if your Excel file format is correct.")
                    # Clean up temp file
                    os.unlink(temp_file_path)
                    return

                
            if st.button("📊 Generate PowerPoint", type="primary"):
                # Create progress bar
                status_text = st.empty()
                try:
                    # Generate presentation with progress updates
                    with st.spinner("Generating PowerPoint..."):
                        presentation_generator.generate_presentation()
                    status_text.text("Presentation generated successfully!")
                    
                    # Read the generated file for download
                    with open(output_path, 'rb') as f:
                        pptx_data = f.read()
                    
                    # Provide download button
                    st.success("✅ Presentation generated successfully!")
                    
                    # Generate filename based on uploaded file
                    original_name = uploaded_file.name.rsplit('.', 1)[0]
                    download_filename = f"{original_name}_presentation.pptx"
                    
                    st.download_button(
                        label="📥 Download PowerPoint Presentation",
                        data=pptx_data,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    
                    # Clean up temp files
                    os.unlink(output_path)
                    
                except Exception as e:
                    st.error(f"❌ Error generating presentation: {str(e)}")
                    st.error("Please check your data format and try again.")
                    
                    # Show detailed error for debugging
                    with st.expander("🔍 Detailed Error Information"):
                        st.code(traceback.format_exc())
            
            # Clean up temp input file
            if os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
                
        except Exception as e:
            st.error(f"❌ Unexpected error: {str(e)}")
            # Clean up temp file
            if os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
    
    else:
        st.info("👆 Please upload an Excel file to get started.")

if __name__ == "__main__":
    main()