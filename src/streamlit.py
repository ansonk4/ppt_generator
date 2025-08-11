import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO
from dataclasses import dataclass
import traceback

# Import your existing modules
from presentation_generator import Config, PresentationGenerator

def display_validation_errors(validation_results: list[dict]) -> None:
    try:
        for result_dict in validation_results:
            if not result_dict:  # If the dictionary is not empty
                continue

            st.warning("❌ Invalid Data Found")
            
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

    except Exception as e:
        st.error("An error occurred while displaying validation errors.")
        st.write(traceback.format_exc())


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
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_output:
            output_path = tmp_output.name

        config = Config(data_file=temp_file_path, output_path=output_path)
        presentation_generator = PresentationGenerator(config)

        # data validation
        with st.spinner("Validating data..."):
            validation_columns_results = presentation_generator.validate_columns()
            if validation_columns_results:
                st.error("❌ Missing Required Columns")
                st.write("The following columns are missing from your Excel file:")
                for col in validation_columns_results:
                    st.write(f"• {col}")
                return

            validation_values_results = presentation_generator.validate_values()

        all_valid = all(not result for result in validation_values_results if result)
        if all_valid:
            st.success("✅ Data validation passed! No issues found.")
        else:
            display_validation_errors(validation_values_results)
            st.warning("⚠️ Warning: Issues were found in your data. You may still proceed, the system will replace any invalid data with NA, but the presentation may not generate correctly.")

        # ppt generation
        if st.button("📊 Generate PowerPoint", type="primary"):
            if not all_valid:
                presentation_generator.replace_invalid_values(validation_values_results)
            
            # Generate presentation with progress updates
            with st.spinner("Generating PowerPoint..."):
                presentation_generator.generate_presentation()

            # Read the generated file for download
            with open(output_path, 'rb') as f:
                pptx_data = f.read()

            # Generate filename based on uploaded file
            original_name = uploaded_file.name.rsplit('.', 1)[0]
            download_filename = f"{original_name}_presentation.pptx"
            
            st.success("Presentation generated successfully! You can download it below.")
            st.download_button(
                label="📥 Download PowerPoint Presentation",
                data=pptx_data,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
            os.unlink(output_path)
                

main()