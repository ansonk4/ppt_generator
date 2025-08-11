import pandas as pd
from dataclasses import dataclass
import streamlit as st

from ppt_generator import PptGenerator
from data_reader import DataReader
from data_validator import DataValidator
from processors.background_processor import BackgroundProcessor
from processors.after_dse_processor import AfterDSEProcessor
from processors.major_processor import MajorProcessor
from processors.job_processor import JobProcessor
from processors.stem_processor import StemProcessor
from processors.gba_processor import GBAProcessor

@dataclass
class Config:
    data_file: str = "data/2025data.xlsx"
    output_path: str = "output/presentation.pptx"


class PresentationGenerator:
    def __init__(self, config: Config):
        self.data_reader = DataReader(config.data_file)
        self.data_validator = DataValidator(self.data_reader.data)
        self.ppt_generator = PptGenerator()
        self.output_path = config.output_path

        self.background_processor = BackgroundProcessor(self.data_reader, self.ppt_generator)
        self.after_dse_processor = AfterDSEProcessor(self.data_reader, self.ppt_generator)
        self.major_processor = MajorProcessor(self.data_reader, self.ppt_generator)
        self.job_processor = JobProcessor(self.data_reader, self.ppt_generator)
        self.stem_processor = StemProcessor(self.data_reader, self.ppt_generator)
        self.gba_processor = GBAProcessor(self.data_reader, self.ppt_generator)

    def validate_columns(self) -> list[str]:
        """Validate if the required columns are present in the data"""
        return self.data_validator.validate_column()

    def validate_values(self) -> list[dict]:
        """Validate specific columns for acceptable values"""
        return self.data_validator.validate_value()
    
    def replace_invalid_values(self, validate_results: list[dict]) -> None:
        """Replace invalid values in the data with NaN"""
        self.data_reader.replace_invalid_values(validate_results)

    def generate_presentation(self):
        try:
            self.ppt_generator.create_title_slide("2025年\nDSE考生問卷調查\n未來勞動力供應預測")
        except Exception as e:
            st.write(f"Error creating title slide: {e}")
        
        try:
            self.background_processor.process_background_pages()
        except Exception as e:
            st.write(f"Error processing background pages: {e}")
            
        try:
            self.after_dse_processor.process_after_dse_pages()
        except Exception as e:
            st.write(f"Error processing after DSE pages: {e}")
            
        try:
            self.major_processor.process_major_pages()
        except Exception as e:
            st.write(f"Error processing major pages: {e}")
            
        try:
            self.job_processor.process_job_pages()
        except Exception as e:
            st.write(f"Error processing job pages: {e}")
            
        try:
            self.stem_processor.process_stem_pages()
        except Exception as e:
            st.write(f"Error processing STEM pages: {e}")
            
        try:
            self.gba_processor.process_gba_pages()
        except Exception as e:
            st.write(f"Error processing GBA pages: {e}")

        try:
            self.ppt_generator.add_image_header_footer_to_all_slides("img/logo.png")
        except Exception as e:
            st.write(f"Error adding image header/footer to slides: {e}")
            
        try:
            self.ppt_generator.save(self.output_path)
        except Exception as e:
            st.write(f"Error saving presentation: {e}")


if __name__ == "__main__":
    config = Config()
    presentation_generator = PresentationGenerator(config)
    presentation_generator.generate_presentation()
