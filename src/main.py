import pandas as pd
from dataclasses import dataclass

from ppt_generator import PptGenerator
from data_reader import DataReader
from processors.background_processor import BackgroundProcessor
from processors.after_dse_processor import AfterDSEProcessor
from processors.major_processor import MajorProcessor
from processors.job_processor import JobProcessor
from processors.stem_processor import StemProcessor

@dataclass
class Config:
    data_file: str = "data/data2.xlsx"
    output_path: str = "output/presentation.pptx"


class PresentationGenerator:
    def __init__(self, config: Config):
        self.data_reader = DataReader(config.data_file)
        self.ppt_generator = PptGenerator()
        self.output_path = config.output_path

        self.background_processor = BackgroundProcessor(self.data_reader, self.ppt_generator)
        self.after_dse_processor = AfterDSEProcessor(self.data_reader, self.ppt_generator)
        self.major_processor = MajorProcessor(self.data_reader, self.ppt_generator)
        self.job_processor = JobProcessor(self.data_reader, self.ppt_generator)
        self.stem_processor = StemProcessor(self.data_reader, self.ppt_generator)

    def generate_presentation(self):

        self.background_processor.process_background_pages()
        self.after_dse_processor.process_after_dse_pages()
        self.major_processor.process_major_pages()
        self.job_processor.process_job_pages()
        self.stem_processor.process_stem_pages()

        self.ppt_generator.add_image_header_footer_to_all_slides("img/logo.png")
        self.ppt_generator.save(self.output_path)



if __name__ == "__main__":
    config = Config()
    presentation_generator = PresentationGenerator(config)
    presentation_generator.generate_presentation()
