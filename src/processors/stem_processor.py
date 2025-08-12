import pandas as pd 
import streamlit as st
from data_reader import DataReader
from ppt_generator import PptGenerator
from pptx.enum.chart import XL_LEGEND_POSITION
from datetime import datetime

class StemProcessor:
    def __init__(self, data_reader: DataReader, ppt_generator: PptGenerator):
        self.data_reader = data_reader
        self.ppt_generator = ppt_generator

    
    def _process_page1(self):
        self.ppt_generator.create_blank_slide("DSE考生STEM學習項目參與分佈")

        stem_dis = self.data_reader.get_col_distribution("參加STEM", normalize=True, return_dict=False)
        current_year = str(datetime.now().year)
        stem_dis = stem_dis.rename(columns={"distribution": current_year})
        
        self.ppt_generator.add_donut_chart(
            stem_dis,
            "參加STEM", current_year,
            to_percent=True,
            title="參加STEM學習項目分佈",
            has_legend=False,
            has_data_labels=True,
            x=5.5, y=2, cx=4, cy=4
        )

        stem_dis[current_year] = stem_dis[current_year].apply(lambda x: f"{x:.1%}")

        self.ppt_generator.add_table(
            stem_dis,
            rows=3, cols=6,
            index=False,
            font_size=12,
            x=1, y=3, cx=4, cy=2,
        )
   
    def _process_major_or_job_page(self, title: str, cols: list[str], target_major_or_job: list[str | int]):
        self.ppt_generator.create_blank_slide(title)

        stem_data = self.data_reader.get_combined_distribution(
            columns=cols,
            filtered_column="參加STEM",
            filter_value="有"
        )

        no_stem_data = self.data_reader.get_combined_distribution(
            columns=cols,
            filtered_column="參加STEM",
            filter_value="沒有"
        )

        # 58: Math
        stem_filtered = stem_data[stem_data[cols[0]].isin(target_major_or_job)]
        no_stem_filtered = no_stem_data[no_stem_data[cols[0]].isin(target_major_or_job)]

        # Merge the two DataFrames for bar chart
        merged_df = pd.merge(
            stem_filtered,
            no_stem_filtered,
            on=cols[0],
            suffixes=("_stem", "_no_stem")
        )

        self.ppt_generator.add_bar_chart(
            merged_df.rename(columns={"distribution_stem": "參加STEM", "distribution_no_stem": "沒有參加STEM"}),
            category_column=cols[0],
            value_columns=["參加STEM", "沒有參加STEM"],
            to_percentage=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=6, y=2, cx=4, cy=4
        )

        
        stem_data["distribution"] = stem_data["distribution"].apply(lambda x: f"{x:.1%}")
        no_stem_data["distribution"] = no_stem_data["distribution"].apply(lambda x: f"{x:.1%}")

        self.ppt_generator.add_table(
            stem_data.head(10).rename(columns={cols[0]: "參加STEM", "distribution": "百分比"}),
            rows=11, cols=2,
            index=False,
            font_size=12,
            x=0.5, y=2, cx=2.7, cy=4.5,
        )

        self.ppt_generator.add_table(
            no_stem_data.head(10).rename(columns={cols[0]: "沒有參加STEM", "distribution": "百分比"}),
            rows=11, cols=2,
            index=False,
            font_size=12,
            x=3, y=2, cx=2.7, cy=4.5,
        )

    def _process_page2(self):
        self.ppt_generator.create_blank_slide("STEM學習項目影響程度")
        data = self.data_reader.get_col_distribution(
            "STEM影響職業選擇程度", normalize=True, return_dict=False, exclude=0,
        )
        self.ppt_generator.add_donut_chart(
            data, "STEM影響職業選擇程度", "distribution",
            to_percent=True,
            title="職業選擇影響程度",
            has_legend=False,
            has_data_labels=True,
            x=0.1, y=2, cx=4, cy=4
        )

        cols = ["領導能力", "團隊合作", "創新思維", "科學知識", "解難能力"]

        data = [
            self.data_reader.get_col_distribution(
                col, normalize=True, return_dict=True, exclude=0,
            ) for col in cols
        ]

        data = pd.DataFrame(data)
        data["index"] = cols

        data = data.dropna(axis=1)

        # norm by row after drop NA to make sure sum is 1
        value_columns = [col for col in data.columns if col != "index"]
        data[value_columns] = data[value_columns].div(data[value_columns].sum(axis=1), axis=0)

        self.ppt_generator.add_stacked_bar(
            data,
            category_column="index",
            value_columns=value_columns,
            title="STEM學習項目對各方面能力的影響",
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=3.6, y=2, cx=6, cy=5
        )

    def process_stem_pages(self):
        try:
            self.ppt_generator.create_section_slide("STEM學習項目對選科及就業取向影響")
        except Exception as e:
            st.error(f"Failed to create STEM section slide: {str(e)}")
            
        try:
            self._process_page1()
        except Exception as e:
            st.error(f"Failed to process STEM page 1: {str(e)}")
            
        try:
            self._process_page2()
        except Exception as e:
            st.error(f"Failed to process STEM page 2: {str(e)}")
            
        try:
            self._process_major_or_job_page("STEM參與率不同下受歡迎主修科目分佈", ["希望修讀", "希望修讀_A", "希望修讀_B"], ["電腦工程", "電腦科學"])
        except Exception as e:
            st.error(f"Failed to process STEM popular majors page: {str(e)}")
            
        try:
            self._process_major_or_job_page("STEM參與率不同下不受歡迎主修科目分佈", ["不希望修讀", "不希望修讀_A", "不希望修讀_B"], ["數學"])
        except Exception as e:
            st.error(f"Failed to process STEM unpopular majors page: {str(e)}")
            
        try:
            self._process_major_or_job_page("STEM參與率不同下受歡迎職業分佈", ["希望從事", "希望從事_A", "希望從事_B"], ["資訊科技", "電腦工程"])
        except Exception as e:
            st.error(f"Failed to process STEM popular jobs page: {str(e)}")
            
        try:
            self._process_major_or_job_page("STEM參與率不同下不受歡迎職業分佈", ["不希望從事", "不希望從事_A", "不希望從事_B"], ["電腦工程"])
        except Exception as e:
            st.error(f"Failed to process STEM unpopular jobs page: {str(e)}")
