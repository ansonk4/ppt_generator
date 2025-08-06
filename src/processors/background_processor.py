import pandas as pd
from data_reader import DataReader
from ppt_generator import PptGenerator
from pptx.enum.chart import XL_LEGEND_POSITION

class BackgroundProcessor:
    def __init__(self, data_reader: DataReader, ppt_generator: PptGenerator):
        self.data_reader = data_reader
        self.ppt_generator = ppt_generator

    def _process_background_page1(self):
        self.ppt_generator.create_blank_slide("背景資料")

        col = "Banding"
        banding_dis = self.data_reader.get_col_distribution(col, normalize=False)
        self.ppt_generator.add_donut_chart(
            banding_dis, col, 'distribution',
            to_percent=False, title="受訪學生",
            sort=False,
            x=1, y=2, cx=4, cy=4
        )
        
        gender = self.data_reader.get_col_distribution("性別", normalize=False, return_dict=True)
        male_count, female_count = gender.get("1.0", 0), gender.get("2.0", 0)
        if male_count > 0 and female_count > 0:
            if male_count >= female_count:
                ratio = f"1 : {male_count / female_count:.2f}"
            else:
                ratio = f"{female_count / male_count:.2f} : 1"
        else:
            ratio = "N/A"

        text = "".join([
            f"{self.data_reader.data['學校編號'].nunique()}間中學\n\n",
            f"{len(self.data_reader.data)}受訪學生\n\n",
            f"男 : 女 = {ratio}",
        ])

        self.ppt_generator.add_textbox(
            text,
            x=6, y=2, cx=3, cy=2,
            font_size=18,
        )
        
    
    def _process_background_page23(self):
        self.ppt_generator.create_blank_slide("考生背景")

        col = "父母教育程度"
        edu_bg = self.data_reader.get_col_distribution(col, normalize=False)
        self.ppt_generator.add_donut_chart(
            edu_bg, col, "distribution",
            to_percent=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            title="父母教育程度",
            x=1, y=2, cx=4, cy=4
        )

        col = "高中選修學科"
        edu_bg = self.data_reader.get_col_distribution(col, normalize=False)
        self.ppt_generator.add_donut_chart(
            edu_bg, col, "distribution",
            title="高中選修學科",
            to_percent=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=5, y=2, cx=4, cy=4
        )

        self.ppt_generator.create_blank_slide("考生中五成績")
        chin = self.data_reader.get_col_distribution("中文成績", normalize=True, return_dict=True)
        eng = self.data_reader.get_col_distribution("英文成績", normalize=True, return_dict=True)
        math = self.data_reader.get_col_distribution("數學成績", normalize=True, return_dict=True)

        # Merge the three DataFrames on their index (assumed to be the grade/category)
        chin_series = pd.Series(chin, name="中文成績")
        eng_series = pd.Series(eng, name="英文成績")
        math_series = pd.Series(math, name="數學成績")
        merged_scores = pd.concat([chin_series, eng_series, math_series], axis=1)
        merged_scores = merged_scores.reset_index().rename(columns={"index": "score"})

        self.ppt_generator.add_bar_chart(
            merged_scores,
            category_column="score",
            value_columns=["中文成績", "英文成績", "數學成績"],
            title="2025考生中五成績",
            to_percentage=True,
            font_size=12,
            has_legend=True,
            x=1, y=1.5, cx=8, cy=5
        )

    def _process_background_page4(self):
        self.ppt_generator.create_blank_slide("中五成績不理想考生希望主修的科目")

        # Get slide width and height from ppt_generator
        slide_width = self.ppt_generator.prs.slide_width
        slide_height = self.ppt_generator.prs.slide_height

        cx, cy = 3, 5
        left_margin = 0.5
        graph_y = 1.7

        chin = self.data_reader.get_combined_distribution(
            columns=["希望修讀", "希望修讀_A", "希望修讀_B"],
            filtered_column="中文成績",
            filter_value=1,
        ).head(5)
        
        self.ppt_generator.add_bar_chart(
            chin,
            category_column="希望修讀",
            value_columns=["distribution"],
            title="中文25-49分",
            has_legend=False,
            to_percentage=True,
            font_size=12,
            x=left_margin, y=graph_y, cx=cx, cy=cy
        )

        eng = self.data_reader.get_combined_distribution(
            columns=["希望修讀", "希望修讀_A", "希望修讀_B"],
            filtered_column="英文成績",
            filter_value=1,
        ).head(5) 

        self.ppt_generator.add_bar_chart(
            eng,
            category_column="希望修讀",
            value_columns=["distribution"],
            title="英文25-49分",
            has_legend=False,
            font_size=12,
            to_percentage=True,
            x=left_margin + cx, y=graph_y, cx=cx, cy=cy
        )

        math = self.data_reader.get_combined_distribution(
            columns=["希望修讀", "希望修讀_A", "希望修讀_B"],
            filtered_column="數學成績",
            filter_value=1,

        ).head(5)

        self.ppt_generator.add_bar_chart(
            math,
            category_column="希望修讀",
            value_columns=["distribution"],
            title="數學25-49分",
            has_legend=False,
            to_percentage=True,
            font_size=12,
            x=left_margin + cx * 2, y=graph_y, cx=cx, cy=cy
        )

    def process_background_pages(self):
        """Generate all background pages in sequence."""
        self._process_background_page1()
        self._process_background_page23()
        self._process_background_page4()
