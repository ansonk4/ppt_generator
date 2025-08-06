import pandas as pd 
from data_reader import DataReader
from ppt_generator import PptGenerator
from pptx.enum.chart import XL_LEGEND_POSITION
from datetime import datetime

class MajorProcessor:
    def __init__(self, data_reader: DataReader, ppt_generator: PptGenerator):
        self.data_reader = data_reader
        self.ppt_generator = ppt_generator

    def _process_page1(self):
        self.ppt_generator.create_blank_slide("受歡迎主修科目")

        data = self.data_reader.get_combined_distribution(
            ["希望修讀", "希望修讀_A", "希望修讀_B"],
        )[:5]


        self.ppt_generator.add_bar_chart(
            data,
            "希望修讀",
            ["distribution"],
            to_percentage=True,
            has_legend=False,
            x=3.5, cx=6
        )

        major = data["希望修讀"].tolist()
        distribution = data["distribution"].tolist()
        text = (
            f"第一位 {major[0]} {distribution[0]*100:.1f}%\n",
            f"第二位 {major[1]} {distribution[1]*100:.1f}%\n",
            f"第三位 {major[2]} {distribution[2]*100:.1f}%\n",
            f"第四位 {major[3]} {distribution[3]*100:.1f}%\n",
            f"第五位 {major[4]} {distribution[4]*100:.1f}%\n",
        )
        self.ppt_generator.add_textbox(
            "\n".join(text),
            x=0.5, cx=3, font_size=19
        )
    
    def _process_page2(self):
        self.ppt_generator.create_blank_slide("受歡迎主修科目走勢")
        self.ppt_generator.create_blank_slide("最受男女歡迎主修科目排名")

        male_data = self.data_reader.get_combined_distribution(
            ["希望修讀", "希望修讀_A", "希望修讀_B"],
            filtered_column="性別",
            filter_value=1
        )[:5]

        female_data = self.data_reader.get_combined_distribution(
           ["希望修讀", "希望修讀_A", "希望修讀_B"],
           filtered_column="性別",
           filter_value=2
       )[:5]

        self.ppt_generator.add_bar_chart(
            male_data,
            "希望修讀",
            ["distribution"],
            to_percentage=True,
            has_legend=False,
            horizontal=True,
            hide_y_axis=True,
            opposite_tick_labels=True,
            reserve_value_axis=True,
            color=(102, 204, 255),
            x=0.5, cx=4, y=3, cy=4
        )

        self.ppt_generator.add_bar_chart(
            female_data,
            "希望修讀",
            ["distribution"],
            to_percentage=True,
            has_legend=False,
            horizontal=True,
            hide_y_axis=True,
            opposite_tick_labels=True,
            color=(255, 153, 204),
            x=5.5, cx=4, y=3, cy=4
        )

        self.ppt_generator.add_img("img/male.png", x=3, y=1.5)
        self.ppt_generator.add_img("img/female.png", x=5.5, y=1.5)


    def _process_page3(self):
        self.ppt_generator.create_blank_slide("不受歡迎主修科目")
        
        data = self.data_reader.get_combined_distribution(
            ["不希望修讀", "不希望修讀_A", "不希望修讀_B"],
        )[:5]


        self.ppt_generator.add_bar_chart(
            data,
            "不希望修讀",
            ["distribution"],
            to_percentage=True,
            has_legend=False,
            x=3.5, cx=6
        )

        major = data["不希望修讀"].tolist()
        distribution = data["distribution"].tolist()
        text = (
            f"第一位 {major[0]} {distribution[0]*100:.1f}%\n",
            f"第二位 {major[1]} {distribution[1]*100:.1f}%\n",
            f"第三位 {major[2]} {distribution[2]*100:.1f}%\n",
            f"第四位 {major[3]} {distribution[3]*100:.1f}%\n",
            f"第五位 {major[4]} {distribution[4]*100:.1f}%\n",
        )
        self.ppt_generator.add_textbox(
            "\n".join(text),
            x=0.5, cx=3, font_size=19
        )



    def _process_page4(self):
        self.ppt_generator.create_blank_slide("最不受男女歡迎主修科目")

        male_data = self.data_reader.get_combined_distribution(
            ["不希望修讀", "不希望修讀_A", "不希望修讀_B"],
            filtered_column="性別",
            filter_value=1
        )[:5]

        female_data = self.data_reader.get_combined_distribution(
            ["不希望修讀", "不希望修讀_A", "不希望修讀_B"],
           filtered_column="性別",
           filter_value=2
       )[:5]

        self.ppt_generator.add_bar_chart(
            male_data,
            "不希望修讀",
            ["distribution"],
            to_percentage=True,
            has_legend=False,
            horizontal=True,
            hide_y_axis=True,
            opposite_tick_labels=True,
            reserve_value_axis=True,
            color=(102, 204, 255),
            x=0.5, cx=4, y=3, cy=4
        )

        self.ppt_generator.add_bar_chart(
            female_data,
            "不希望修讀",
            ["distribution"],
            to_percentage=True,
            has_legend=False,
            horizontal=True,
            hide_y_axis=True,
            opposite_tick_labels=True,
            color=(255, 153, 204),
            x=5.5, cx=4, y=3, cy=4
        )

        self.ppt_generator.add_img("img/male.png", x=3, y=1.5)
        self.ppt_generator.add_img("img/female.png", x=5.5, y=1.5)



    def process_major_pages(self):
        self._process_page1()
        self._process_page2()
        self._process_page3()
        self._process_page4()
