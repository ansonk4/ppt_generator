import pandas as pd 
from data_reader import DataReader
from ppt_generator import PptGenerator
from pptx.enum.chart import XL_LEGEND_POSITION
from datetime import datetime


class GBAProcessor:
    def __init__(self, data_reader: DataReader, ppt_generator: PptGenerator):
        self.data_reader = data_reader
        self.ppt_generator = ppt_generator

    def _process_page1(self):
        self.ppt_generator.create_blank_slide("大灣區政策\n對選科及就業取向影響")

        self.ppt_generator.create_blank_slide("DSE考生對大灣區政策了解")


        self.ppt_generator.add_textbox(
            [
                "粵港澳大灣區（大灣區）包括香港、澳門兩個特別行政區，",
                "和廣東省九個城市 (廣州、深圳、珠海、佛山、惠州、東莞、中山、江門、肇慶)"
            ],
            x=0.5, y=1, cx=8, cy=1,
            font_size=14
        )       

        
        gba = self.data_reader.get_col_distribution(
            "大灣區了解", normalize=True, return_dict=False
        )

        self.ppt_generator.add_donut_chart(
            gba, "大灣區了解", "distribution",
            to_percent=True,
            title="DSE考生對大灣區政策了解",
            has_legend=False,
            has_data_labels=True,
            x=5, y=2, cx=4, cy=4
        )




    def _process_page2(self):
        self.ppt_generator.create_blank_slide("大灣區政策香港定位")
        self.ppt_generator.create_blank_slide("考生接收大灣區資訊來源")
        
        cols = [
            "公社科","內地考察",
            "政府資訊","新聞媒體","網上資訊",
            "內地交流","校內講座","朋輩及老師"
        ]
        data = [
            self.data_reader.get_col_distribution(
                col_name, 
                normalize=True,
                exclude=0,
                return_dict=True
            ) for col_name in cols]

        data = pd.DataFrame(data, index=cols)
        data["index"] = cols
        self.ppt_generator.add_bar_chart(
            data,
            "index",
            ["1.0"],
            has_legend=False,
            to_percentage=True
        )

    def _process_gba_major_or_job_page(self, title: str, cols: str, target_major_or_job: list[str | int]):
        self.ppt_generator.create_blank_slide(title)

        stem_data = self.data_reader.get_combined_distribution(
            columns=cols,
            filtered_column="大灣區了解",
            filter_value=1
        )

        no_stem_data = self.data_reader.get_combined_distribution(
            columns=cols,
            filtered_column="大灣區了解",
            filter_value=3
        )

        # 58: Math
        stem_filtered = stem_data[stem_data[cols[0]].isin(target_major_or_job)]
        no_stem_filtered = no_stem_data[no_stem_data[cols[0]].isin(target_major_or_job)]

        # Merge the two DataFrames for bar chart
        merged_df = pd.merge(
            stem_filtered,
            no_stem_filtered,
            on=cols[0],
            suffixes=("_gba", "_no_gba")
        )

        self.ppt_generator.add_bar_chart(
            merged_df.rename(columns={"distribution_gba": "對大灣區了解", "distribution_no_gba": "對大灣區不了解"}),
            category_column=cols[0],
            value_columns=["對大灣區了解", "對大灣區不了解"],
            to_percentage=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=6, y=2, cx=4, cy=4
        )

        stem_data["distribution"] = stem_data["distribution"].apply(lambda x: f"{x:.1%}")
        no_stem_data["distribution"] = no_stem_data["distribution"].apply(lambda x: f"{x:.1%}")

        self.ppt_generator.add_table(
            stem_data.head(10).rename(columns={cols[0]: "對大灣區了解", "distribution": "百分比"}),
            rows=11, cols=2,
            index=False,
            font_size=12,
            x=0.5, y=2, cx=2.7, cy=4.5,
        )

        self.ppt_generator.add_table(
            no_stem_data.head(10).rename(columns={cols[0]: "對大灣區了解", "distribution": "百分比"}),
            rows=11, cols=2,
            index=False,
            font_size=12,
            x=3, y=2, cx=2.7, cy=4.5,
        )


    def _process_gba_page3(self):
        self.ppt_generator.create_blank_slide("考生對大灣區擇業的影響因素")

        cols = [
            "個人興趣及性格", "個人能力", "晉升機會",
            "工作性質", "行業前景", "工作環境",
            "工作量", "薪水福利", "生活成本", "國家貢獻"
        ]

        data = [
            self.data_reader.get_col_distribution(
                col_name + "_gba", 
                normalize=True,
                return_dict=True
            ) for col_name in cols]

        data = pd.DataFrame(data).sort_values(by="1.0", ascending=True)
        data["index"] = cols

        self.ppt_generator.add_bar_chart(
            data,
            "index",
            ["1.0"],
            has_legend=False,
            to_percentage=True,
            horizontal=True,
            title="考生對大灣區擇業的影響因素",
            x=1, y=1.5, cx=8, cy=5.5
        )

    def process_gba_pages(self):
        self._process_page1()
        self._process_page2()

        self._process_gba_major_or_job_page("對大灣區政策了解不同程度下受歡迎科目", ["希望修讀", "希望修讀_A", "希望修讀_B"], [25, 61])
        self._process_gba_major_or_job_page("對大灣區政策了解程度不同下不受歡迎科目", ["不希望修讀", "不希望修讀_A", "不希望修讀_B"], [58])
        self._process_gba_major_or_job_page("對大灣區政策了解程度不同下受歡職業", ["希望從事", "希望從事_A", "希望從事_B"], [24, 48])
        # self._process_major_or_job_page("不受歡迎職業", ["不希望從事", "不希望從事_A", "不希望從事_B"], [24])

        self._process_gba_page3()
        
