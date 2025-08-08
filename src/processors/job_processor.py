import pandas as pd 
from data_reader import DataReader
from ppt_generator import PptGenerator
from pptx.enum.chart import XL_LEGEND_POSITION
from datetime import datetime

class JobProcessor:
    def __init__(self, data_reader: DataReader, ppt_generator: PptGenerator):
        self.data_reader = data_reader
        self.ppt_generator = ppt_generator


    def _process_page1(self):
        self.ppt_generator.create_blank_slide("未來工作地點")
        col = "工作地方"
        location = self.data_reader.get_col_distribution(col, normalize=True, return_dict=False)

        self.ppt_generator.add_pie_chart(
            location, col, 'distribution',
            to_percent=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=5, y=2, cx=5, cy=5,
        )

        self.ppt_generator.add_bar_chart(
            location, col, ['distribution'],
            title='未來工作地點分佈',
            to_percentage=True,
            x=0.5, y=2, cx=5, cy=5,
        )

    def _process_page_filtered_by_location(self, location: str):
        self.ppt_generator.create_blank_slide(f"未來工作地點：{location}")

        filter_map = {
            "香港": "香港",
            "內地": "內地",
            "亞洲": "國外 - 亞洲",
            "歐美澳": "國外 - 歐美澳"
        }
        gender = self.data_reader.get_col_distribution(
            "性別", normalize=True, return_dict=False,
            filter_column="工作地方", filter_value=filter_map[location]
        )

        self.ppt_generator.add_donut_chart(
            gender, "性別", "distribution",
            to_percent=True,
            title=f"選擇在{location}工作的DSE考生",
            has_legend=False,
            small_title=True,
            has_data_labels=True,
            x=0.3, y=1.2, cx=3, cy=3
        )   

        gba = self.data_reader.get_col_distribution(
            "大灣區了解", normalize=True, return_dict=False,
            filter_column="工作地方", filter_value=filter_map[location]
        )
        self.ppt_generator.add_donut_chart(
            gba, "大灣區了解", "distribution",
            to_percent=True,
            title=f"選擇在{location}工作的DSE考生對大灣區政策的了解程度",
            has_legend=False,
            has_data_labels=True,
            small_title=True,
            x=0.3, y=4, cx=3.5, cy=3.5
        )

        gba = self.data_reader.get_col_distribution(
            "高中選修學科", normalize=True, return_dict=False,
            filter_column="工作地方", filter_value=filter_map[location]
        )
        self.ppt_generator.add_donut_chart(
            gba, "高中選修學科", "distribution",
            to_percent=True,
            title=f"選擇在{location}工作的DSE考生的高中選修學科",
            has_legend=False,
            has_data_labels=True,
            small_title=True,
            x=2.8, y=2.5, cx=3.5, cy=3.5
        )

        major = self.data_reader.get_combined_distribution(
            columns=["希望修讀", "希望修讀_A", "希望修讀_B"],
            filtered_column="工作地方",
            filter_value=filter_map[location]
        )[:5]

        self.ppt_generator.add_bar_chart(
            major,
            category_column="希望修讀",
            value_columns=["distribution"],
            title="選擇在香港工作的DSE考生希望修讀的科目",
            has_legend=False,
            to_percentage=True,
            small_title=True,
            x=6, y=1.2, cx=4, cy=3
        )
        
        job = self.data_reader.get_combined_distribution(
            columns=["希望從事", "希望從事_A", "希望從事_B"],
            filtered_column="工作地方",
            filter_value=filter_map[location]
        )[:5]

        self.ppt_generator.add_bar_chart(
            job,
            category_column="希望從事",
            value_columns=["distribution"],
            title=f"選擇在{location}工作的DSE考生希望從事的工作",
            has_legend=False,
            small_title=True,
            to_percentage=True,
            x=6, y=4, cx=4, cy=3
        )


    def _process_page2(self):
        self.ppt_generator.create_blank_slide("考生擇業條件")
        cols = [
            "個人能力", "個人興趣性格", "成就感", "家庭因素", "人際關係",
            "工作性質", "工作模式", "工作量", "工作環境", "薪水及褔利",
            "晉升機會", "發展前景", "社會貢獻", "社會地位"
        ]
        
        dis = [
            self.data_reader.get_col_distribution(col + "_B", normalize=True, return_dict=True) for col in cols
        ]
        dis = pd.DataFrame(dis, index=cols).reset_index()

        target_col = "十分重要"
        dis = dis.sort_values(by=target_col, ascending=True)

        self.ppt_generator.add_bar_chart(
            dis,
            category_column="index",
            value_columns=[target_col],
            title="2025考生擇業條件",
            has_legend=False,
            to_percentage=True,
            horizontal=True,
            y=1.2, cy=5.8
        )


    def _process_major_preference_page(self, title:str, cols: list[str]):
        self.ppt_generator.create_blank_slide(title)

        job_data = self.data_reader.get_combined_distribution(
            columns=cols,
            return_dict=False
        ).head(5)

        self.ppt_generator.add_bar_chart(
            job_data,
            category_column=cols[0],
            value_columns=["distribution"],
            has_legend=False,
            to_percentage=True,
            x=3.5, cx=6
        )

        jobs = job_data[cols[0]].tolist()
        distribution = job_data["distribution"].tolist()
        text = (
            f"第一位 {jobs[0]} {distribution[0]*100:.1f}%\n",
            f"第二位 {jobs[1]} {distribution[1]*100:.1f}%\n",
            f"第三位 {jobs[2]} {distribution[2]*100:.1f}%\n",
            f"第四位 {jobs[3]} {distribution[3]*100:.1f}%\n",
            f"第五位 {jobs[4]} {distribution[4]*100:.1f}%\n",
        )
        self.ppt_generator.add_textbox(
            "\n".join(text),
            x=0.5, cx=3, font_size=19
        )


    def _process_gender_major_preference_page(self, title: str, cols: list[str]):
        self.ppt_generator.create_blank_slide(title)

        male_data = self.data_reader.get_combined_distribution(
            cols,
            filtered_column="性別",
            filter_value="男"
        )[:5]

        female_data = self.data_reader.get_combined_distribution(
           cols,
           filtered_column="性別",
           filter_value="女"
       )[:5]

        self.ppt_generator.add_bar_chart(
            male_data,
            cols[0],
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
            cols[0],
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

        for i, text in enumerate(["1st", "2nd", "3rd", "4th", "5th"]):
            self.ppt_generator.add_textbox(
                text,
                x=4.7, y=3 + i * 0.75, cx=0.5, cy=0.5
            )
    
    def _process_least_popular_job(self):
        least_popular_job = self.data_reader.get_combined_distribution(
            ["不希望從事", "不希望從事_A", "不希望從事_B"],
        ).head(1)["不希望從事"].values[0]

        self.ppt_generator.create_blank_slide(f"最不受歡迎職業：{least_popular_job} (背景資料)")

        gender = self.data_reader.get_col_distribution(
            "性別",  
            ["不希望從事", "不希望從事_A", "不希望從事_B"],
            least_popular_job,
            normalize=True, 
        
        )
        self.ppt_generator.add_donut_chart(
            gender,
            "性別", "distribution",
            has_legend=False,
            to_percent=True,
            has_data_labels=True,
            x=0.0, y=1.7, cx=3, cy=3,
        )

        banding = self.data_reader.get_col_distribution(
            "Banding", 
            ["不希望從事", "不希望從事_A", "不希望從事_B"],
            least_popular_job,
            normalize=True,
        )
        self.ppt_generator.add_donut_chart(
            banding,
            "Banding", "distribution",
            has_legend=False,
            to_percent=True,
            has_data_labels=True,
            x=2.3, y=1.7, cx=3, cy=3,
        )

        elective = self.data_reader.get_col_distribution(
            "高中選修學科", 
            ["不希望從事", "不希望從事_A", "不希望從事_B"],
            least_popular_job,
            normalize=True,
        )
        self.ppt_generator.add_bar_chart(
            elective,
            "高中選修學科",
            ["distribution"],
            title="高中選修學科",
            to_percentage=True,
            has_legend=False,
            x=0.5, y=4.5, cx=4.5, cy=2.8,
        )

        edu_bg = self.data_reader.get_col_distribution(
            "父母教育程度", 
            ["不希望從事", "不希望從事_A", "不希望從事_B"],
            least_popular_job,
            normalize=True,
        )
        self.ppt_generator.add_bar_chart(
            edu_bg,
            "父母教育程度",
            ["distribution"],
            title="父母教育程度",
            to_percentage=True,
            has_legend=False,
            x=5, y=4, cx=5, cy=3.5,
        )

        cols = ["中文成績", "英文成績", "數學成績"]
        grade_data = [
            self.data_reader.get_col_distribution(
                col, normalize=True, return_dict=True,
                filter_column=["不希望從事", "不希望從事_A", "不希望從事_B"],
                filter_value=least_popular_job,
            ) for col in cols
        ]
        grade_data = pd.DataFrame(grade_data, index=cols).reset_index()
        self.ppt_generator.add_bar_chart(
            grade_data,
            category_column="index",
            value_columns=grade_data.columns[grade_data.columns != "index"].tolist(),
            to_percentage=True,
            font_size=12,
            has_legend=True,
            x=5, y=1.2, cx=4.5, cy=2.8
        )
    def _process_page5(self):

        self.ppt_generator.create_blank_slide("從事與大學主修科目相關工作的可能性")

        data = self.data_reader.get_col_distribution(
            "從事相關工作", normalize=True, return_dict=False
        )

        self.ppt_generator.add_donut_chart(
            data, "從事相關工作", "distribution",
            to_percent=True,
            title="從事與大學主修科目相關工作的可能性",
            has_legend=False,
            has_data_labels=True,
            x=1, y=2, cx=4, cy=4
        )

        self.ppt_generator.create_blank_slide("從事與大學主修科目相關工作：選擇絕對會")

        self.ppt_generator.add_donut_chart(
            data, "從事相關工作", "distribution",
            to_percent=True,
            title="從事與大學主修科目相關工作的可能性",
            has_legend=False,
            has_data_labels=True,
            x=0.5, y=2, cx=4, cy=4
        )

        major_data = self.data_reader.get_combined_distribution(
            columns=["希望修讀", "希望修讀_A", "希望修讀_B"],
            filtered_column="從事相關工作",
            filter_value="絕對會"
        ).head(10)

        job_data = self.data_reader.get_combined_distribution(
            columns=["希望從事", "希望從事_A", "希望從事_B"],
            filtered_column="從事相關工作",
            filter_value="絕對不會"
        ).head(10)

        major_data["百分比"] = major_data["distribution"].apply(lambda x: f"{round(x * 100, 1)}%")
        major_data = major_data.drop(columns=["distribution"])
        job_data["百分比"] = job_data["distribution"].apply(lambda x: f"{round(x * 100, 1)}%")
        job_data = job_data.drop(columns=["distribution"])

        self.ppt_generator.add_table(
            major_data,
            index=False,
            x=4.5, y=2.5, cx=2, cy=3.5,
            font_size=14
        )

        self.ppt_generator.add_table(
            job_data,
            index=False,
            x=7, y=2.5, cx=2, cy=3.5,
            font_size=14
        )
    
    def process_job_pages(self):
        self.ppt_generator.create_section_slide("考生未來工作取向")
        self._process_page1()

        self._process_page_filtered_by_location("香港")
        self._process_page_filtered_by_location("內地")
        self._process_page_filtered_by_location("亞洲")
        self._process_page_filtered_by_location("歐美澳")

        self._process_page2()

        self._process_major_preference_page("受歡迎職業", ["希望從事", "希望從事_A", "希望從事_B"])
        self.ppt_generator.create_blank_slide("受歡迎職業走勢")
        self._process_gender_major_preference_page("受男女歡迎職業排名", ["希望從事", "希望從事_A", "希望從事_B"])
        
        self._process_major_preference_page("不受歡迎職業", ["不希望從事", "不希望從事_A", "不希望從事_B"])
        self.ppt_generator.create_blank_slide("不受歡迎職業走勢")

        self._process_least_popular_job()

        self._process_gender_major_preference_page("不受男女歡迎職業排名", ["不希望從事", "不希望從事_A", "不希望從事_B"])

        self._process_page5()
