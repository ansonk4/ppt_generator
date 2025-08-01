import pandas as pd
from data_reader import DataReader
from ppt_generator import PptGenerator
from pptx.enum.chart import XL_LEGEND_POSITION
from datetime import datetime

class AfterDSEProcessor:
    def __init__(self, data_reader: DataReader, ppt_generator: PptGenerator):
        self.data_reader = data_reader
        self.ppt_generator = ppt_generator
    

    def _process_page1(self):
        self.ppt_generator.create_blank_slide("考生DSE後第一階段計劃")

        plan_dis = self.data_reader.get_binary_distribution(
            ["大學", "副學士", "文憑", "高級文憑", "工作", "工作假期"],
            return_dict=False
        )

        self.ppt_generator.add_pie_chart(
            plan_dis, 'category', 'distribution',
            title='第一階段計劃',
            to_percent=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=5, y=2, cx=5, cy=5,
        )
        plan_dis.set_index('category', inplace=True)
        current_year = str(datetime.now().year)
        plan_dis[current_year] = (plan_dis['distribution'] * 100).round(1).astype(str) + '%'
        plan_dis.drop(columns=['distribution'], inplace=True)

        self.ppt_generator.add_table(
            plan_dis, rows=8, cols=8,
            x=0.5, y=2, cx=5, cy=3.5,
            font_size=12,
        )  
    
    def _process_page2(self):
        self.ppt_generator.create_blank_slide("考生DSE後第二階段計劃")
        col = "試後計劃"
        plan_dis = self.data_reader.get_col_distribution(
            col,
            normalize=True,
            return_dict=False,
        )
        

        self.ppt_generator.add_pie_chart(
            plan_dis, col, 'distribution',
            title='第二階段計劃',
            to_percent=True,
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            x=5, y=2, cx=5, cy=5,
        )

        plan_dis.set_index(col, inplace=True)
        current_year = str(datetime.now().year)
        plan_dis[current_year] = (plan_dis['distribution'] * 100).round(1).astype(str) + '%'
        plan_dis.drop(columns=['distribution'], inplace=True)

        self.ppt_generator.add_table(
            plan_dis, rows=5, cols=8,
            x=0.5, y=2, cx=5, cy=3.5,
            font_size=12,
        )

    def _process_page34(self):
        self.ppt_generator.create_blank_slide("考生升學地方")

        location_dis = []
        for i in range(1, 5):
            location_dis.append(
                self.data_reader.get_binary_distribution(
                    ["香港", "內地", "亞洲", "歐美澳"],
                    value=i,
                    return_dict=False
                ).rename(columns={'distribution': str(datetime.now().year)})
            )

        self.ppt_generator.add_bar_chart(
            location_dis[0], 'category', [str(datetime.now().year)],
            title='首位選擇',
            to_percentage=True,
            x=1, y=2, cx=8, cy=4
        )

        self.ppt_generator.create_blank_slide("考生升學地方")

        self.ppt_generator.add_bar_chart(
            location_dis[1], 'category', [str(datetime.now().year)],
            title='第二位選擇',
            to_percentage=True,
            hide_y_axis=True,
            x=1, y=1, cx=8, cy=3
        )

        self.ppt_generator.add_bar_chart(
            location_dis[2], 'category', [str(datetime.now().year)],
            title='第三位選擇',
            to_percentage=True,
            hide_y_axis=True,
            x=0.5, y=4, cx=4.5, cy=3
        )
        self.ppt_generator.add_bar_chart(
            location_dis[3], 'category', [str(datetime.now().year)],
            title='第四位選擇',
            to_percentage=True,
            hide_y_axis=True,
            x=5.5, y=4, cx=4.5, cy=3
        )

    def _process_page56(self):
        self.ppt_generator.create_blank_slide("考生升學地方 (學校Banding) ")

        bandings_location_dis = [self.data_reader.get_binary_distribution(
            ["香港", "內地", "亞洲", "歐美澳"],
            filter_column="Banding",
            filter_value=banding,
            return_dict=False
        ).set_index('category').rename(columns={'distribution': f'Banding {banding}'}) for banding in [1, 2, 3]]

        bandings_location_df = pd.concat(bandings_location_dis, axis=1)
        bandings_location_df = bandings_location_df.reset_index().rename(columns={"index": "category"})

        self.ppt_generator.add_bar_chart(
            bandings_location_df,
            category_column='category',
            value_columns=['Banding 1', 'Banding 2', 'Banding 3'],
            title='2025 考生升學地方',
            to_percentage=True,
            x=1, y=2, cx=8, cy=4
        )

        self.ppt_generator.create_blank_slide("考生希望升讀的大學")

        university_dis = self.data_reader.get_binary_distribution(
            [
            "浸會大學", "中文大學", "城市大學", "教育大學", "恒生大學", "香港大學",
            "嶺南大學", "都會大學", "理工大學", "聖方濟各大學", "樹仁大學", "科技大學", "自資學院"
            ],
            return_dict=False
        )
        
        self.ppt_generator.add_bar_chart(
            university_dis,
            category_column='category',
            value_columns=['distribution'],
            title='2025考生希望升讀的大學',
            to_percentage=True,
            has_legend=False,
            x=1, y=1.5, cx=8, cy=5.5
        )


    def _process_page7(self):
        self.ppt_generator.create_blank_slide("受歡迎主修科目 (按考生希望升讀的大學)")


    
    def _process_page8(self):
        self.ppt_generator.create_blank_slide("考生接收升學及就業資訊活動和成效")

        cols = [
            "大學入學講座",
            "升學展覽",
            "職業博覽",
            "生涯規劃",
            "團體師友",
            "工作影子",
        ]
        data_A, data_B = [], []


        data_A = [self.data_reader.get_col_distribution(col + "_A", normalize=True, return_dict=True) for col in cols]
        data_B = [
            self.data_reader.get_col_distribution(
                col + "_B",
                normalize=True,
                return_dict=True, 
                filter_column=col + "_A", 
                filter_value=1
            ) for col in cols
        ]
        
        df_A = pd.DataFrame(data_A)
        df_A['category'] = cols
        df_A = df_A.rename(columns={'1.0': 'distribution'})

        df_B = pd.DataFrame(data_B)
        # Drop 0 and Normalize each row
        if '0.0' in df_B.columns:
            df_B = df_B.drop(columns=['0.0'])
        value_cols = [col for col in df_B.columns]
        df_B[value_cols] = df_B[value_cols].div(df_B[value_cols].sum(axis=1), axis=0)
        df_B['category'] = cols

        self.ppt_generator.add_bar_chart(
            df_A,
            category_column='category',
            value_columns=['distribution'],
            title='接收升學及就業資訊渠道',
            to_percentage=True,
            has_legend=False,
            x=0.5, y=2, cx=4.5, cy=5
        )

        self.ppt_generator.add_stacked_bar(
            df_B,
            category_column='category',
            value_columns=df_B.columns[df_B.columns != 'category'].tolist(),
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            title='升學及就業資訊活動成效',
            x=5, y=2, cx=5, cy=4
        )


    def _process_page9(self):
        self.ppt_generator.create_blank_slide("影響考生選科因素")

        cols = ["學科知識", "院校因素", "大學學費", "助學金", "主要行業", "朋輩老師", 
            "家庭因素", "預期收入", "DSE成績", "高中選修科目"]

        data = [self.data_reader.get_col_distribution(col, normalize=True, return_dict=True) for col in cols]
        
        df = pd.DataFrame(data)
        df['category'] = cols

        self.ppt_generator.add_stacked_bar(
            df,
            category_column='category',
            value_columns=df.columns[df.columns != 'category'].tolist(),
            legend_position=XL_LEGEND_POSITION.BOTTOM,
            title='影響考生選科因素',
            x=0.5, y=1.5, cx=9, cy=5
        )

    def process_after_dse_pages(self):
        self._process_page1()
        self._process_page2()
        self._process_page34()
        self._process_page56()
        self._process_page7()
        self._process_page8()
        self._process_page9()



if __name__ == "__main__":
    # Example usage
    data_reader = DataReader("data/data2.xlsx")
    ppt_generator = PptGenerator()
    processor = AfterDSEProcessor(data_reader, ppt_generator)

    processor._process_page1()
