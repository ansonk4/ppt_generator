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


    def process_gba_pages(self):
        self._process_page1()
        self._process_page2()
    

    
        
