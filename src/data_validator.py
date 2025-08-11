import pandas as pd
import numpy as np


class DataValidator:
    """Class to validate data in a DataFrame"""
    def __init__(self, data: pd.DataFrame):
        self.data = data

    def validate_column(self) -> list[str]:
        """Validate if the column exists and has the expected type"""
        column_required = set([
            #after dse
            "大學", "副學士", "文憑", "高級文憑", "工作", "工作假期", "其他", 
            "試後計劃",
            "香港", "內地", "亞洲", "歐美澳",
            "浸會大學", "中文大學", "城市大學", "教育大學", "恒生大學", "香港大學",
            "嶺南大學", "都會大學", "理工大學", "聖方濟各大學", "樹仁大學", "科技大學", "自資學院",
            #background
            "希望修讀", "希望修讀_A", "希望修讀_B",
            "不希望修讀", "不希望修讀_A", "不希望修讀_B",
            "大學入學講座_A","升學展覽_A","職業博覽_A","生涯規劃_A","團體師友_A","工作影子_A",
            "大學入學講座_B","升學展覽_B","職業博覽_B","生涯規劃_B","團體師友_B","工作影子_B",
            "學科知識", "院校因素", "大學學費", "助學金", "主要行業", "朋輩老師", 
            "家庭因素", "預期收入", "DSE成績", "高中選修科目",
            "中文成績", "英文成績", "數學成績",
            "性別", "Banding", "學校編號",
            "父母教育程度", "高中選修學科",
            #GBA
            "大灣區了解",
            "公社科","內地考察","政府資訊","新聞媒體","網上資訊","內地交流","校內講座","朋輩及老師",
            "個人興趣及性格_gba", "個人能力_gba", "晉升機會_gba",
            "工作性質_gba", "行業前景_gba", "工作環境_gba",
            "工作量_gba", "薪水福利_gba", "生活成本_gba", "國家貢獻_gba",
            #job
            "工作地方",
            "個人能力_B", "個人興趣性格_B", "成就感_B", "家庭因素_B", "人際關係_B",
            "工作性質_B", "工作模式_B", "工作量_B", "工作環境_B", "薪水及褔利_B",
            "晉升機會_B", "發展前景_B", "社會貢獻_B", "社會地位_B",
            "希望從事", "希望從事_A", "希望從事_B",
            "不希望從事", "不希望從事_A", "不希望從事_B",
            "從事相關工作",
            #Stem
            "參加STEM", "STEM影響職業選擇程度", "領導能力", "團隊合作", "創新思維", "科學知識", "解難能力"
        ])

        columns_present = set(self.data.columns)
        if not column_required.issubset(columns_present):
            missing_columns = column_required - columns_present
            return list(missing_columns)
        else:
            return []
    
    def validate_cols(self) -> list[dict]:
        """Validate specific columns for acceptable values"""

        validation_result = []
        validation_result.append(self._validate_col("性別", ["男", "女"]))
        validation_result.append(self._validate_col("Banding", ["Band 1", "Band 2", "Band 3"]))
        validation_result.append(self._validate_col(["大學入學講座_A","升學展覽_A","職業博覽_A","生涯規劃_A","團體師友_A","工作影子_A"], ["有", "沒有"]))
        validation_result.append(self._validate_col("大灣區了解", ["完全不了解", "不太了解", "了解", "非常了解"]))
        validation_result.append(self._validate_col(["公社科","內地考察","政府資訊","新聞媒體","網上資訊","內地交流","校內講座","朋輩及老師"], ["曾經 / 希望參與", "沒有 / 不會參與"]))
        validation_result.append(self._validate_col("工作地方", ["香港", "內地", "國外 - 亞洲", "國外 - 歐美澳"]))
        validation_result.append(self._validate_col(
            ["個人能力_B", "個人興趣性格_B", "成就感_B", "家庭因素_B", "人際關係_B",
            "工作性質_B", "工作模式_B", "工作量_B", "工作環境_B", "薪水及褔利_B",
            "晉升機會_B", "發展前景_B", "社會貢獻_B", "社會地位_B"],
            ["十分重要", "重要", "不太重要", "不重要"]
        ))
        validation_result.append(self._validate_col("參加STEM", ["有", "沒有"]))
        validation_result.append(self._validate_col(["中文成績", "英文成績", "數學成績"], ["< 25 分", "25-49 分", "50-75 分", "> 75 分"]))
        validation_result.append(self._validate_col("從事相關工作", ["絕對不會", "可能不會", "不確定", "可能會", "絕對會"]))
        validation_result.append(self._validate_col([
            "浸會大學", "中文大學", "城市大學", "教育大學", "恒生大學", "香港大學",
            "嶺南大學", "都會大學", "理工大學", "聖方濟各大學", "樹仁大學", "科技大學", "自資學院"],
            [1, 0]
        ))

        return validation_result

    def _validate_col(self, column: str | list[str], acceptable_values: list[str]) -> dict[list[tuple[int, str]]]:
        """Validate if the column exists and contains acceptable values"""
        column = [column] if isinstance(column, str) else column
        acceptable_values = acceptable_values + [np.nan]
        result = {}
        for col in column:
            invalid_mask = ~self.data[col].isin(acceptable_values)
            invalid_values = self.data[col][invalid_mask]
            if not invalid_values.empty:
                invalid_row_ids = self.data.index[invalid_mask].tolist()
                result[col] = [(row_id, value) for row_id, value in zip(invalid_row_ids, invalid_values)]
                
        if result:
            result["acceptable_values"] = acceptable_values

        return result
