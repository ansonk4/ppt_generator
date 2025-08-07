import pandas as pd
import numpy as np
import yaml

class DataReader:
    """Class to read data from an Excel file"""
    def __init__(self, file_path: str):
        self.file_path = file_path

        try:
            df = pd.read_excel(self.file_path)
            # df = df.apply(lambda x: pd.to_numeric(x, errors='coerce'))
            df = df.replace(999, np.nan)
            df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
            # Replace any cell that is empty or contains only spaces with np.nan
            df = df.map(lambda x: np.nan if isinstance(x, str) and x.strip() == "" else x)

            self.data = df
            print(f"Data read successfully from {self.file_path}")
        except Exception as e:
            raise ValueError("Failed to read data from Excel file.")

    
    def get_col_distribution(
        self, 
        column_name: str,
        filter_column: str | None = None,
        filter_value: str | int | None = None,
        normalize: bool = True,
        exclude: float | int | None = None,
        return_dict: bool = False
    ) -> pd.DataFrame | dict:
        """Get the distribution of a specified column"""
        data = self.data
        if filter_column is not None and filter_value is not None:
            data = data[data[filter_column] == filter_value]

        if exclude is not None:
            data = data[data[column_name] != exclude]
        
        if column_name not in data.columns:
            print(f"Column {column_name} does not exist in the data.")
            return {}

        data = data.dropna(subset=[column_name])

        distribution = data[column_name].value_counts(normalize=normalize).to_dict()
        distribution = {str(float(k)) if isinstance(k, int) else str(k): v for k, v in distribution.items()}

        if return_dict:
            return distribution

        return pd.DataFrame(distribution.items(), columns=[column_name, 'distribution'])
        

    def get_binary_distribution(
        self, 
        columns: list[str], 
        value: int = 1, 
        unique: bool = False,
        filter_column: str | None = None,
        filter_value: str | int | None = None,
        return_dict: bool = True
    ) -> dict[str, float] | pd.DataFrame:
        result = {}

        data = self.data
        if filter_column is not None and filter_value is not None:
            data = self.data[self.data[filter_column] == filter_value]

        data = data.dropna(subset=columns)

        # Drop rows where more than one target value exists in the specified columns
        if unique:
            target_mask = data[columns].apply(lambda row: (pd.to_numeric(row, errors='coerce') == value).sum(), axis=1)
            data = data[target_mask <= 1]

        if len(data) == 0:
            print("No valid rows found in the DataFrame.")
            return result
            
        for col in columns:
            if col not in data.columns:
                print(f"Warning: Column '{col}' not found in DataFrame")
                result[col] = 0.0
            else:
                # Count the number of 1s in each column, handle non-numeric values as 0
                numeric_col = pd.to_numeric(data[col], errors='coerce').fillna(0)
                count = int((numeric_col == value).sum())
                result[col] = count / len(data)

        if unique:
            # Normalize so the sum of result is 100
            total = sum(result.values())
            result = {k: v / total for k, v in result.items()}

        if return_dict:
            return result       

        return pd.DataFrame(result.items(), columns=['category', 'distribution'])


    def get_combined_distribution(
        self,
        columns: list[str],
        filtered_column: str | None = None,
        filter_value: str | int | None = None,
        return_dict: bool = False,
    ) -> pd.DataFrame:
        """
        Get the normalized combined distribution of multiple specified columns.
        Optionally filter the data by a column and value.
        """
        result = {}

        data = self.data
        if filtered_column is not None and filter_value is not None:
            data = data[data[filtered_column] == filter_value]

        total_count = len(data)

        for col in columns:
            if col in data.columns:
                distribution = data[col].value_counts().to_dict()
                for key, value in distribution.items():
                    result[key] = result.get(key, 0) + value
            else:
                print(f"Column {col} does not exist in the data.")

        if total_count > 0:
            # Normalize the result
            result = {k: v / total_count for k, v in result.items()}
            result = dict(sorted(result.items(), key=lambda item: item[1], reverse=True))

        if return_dict:
            return result

        return pd.DataFrame(result.items(), columns=[columns[0], 'distribution'])


if __name__ == "__main__":

    # Example usage
    reader = DataReader("data/data2.xlsx")
    plan_dis = reader.get_binary_distribution(
        ["大學", "副學士", "文憑", "高級文憑", "工作", "工作假期"]
    )
    print(plan_dis)