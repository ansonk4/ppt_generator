import data_reader

reader = data_reader.DataReader("data/data2.xlsx")
df = reader.get_combined_distribution(["希望修讀", "希望修讀_A", "希望修讀_B"]).head(5)
print(df)