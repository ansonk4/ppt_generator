import pandas as pd
import numpy as np

df = pd.read_excel("data/2025data.xlsx").iloc[0:500]
df = df.drop(columns=["Email "])
cols = ["公社科","內地考察","政府資訊","新聞媒體","網上資訊","內地交流","校內講座","朋輩及老師"]
df.to_excel("sample_data/sample.xlsx", index=False)