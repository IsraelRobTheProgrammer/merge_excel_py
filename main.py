import pandas as pd

siwes_df = pd.read_excel("SIWES 2.xlsx")
exam_df = pd.read_excel("exam.xlsx")
ca_df = pd.read_excel("CA.xlsx")

merged_df = siwes_df.merge(exam_df, on="MAT NO").merge(ca_df, on="MAT NO")
merged_df = merged_df.drop(columns=["NAME_y"]).rename(columns={"NAME_x": "NAME"})
merged_df.to_excel("merged_res.xlsx", index=False)
