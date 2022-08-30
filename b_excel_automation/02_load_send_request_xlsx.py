import pandas as pd

# 파일 불러와서 첫 행을 컬럼으로 지정 후 0, 1행 지움
df = pd.read_excel('sendRequest.xlsx', engine='openpyxl')
processed_df = df.rename(columns=df.iloc[1])
processed_df = processed_df.drop([processed_df.index[0], processed_df.index[1]])
processed_df = processed_df.reset_index(drop=True)
print(processed_df)


