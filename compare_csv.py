import pandas as pd
data_path = 'data\\'
csv_file1 = data_path + 'Plexxis_Tax Details_Ck_Date_121523.csv'
csv_file2 = data_path + 'Zenefits_payroll_detail_Ck_Date_121523.csv'

df1 = pd.read_csv(csv_file1)
df2 = pd.read_csv(csv_file2)

common = df1.merge(df2,on=['Last Name','First Name'])
print(common)

