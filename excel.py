import pandas as pd

all_sheets = pd.read_excel('ttt.xls' , sheet_name=None)
for sheet_name, df in all_sheets.items():
    print(f"Sheet Name: {sheet_name}")
    print(df)
    print("\n")
# selected_c=df['First Name']
# selected_Id=df['Id']
# for i in range(0,length):
#     if(selected_Id[i]>3000):
#        selected_c[i] = 'Yes'
    
# df.to_excel('ttt.xlsx', index=False)

