import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl import Workbook
start=0

end=0
result_file="/Users/guqingxian/Desktop/学习/Data Mining in Excel/Results_Exercises/results.xlsx"
df_SA = pd.read_excel('/Users/guqingxian/Desktop/学习/Data Mining in Excel/Data Set/Stand-Alone.xlsx')
result=pd.read_excel(result_file,header=None,sheet_name='Total 30 Traces')

def search_cells_by_value(search_value):
    mask = df_SA['No.'] == search_value
    print(mask)
    matching_indices = df_SA.index[mask].tolist()
    return matching_indices


No=search_cells_by_value(1)
No.append(df_SA.shape[0])
# print(len(No))
# print(type(No[1]))
# print(No)
for i in range(30):
    result.loc[i+2,0]=i+1
    Sum=df_SA['Time'].iloc[No[i]:No[i+1]].sum()
    result.loc[i+2,1]=int(Sum)
    print(int(Sum))


result.to_excel(result_file,index=False,header=None,sheet_name='Total 30 Traces')





# for i in range(30):
#
#     No2=find_row_index_SA('1')
#     Sum=df_SA['Time'].iloc[No1:No2].sum()
#     No1=No2
#     # print(Sum)


