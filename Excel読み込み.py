import openpyxl

#import
import pandas as pd
import glob


#makinframe
df = pd.DataFrame(columns=["発注日","発注コード","発注先名","納期","注文NO","識別","商品名","","数量","単位","単価","値段","発注NO","納品番号","納品書"])

#getlist
path_str = "*.xlsx"
paths = glob.glob(path_str)

for file_path in paths:

     wb = {}
     wb = openpyxl.load_workbook(file_path)
     sheet = wb["Sheet1"]

     for row in sheet.iter_rows(min_row=2,max_col=30):

          values1 = []
          for col in row:
               values1.append(col.value)
               print(values1)
