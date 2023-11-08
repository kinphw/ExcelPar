import pandas as pd
import numpy as np

# 금호석유화학
########################################################################################################################

path = r'C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_금호석유화학\01_PBC가공'
path = path.replace('\\','/')
path = path + '/' + 'rawGLPY.parquet'
dfPYraw = pd.read_parquet(path)

path = r'C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_금호석유화학\01_PBC가공'
path = path.replace('\\','/')
path = path + '/' + 'rawGLCY.parquet'
dfCYraw = pd.read_parquet(path)

def Doit(df:pd.DataFrame, fileName:str):    
    c1 = df.columns[2]
    c2 = df.columns[3]

    # 일단 숫자로 바꾼다.
    #dfCYraw[c2].astype('int64')
    df[c2] = pd.to_numeric(df[c2].replace("[,]","",regex=True))

    # 차대변경
    
    df[c2] = np.where(df[c1] == 'S', df[c2], df[c2] * -1)

    #df.groupby('계정')['현지통화금액'].sum().to_excel("PY_pivot.xlsx")

    df.to_parquet(fileName)

Doit(dfPYraw, 'dfPYraw_re.parquet')
Doit(dfCYraw, 'dfCYraw_re.parquet')

########################################################################################################################

