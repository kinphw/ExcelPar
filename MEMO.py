# 롯데하이마트
########################################################################################################################

from ExcelPar.mylib import myFileDialog as myfd
import pandas as pd
from ExcelPar.mylib.ForSAP import ForSAP
from ExcelPar.ConcatTextFiles import Object2String
import numpy as np

path = myfd.askopenfilename()
df = pd.read_csv(path,sep='\t', chunksize=1000000, encoding='cp949')
dfCon = pd.DataFrame()
li = ['전표번호','전기일자','전표적요','차대변구분자','전표금액','계정과목코드']
for cnt, chunk in enumerate(df):
     print(cnt,"회")
     dfCon = pd.concat([dfCon,chunk[li]])

Object2String(dfCon)
dfCon.to_parquet("GLCY_tmp.parquet")

ForSAP.ColumnStrToInt(dfCon, '전표금액')
ForSAP.Multiple100(dfCon, '전표금액')
dfCon['전표금액'] = np.where(dfCon['차대변구분자'] == 'D', dfCon['전표금액'], dfCon['전표금액'] * -1)

#계정과목코드 클렌징 (00이 섞여있음)
dfCon['계정과목코드'] = pd.to_numeric(dfCon['계정과목코드'], errors='coerce').fillna(0).astype('float64').astype('int64').astype('str')
dfCon['계정과목코드'] = dfCon['계정과목코드'].replace(" ","")

dfCon.groupby('계정과목코드')['전표금액'].sum().to_excel("testCY.xlsx")

dfCon.to_parquet("rawGLPY_F.parquet")

# 재처리

df = pd.read_parquet(path)



#삭제
df[df['전기일자'] == 'BLDAT']

df = df.reset_index()
df = df.drop(df[df['전기일자'] == 'BLDAT'].index)

df.to_parquet("rawGLCY_F_RE.parquet")

