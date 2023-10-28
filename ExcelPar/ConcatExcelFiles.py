#경로 내 엑셀파일을 합치는 코드

import ExcelPar.mylib.myFileDialog as myfd
import os
import glob
import pandas as pd

path = myfd.askdirectory()
list = glob.glob(path + "/" + "*.xls*")

len(list)

dfCon = pd.DataFrame()

for i in list:
    print(i,"번째 시작:")
    df = pd.read_excel(i)
    df.columns = [i.strip() for i in df.columns] #추가코드. strip
    dfCon = pd.concat([dfCon,df])
    print(i,"번째 종료:")

dfCon.shape
dfCon.columns
dfCon.columns

#검증
dfTmp = dfCon[['G/L 계정','금액(현지 통화)']]
dfTmp.loc[:, '금액(현지 통화)'] = pd.to_numeric(dfCon['금액(현지 통화)'], errors='coerce').fillna(0)
dfTmp.groupby('G/L 계정').sum().to_excel('test_PY2.xlsx')

filename = "rawGLPY.tsv"
dfCon.to_csv(filename,sep="\t",index=None)