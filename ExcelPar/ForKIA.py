# 가공구문 FOR KIA...
# 폴더들을 순환하면서 전표금액 가공

import os
import csv
import sys
import numpy as np
import gc

import pandas as pd
import dask.dataframe as dd
from dask.diagnostics import ProgressBar

def DoKIA(i:str):
    #1. 읽는다.
    #tmp = r'C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\22_0'
    #tmp = tmp.replace('\\','/')
    #tmp = tmp + '/' + '*.parquet'
    tmp = '*.parquet'
    dfCY = dd.read_parquet(tmp)
    dfConCY = dfCY.compute()    
    print("LOADED")

    #2. 금액 전처리
    dfTmp = dfConCY['현지통화금액']
    dfTmp = pd.DataFrame(dfTmp)
    # fillna
    #dfTmp['현지통화금액'] = dfTmp['현지통화금액'].fillna('0')
    # trim
    #dfTmp['현지통화금액'] = dfTmp['현지통화금액'].apply(str.strip)
    # 숫자로
    dfTmp['현지통화금액'] = pd.to_numeric(dfTmp['현지통화금액'], errors='coerce')
    # fillna
    dfTmp['현지통화금액'] = dfTmp['현지통화금액'].fillna(0)
    # *100
    dfTmp['현지통화금액'] = dfTmp['현지통화금액']*100
    dfConCY['현지통화금액']=dfTmp['현지통화금액']    

    #3. 차대 마이너스
    import numpy as np
    dfTmp = dfConCY[['차대변지시자','현지통화금액']]
    dfTmp['차대변지시자'] = dfTmp['차대변지시자'].fillna('S')
    dfTmp['현지통화금액'] = np.where(dfTmp['차대변지시자'] =='S', dfTmp['현지통화금액'], dfTmp['현지통화금액'] * -1)
    dfConCY['현지통화금액'] = dfTmp['현지통화금액']

    print("AMOUNT DONE")
    #4. 불필요한 행 삭제
    dfConCY.drop(columns=['차대변지시자'],inplace=True)

    print("차대변 DELETED")

    #5. 계정과목 CHECK - unsigned int 로
    dfConCY['계정코드'] = pd.to_numeric(dfConCY['계정코드'],downcast='unsigned',errors='coerce')
    print("ACCT DONE")

    #5. 메모리관리
    del dfTmp
    dfTmp = pd.DataFrame()

    # Save
    fileName = i+'.parquet'
    dfConCY.to_parquet(fileName)    
    print("Saved:",fileName)

def Run():

    ProgressBar().register()

    rootPath = r'C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사'
    rootPath = rootPath.replace('\\','/')
    
    li =['22_0','22_1','22_2','22_3','23_1','23_2','23_3']
    for i in li:
        os.chdir(rootPath+'/'+i)
        print(i)
        DoKIA(i)

if __name__=='__main__':
    Run()
    

