###
# Excel PAR
# 12. 전처리_T/B
# 231019. BSPL 구분자로 전기액 조절하는 코드 추가
###

##################################################################

#전역부

from mylib import myFileDialog as myfd
import pandas as pd
import clipboard

##################################################################

#1. TB 가공 >> Excel

##################################################################

#2. TB Import
df = pd.read_excel(myfd.askopenfilename(), sheet_name="IMPORT")

#전처리 - NaN to 0
df = df.fillna(0)

# 이 단계에서 df는 rawdata

##################################################################

#2) TB 컬럼매핑 >> 엑셀로 수행

#헤더 추출
tmp = df.columns.to_list()
tmp = map(str,tmp)
tmp = list(tmp)
tmp = ','.join(tmp)
clipboard.copy(tmp)

##################################################################

#3) GL DF 컬럼 전처리
dfMap = pd.read_excel(myfd.askopenfilename(), sheet_name="MAP_TB")

## a. MAP 대상 먼저 전처리
dfMapMap = dfMap[dfMap['방법'] == 'MAP']
dfTB = pd.DataFrame()
for i in range(0, dfMapMap.shape[0]):    
    try:
        dfTB[dfMapMap.iloc[i]["tobe"]] = df[dfMapMap.iloc[i]["asis"]]
    except:
        dfTB[dfMapMap.iloc[i]["tobe"]] = df[str(dfMapMap.iloc[i]["asis"])]

## b. KEYIN 대상 추가 전처리
dfMapKeyin = dfMap[dfMap['방법'] == 'KEYIN']

for i in range(0, dfMapKeyin.shape[0]):        
    dfTB[dfMapKeyin.iloc[i]["tobe"]] = dfMapKeyin.iloc[i]["asis"]

## c . 수기

import numpy as np

dfTB["CY"] = dfTB["당기말"]
dfTB["PY"] = np.where(dfTB["BSPL"] == "BS", dfTB["전기말"], dfTB["전년동기말"]) #조건식으로 브로드캐스팅
dfTB["PY1"] = dfTB["전전기말"]

dfTB["Company code"] = dfTB["계정과목코드"] + "_" + dfTB["계정과목명"] # Company Code

## 4) FSLine 추가

from mylib import myFileDialog as myfd
import pandas as pd

#계정과목 매핑을 읽는다.
dfAcc = pd.read_excel(myfd.askopenfilename())
dfTBShape = dfTB.shape
dfTBJoin = dfTB.merge(dfAcc,how='left',left_on='계정과목코드',right_on='DetailCode')
dfTBJoin[dfTBJoin["FSName"].isna()] #Empty여야 함
dfTB = dfTBJoin

##################################################################

# EXPORT
dfTB.to_csv("./raw/dfTB.tsv", sep="\t", index=None)