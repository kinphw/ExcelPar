###
# Excel PAR
# 11. 전처리_G/L
###

##################################################################

#전역부

from mylib import myFileDialog as myfd
import pandas as pd
import clipboard

##################################################################

#1) GL_raw 컬럼추출

tmp = myfd.askopenfilename()
df = pd.read_excel(tmp, sheet_name="GL")
a = df.columns.to_list()
tmp = ','.join(a)
clipboard.copy(tmp)

##################################################################

#2) GL 컬럼매핑 >> 엑셀로 수행

##################################################################

#3) GL DF 컬럼 전처리

##################################################################

## CY
df = pd.read_excel(myfd.askopenfilename(), sheet_name="GL")

dfMap = pd.read_excel("12_MAP.xlsx", sheet_name="MAP_GL")

## a. MAP 대상 먼저 전처리
dfMapMap = dfMap[dfMap['방법'] == 'MAP']
dfGL = pd.DataFrame()
for i in range(0, dfMapMap.shape[0]):    
    dfGL[dfMapMap.iloc[i]["tobe"]] = df[dfMapMap.iloc[i]["asis"]]

## b. KEYIN 대상 추가 전처리
dfMapKeyin = dfMap[dfMap['방법'] == 'KEYIN']

for i in range(0, dfMapKeyin.shape[0]):        
    dfGL[dfMapKeyin.iloc[i]["tobe"]] = dfMapKeyin.iloc[i]["asis"]

dfGL['연도'] = 'CY'

## b-1. 추가 수기처리

dfGL["회계월"] = dfGL["회계월"].apply(lambda x: x[-2:]) # 회계월 가공
dfGL["차변금액"] = dfGL["차변금액"].fillna(0)
dfGL["대변금액"] = dfGL["대변금액"].fillna(0)
dfGL["전표금액"] = dfGL["차변금액"] - dfGL["대변금액"] # 전표금액 = 차변잔액 - 대변잔액
dfGL["Company code"] = dfGL["계정과목코드"] + "_" + dfGL["계정과목명"] # Company Code

dfGLCY = dfGL

##################################################################

## PY
df = pd.read_excel(myfd.askopenfilename(), sheet_name="GL")

dfMap = pd.read_excel("12_MAP.xlsx", sheet_name="MAP_GL")

## a. MAP 대상 먼저 전처리
dfMapMap = dfMap[dfMap['방법'] == 'MAP']
dfGL = pd.DataFrame()
for i in range(0, dfMapMap.shape[0]):    
    dfGL[dfMapMap.iloc[i]["tobe"]] = df[dfMapMap.iloc[i]["asis"]]

## b. KEYIN 대상 추가 전처리
dfMapKeyin = dfMap[dfMap['방법'] == 'KEYIN']

for i in range(0, dfMapKeyin.shape[0]):        
    dfGL[dfMapKeyin.iloc[i]["tobe"]] = dfMapKeyin.iloc[i]["asis"]

dfGL['연도'] = 'PY' ##유일하게 달라지는 부분

## b-1. 추가 수기처리

dfGL["회계월"] = dfGL["회계월"].apply(lambda x: x[-2:]) # 회계월 가공
dfGL["차변금액"] = dfGL["차변금액"].fillna(0)
dfGL["대변금액"] = dfGL["대변금액"].fillna(0)
dfGL["전표금액"] = dfGL["차변금액"] - dfGL["대변금액"] # 전표금액 = 차변잔액 - 대변잔액
dfGL["Company code"] = dfGL["계정과목코드"] + "_" + dfGL["계정과목명"] # Company Code

dfGLPY = dfGL

## c. Concatenate
tmp = dfGLCY.shape[0] + dfGLPY.shape[0]
dfGL = pd.concat([dfGLCY, dfGLPY], axis=0)
dfGL.shape[0] == tmp #검증식

###################################

## 4) FSLine 추가

from mylib import myFileDialog as myfd
import pandas as pd

#계정과목 매핑을 읽는다.
dfAcc = pd.read_excel(myfd.askopenfilename())
dfGLShape = dfGL.shape
dfGLJoin = dfGL.merge(dfAcc,how='left',left_on='계정과목코드',right_on='DetailCode')
dfGLJoin[dfGLJoin["FSName"].isna()] #Empty여야 함
dfGL = dfGLJoin

##################################################################

# EXPORT
dfGL.to_csv("./raw/dfGL.tsv", sep="\t", index=None)










