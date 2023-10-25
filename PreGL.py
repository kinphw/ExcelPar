###
# Excel PAR
# 전처리_G/L
# v 0.0.1 DD 231024

## 인터프리터에서 수행을 권장

###
# v 0.0.2 DD 231024 GL 엑셀 순환인식부 추가
##################################################################

#전역부
from mylib import myFileDialog as myfd
import pandas as pd
import clipboard
import glob
##################################################################
#0. 편의를 위한 폴더 이동
tgtdir = myfd.askdirectory()
os.chdir(tgtdir)
##################################################################

# USE IF NEEDED
#PRE-1) GL_raw 컬럼추출
# df = pd.read_excel(myfd.askopenfilename()) #,  sheet_name="GL")
# a = df.columns.to_list()
# tmp = ','.join(a)
# clipboard.copy(tmp)

##################################################################

#2) GL 컬럼매핑 >> 엑셀로 수행

##################################################################

#3) GL DF 컬럼 전처리

##################################################################

## STE1. CY
df = pd.read_csv(myfd.askopenfilename(),sep="\t")#, sheet_name="GL")

# IF EXCEL
df = pd.read_excel(myfd.askopenfilename())

def autoMap(df:pd.DataFrame)-> pd.DataFrame:
    import glob
    filenameImportMap = "ImportMAP.xlsx"
    #tgtdir = myfd.askdirectory()
    filenameImportMap = glob.glob(tgtdir+"/"+filenameImportMap)
    filenameImportMap = filenameImportMap[0]
    dfMap = pd.read_excel(filenameImportMap, sheet_name="MAP_GL")

    ## a. MAP 대상 먼저 전처리
    dfMapMap = dfMap[dfMap['방법'] == 'MAP']
    dfGL = pd.DataFrame()
    for i in range(0, dfMapMap.shape[0]):    
        dfGL[dfMapMap.iloc[i]["tobe"]] = df[dfMapMap.iloc[i]["asis"]]

    ## b. KEYIN 대상 추가 전처리
    dfMapKeyin = dfMap[dfMap['방법'] == 'KEYIN']

    for i in range(0, dfMapKeyin.shape[0]):        
        dfGL[dfMapKeyin.iloc[i]["tobe"]] = dfMapKeyin.iloc[i]["asis"]
    
    return dfGL

dfGL = autoMap(df)

##########################################################################
## 중요
##########################################################################

# b. 수기처리 => 회사에 따라 적절히 변형하여 적용한다.

def preproc(dfGL, year:str = 'CY') -> pd.DataFrame:
    import numpy as np

    dfGL['계정과목코드'].astype(str)

    dfGL['연도'] = str(year)

    #dfGL["회계월"] = dfGL["회계월"].apply(lambda x: x[-2:]) # 회계월 가공
    #dfGL['전기일자'] = dfGL['전기일자'].astype(int).astype(str)
    dfGL["회계월"] = dfGL['회계월'].apply(lambda x: x[5:]).astype(int)
    #dfGL["회계월"] = dfGL["전기일자"].apply(str).apply(lambda x: x[5:7]) # 회계월 가공
    #dfGL['회계월'] = dfGL['회계월'].astype(int)
    dfGL["차변금액"] = np.where(dfGL["전표금액"] > 0, dfGL["전표금액"], 0)
    dfGL["대변금액"] = np.where(dfGL["전표금액"] < 0, abs(dfGL["전표금액"]), 0)    

    return dfGL #처리 후 반환

#dfGL["전표금액"] = dfGL["차변금액"] - dfGL["대변금액"] # 전표금액 = 차변잔액 - 대변잔액
# import numpy as np
# dfGL["연도"] = dfGL["전기일자"].apply(lambda x:x.year)
# dfGL["연도"] = np.where(dfGL["연도"]==2023,"CY","PY")

#계정과목명 붙여야함
##########################################################################
#dfGL.shape[0]

dfGL = preproc(dfGL, 'CY')

dfGLCY = dfGL


##################################################################

## PY : 분리되어 있는 경우에 수행한다.
df = pd.read_csv(myfd.askopenfilename(), sep="\t")#), sheet_name="GL")

# IF EXCEL
df = pd.read_excel(myfd.askopenfilename())

# a. 자동매핑
dfGLPY = autoMap(df)

# b. 추가 수기처리
dfGLPY = preproc(dfGLPY,'PY')

##################################################################

## c. Concatenate
dfGL = pd.DataFrame()
dfGL = pd.concat([dfGLCY, dfGLPY], axis=0)
dfGL.shape[0] == dfGLCY.shape[0] + dfGLPY.shape[0]

###################################
#검증식 : 차대
dfGL['전표금액'].sum()

###################################
# 계정과목명 붙임
###################################
# dfAcct = pd.read_excel(myfd.askopenfilename())#), sheet_name="GL")
# dfAcct = dfAcct[['before','after_1','after_2']]
# dfTmp = dfTmp.rename(columns={'before':'계정과목코드'})

# dfGLNew = dfGL.merge(right=dfTmp,how='left',on="계정과목코드")
# dfGLNew['after_1'].isna().any() ## False여야 함
# dfGL.shape[0] == dfGLNew.shape[0] # True여야 함

# dfGLNew['계정과목코드'] = dfGLNew['after_1']
# dfGLNew['계정과목명'] = dfGLNew['after_2']

# dfGL = dfGLNew
###################################

## 4) FSLine 추가

from mylib import myFileDialog as myfd
import pandas as pd

#계정과목 매핑을 읽는다.
filenameAcctMap = "acctMAP.xlsx"
#tgtdir = myfd.askdirectory()
filenameAcctMap = glob.glob(tgtdir+"/"+filenameAcctMap)
filenameAcctMap = filenameAcctMap[0]
dfAcc = pd.read_excel(filenameAcctMap)

dfAcc["DetailCode"] = dfAcc["DetailCode"].astype(str)
#dfGL["계정과목코드"] = dfGL["계정과목코드"].astype(str)
dfGL["계정과목코드"] = dfGL["계정과목코드"].astype(float).astype(int).astype(str)

dfGLJoin = dfGL.merge(dfAcc,how='left',left_on='계정과목코드',right_on='DetailCode')

# NA Test
dfGLJoin[dfGLJoin["FSName"].isna()] #Empty여야 함
dfGLJoin[dfGLJoin["FSName"].isna()]['계정과목코드'].drop_duplicates().to_csv("val_GL_coa_완전성체크.csv", index=None)
##################################################################

#필수실행
dfGL = dfGLJoin


## b-2 > 작동을 위해 반드시 적용한다.
dfGL["Company code"] = dfGL["계정과목코드"].apply(str) + "_" + dfGL["계정과목명"].apply(str) # Company Code


#검증식 : TB vs GL Recon
# GL LOAD
dfGL.pivot_table(index=["계정과목코드","계정과목명"],columns="연도",values="전표금액",aggfunc='sum').to_excel("검증_GL.xlsx")

# TB LOAD from imported
tbFile = myfd.askopenfilename() #PHW
pd.read_csv(tbFile, encoding="utf-8-sig", sep="\t").to_excel("검증_TB.xlsx")

# -> 이후 별도 엑셀에서 TB GL Recon한다.
##################################################################
# 추출
import os
if not os.path.exists("./imported"):
    os.makedirs("./imported")

# EXPORT
dfGL.to_csv("./imported/dfGL.tsv", sep="\t", index=None)

#dfGLPY['전표금액'].sum()
#dfGLPY.groupby('계정과목코드').sum()





