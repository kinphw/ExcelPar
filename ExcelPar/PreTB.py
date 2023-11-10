###
# Excel PAR
# 전처리_T/B
# v 0.0.2 DD 231110
#
# v 0.0.2 : 추상화
###

# 검증1순위 : 차대일치
# 검증2순위 : TB vs GL recon
##################################################################

#전역부
import clipboard
import os
import numpy as np
import glob
from abc import ABCMeta, abstractmethod

import pandas as pd

from ExcelPar.mylib import myFileDialog as myfd
from ExcelPar.mylib.ErrRetry import ErrRetry

class ABCPreTB(metaclass=ABCMeta):
    #ABSTRACT
    @classmethod
    @abstractmethod    
    def Handler(cls): pass #Process Handler

    @classmethod
    @abstractmethod    
    def MoveFolder(cls): pass #작업폴더로 이동

    @classmethod
    @abstractmethod    
    def ImportTB(cls): pass #TB를 읽는다.

    @classmethod
    @abstractmethod    
    def autoMap(cls): pass #Excel로 읽은 사전정보로 TB를 자동 전처리한다.

    @classmethod
    @abstractmethod    
    def AddFSLineCode(cls) : pass #FSLine Code를 추가한다.

    @classmethod
    @abstractmethod    
    def AdditionalCleansing(cls): pass #필요한 추가 클렌징 실시

    # @classmethod
    # @abstractmethod    
    # def AnalyzeTB(cls): pass #분석목적 생성코드 PreprocessTB에서 가져온다.

    @classmethod
    @abstractmethod    
    def ExportDF(cls) : pass #데이터프레임을 Export한다.

class PreTB(ABCPreTB):
    
    @classmethod
    def Handler(cls):
        print("TB Processing START:")        
        tgtdir = cls.MoveFolder()
        df = cls.ImportTB()
        dfTB = cls.autoMap(df, tgtdir)
        dfTB = cls.AddFSLineCode(dfTB,tgtdir)
        cls.AdditionalCleansing(dfTB) #Call by Obj. Refe이므로 return 불필요)     
        #dfTB = cls.AnalyzeTB(dfTB) #중요성 고려한 분석대상 선정 => 기존 Preprocess 코드 집합
        cls.ExportDF(dfTB)
        print("TB Processing END..:")

    ##################################################################
    #0. 편의를 위한 폴더 이동
    @classmethod
    def MoveFolder(cls)->str:
        tgtdir = myfd.askdirectory("기본 WORK폴더 지정(AcctMAP, ImportMAP 있는)")
        os.chdir(tgtdir)
        print("폴더를 이동했습니다.")
        return tgtdir

    #1. TB 가공 >> Excel

    ##################################################################

    #2. TB Import 
    @classmethod
    @ErrRetry
    def ImportTB(cls) -> pd.DataFrame:
        while True:
            try:
                tmp = input("Excel : 1, csv : 2>>")
                if int(tmp) == 1:
                    df = pd.read_excel(myfd.askopenfilename("TB파일 선택"))#, sheet_name="IMPORT")    
                elif int(tmp) == 2:
                    df = pd.read_csv(myfd.askopenfilename("TB파일을 선택"))
                break
            except Exception as e:
                print(e)
                print("오류. 다시 선택하세요.")

        df = df.fillna(0) #전처리 - NaN to 0
        print("TB Imported")
        return df

    # 이 단계에서 df는 rawdata

    ##################################################################

    #2) TB 컬럼매핑 >> 엑셀로 수행

    #USE IF NEEDED

    #헤더 추출 > 필요하면 사용
    @classmethod
    def ExportHeader(cls, df:pd.DataFrame):
        tmp = df.columns.to_list()
        tmp = map(str,tmp)
        tmp = list(tmp)
        tmp = ','.join(tmp)
        clipboard.copy(tmp)

    ##################################################################


    #3) DF 컬럼 전처리


    #폴더를 지정한 후, 지정 폴더에서 ImportMAP.xlsx를 찾습니다.
    @classmethod
    @ErrRetry
    def autoMap(cls, df:pd.DataFrame, tgtdir:str)->pd.DataFrame :

        print("Auto Mapping. Read ImportMAP.xlsx")
        filenameImportMap = "ImportMAP.xlsx"
        filenameImportMap = glob.glob(tgtdir+"/"+filenameImportMap)
        filenameImportMap = filenameImportMap[0]
        dfMap = pd.read_excel(filenameImportMap, sheet_name="MAP_TB")

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


        print("AUTO-MAP Done")
        return dfTB

    ## 4) FSLine 추가

    #계정과목 매핑을 읽는다.
    @classmethod
    @ErrRetry
    def AddFSLineCode(cls,dfTB:pd.DataFrame, tgtdir:str):
        filenameAcctMap = "acctMAP.xlsx"
        #tgtdir = myfd.askdirectory()
        filenameAcctMap = glob.glob(tgtdir+"/"+filenameAcctMap)
        filenameAcctMap = filenameAcctMap[0]
        dfAcc = pd.read_excel(filenameAcctMap)

        dfAcc["DetailCode"] = dfAcc["DetailCode"].astype('str')
        dfTB["계정과목코드"] = dfTB["계정과목코드"].astype('str')

        dfTBJoin = dfTB.merge(dfAcc,how='left',left_on='계정과목코드',right_on='DetailCode')

        dfTBJoin.shape[0] == dfTB.shape[0]

        #정합성 체크문
        #Empty여야 함 =>
        if dfTBJoin[dfTBJoin["FSName"].isna()].shape[0]  == 0:
            print("매핑결과 이상없습니다.")    
        else:
            print("매핑결과 매핑되지 않은 계정이 있습니다. error.xlsx를 추출합니다.")
            dfTBJoin[dfTBJoin["FSName"].isna()].to_excel("error.xlsx")

        #최종
        dfTB = dfTBJoin

        print("FSLINE Code Added")

        return dfTB

    #########################################
    ## CY/PY 설정_분반기 적용 코드 - 필수

    @classmethod
    def AdditionalCleansing(cls,dfTB:pd.DataFrame):

        #return 불필요
        dfTB["CY"] = dfTB["당기말"]
        dfTB["PY"] = np.where(dfTB["BSPL"] == "BS", dfTB["전기말"], dfTB["전년동기말"]) #조건식으로 브로드캐스팅 ## CY/PY 설정_분반기 적용 코드 - 필수
        dfTB["PY1"] = dfTB["전전기말"]
        dfTB["Company code"] = dfTB["계정과목코드"].apply(str) + "_" + dfTB["계정과목명"].apply(str) # Company Code

        print("CY/PY/PY1/Company Code set")
    ##################################################################

        # VALIDATION  # TRUE가 아니면 문구 진행불가
        if not dfTB['계정과목코드'].shape[0] == dfTB['계정과목코드'].drop_duplicates().shape[0]:
            print("계정코드 중복이 있음. 진행불가. 확인 필요")
        else:
            print("계정코드 중복없음. PASS")

    @classmethod
    def ExportDF(cls,dfTB:pd.DataFrame):
        ##################################################################
        if not os.path.exists("./imported"):
            print("폴더를 생성한다.")
            os.makedirs("./imported")

        # EXPORT
        dfTB.to_csv("./imported/dfTB.tsv", sep="\t", index=None)
        print("dfTB 생성완료")

#ENTRY POINT
def RunPreTB():
    PreTB.Handler()

if __name__=="__main__":
    PreTB.Handler()
