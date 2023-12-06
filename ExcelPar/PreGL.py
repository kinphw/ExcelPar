##################################################################
# Excel PAR
# 전처리_G/L
# v 0.0.1 DD 231024
# v 0.0.2 DD 231024 GL 엑셀 순환인식부 추가
# v 0.0.3 DD 231024 코드로 구현
# v 0.0.4 DD 231104 USE PARQUET. import and export
# v 0.0.5 DD 231110 CY/PY처리를 input에 따라 받도록 바꾸고, 검증식 생성로직 분리
##################################################################

#전역부
import pandas as pd
import clipboard
import glob
import os
import numpy as np
import openpyxl
import gc

try:
    from ExcelPar.mylib import myFileDialog as myfd
except Exception as e:
    print(e)
    from mylib import myFileDialog as myfd    
try:
    from ExcelPar.mylib.ErrRetry import ErrRetry
except Exception as e:
    print(e)    
    from mylib.ErrRetry import ErrRetry

#################################################################

class Const:
    #FOR 비트연산을 위한 상수
    TO연도CYPY = 0b1 << 0
    TO회계월FR전기일자yyyy_mm_dd = 0b1 << 2 #2023-01-01
    TO회계월FR전기일자yyyy_mm = 0b1 << 3
    TO회계월FR전기일자yyyymm = 0b1 << 4 #yyyymmdd도 사용
    TO회계월FR전기일자yyyy_mm_ddDOT = 0b1 << 6 #2023.01.01    
    TO회계월FR전기일자yyyymmdd = 0b1 << 12 #2023.11.04

    TO대변금액FR대변금액MINUS = 0b1 << 5    
    TO차대금액FR전표금액 = 0b1 << 7
    TO전표금액FR차대금액 = 0b1 << 8

    TO계정과목명FRDetailName = 0b1 << 11

    #전역변수
    strCYPY = ''

#변수생성부

def SetGlobal():
    while True:
        Const.strCYPY = input("Select Year(CY or PY)>>")
        if not(Const.strCYPY in ['CY','PY']):
            print("다시 입력하세요")
            continue
        else:
            break

def MoveFolder()->str:  
    ##################################################################
    #0. 편의를 위한 폴더 이동
    tgtdir = myfd.askdirectory("수행폴더를 지정하세요")
    os.chdir(tgtdir)
    return tgtdir
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

class TempDF:   

    #@ErrRetry
    @classmethod
    def SaveTempDF(cls, df:pd.DataFrame = None) -> pd.DataFrame: #True면 저장, False면 불러오기
        Filename = input("임시저장합니다. 저장할파일명 입력하세요>")
        df.to_pickle(Filename)
        print("임시저장완료.",Filename)                   
        return df

    #@ErrRetry
    @classmethod    
    def LoadTempDF(cls, Flag:bool=True) -> pd.DataFrame: #True면 저장, False면 불러오기                
        Filename = input("임시저장 파일을 로드합니다. 로드할파일명>")
        df = pd.read_pickle(Filename)        
        print("임시파일 로드완료")
        return df

class ImportGL:

    #상수 선언
    FLAG_EXCEL_SINGLE = '0'
    FLAG_EXCEL_MULTI = '1'
    FLAG_TSV = '2'
    FLAG_PARQUET = '3'
    FLAG_PICKLE = '4'

    #@ErrRetry
    @classmethod
    def ImportGLWraper(cls) -> pd.DataFrame:
        message = """
        단일시트 엑셀 : 0
        복수시트 엑셀 : 1
        TSV : 2
        PARUQET : 3
        PICKLE(TMP) : 4
        """
        #1-1. 읽는다.
        print(message)       
        
        Flag = input(">>")
        df = cls.ImportGL(Flag)           
        return df    

    @classmethod
    #@ErrRetry
    def ImportGL(cls, Flag:str)->pd.DataFrame:
        
        filename = myfd.askopenfilename("Select GL File")

        match Flag:
            case cls.FLAG_EXCEL_SINGLE:
                print("Read Excel with single sheet")
                df = pd.read_excel(filename)
            case cls.FLAG_EXCEL_MULTI:
                print("Read Excel with multi sheet")
                print("복수 시트를 읽습니다.)")
                wb = openpyxl.load_workbook(filename, read_only=True) #read_only for speed
                print(filename)
                sheetCount = len(wb.sheetnames)
                print("시트수:",sheetCount)
                df = pd.DataFrame()
                for i in range(sheetCount):
                    dfTmp = pd.read_excel(filename,sheet_name=i) #Should be i
                    df = pd.concat([df, dfTmp])
                    print(i,"번째 시트를 합쳤습니다.")
            case cls.FLAG_TSV:
                print("tsv를 읽습니다.")
                df = pd.read_csv(filename, sep="\t")
            case cls.FLAG_PARQUET:
                print("Read Parquet")
                df = pd.read_parquet(filename)        
            case cls.FLAG_PICKLE:
                print("read TEMP file")        
                df = TempDF.LoadTempDF()

        print("데이터프레임을 반환합니다.")
        return df

@ErrRetry
def AutoMap(df:pd.DataFrame, tgtdir:str)-> pd.DataFrame:
    
    filenameImportMap = "ImportMAP.xlsx"
    #tgtdir = myfd.askdirectory()
    filenameImportMap = glob.glob(tgtdir+"/"+filenameImportMap)
    filenameImportMap = filenameImportMap[0]
    while True:
        try:
            dfMap = pd.read_excel(filenameImportMap, sheet_name="MAP_GL")
            break
        except Exception as e:
            print(e)
            print("ImportMAP.xlsx 파일이 열려 있는 것 같습니다.")
            input(">>")

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

@ErrRetry
def ReadFlag(tgtdir:str)->bin:
    print("ImportMap.xlsx에서 Flag를 읽습니다.")
    filenameImportMap = "ImportMAP.xlsx"
    #tgtdir = myfd.askdirectory()
    filenameImportMap = glob.glob(tgtdir+"/"+filenameImportMap)
    filenameImportMap = filenameImportMap[0]
    dfMap = pd.read_excel(filenameImportMap, sheet_name="MAP_FLAG",header=None)    
    dfMap.columns = ['Flag', 'Check']

    cond = dfMap.loc[:,'Check'] == 'O'
    FlagOld = dfMap[cond].loc[:,'Flag'].to_list()
    print("Flag:",FlagOld)
    FlagNew = 0b0
    for i,j in enumerate(FlagOld):
        FlagNew = FlagNew | getattr(Const,j)

    print("Flag:",bin(FlagNew))
    return FlagNew

##########################################################################
## 중요
##########################################################################

@ErrRetry
def UserDefinedProc(dfGL, year:str = 'CY', Flag:bin = 0b0) -> pd.DataFrame:    
    # 수기처리 => 회사에 따라 적절히 변형하여 적용한다.    
    #dfGL = dfGL

    #공통처리
    dfGL['계정과목코드'].astype(str)
    dfGL['거래처코드'] = dfGL['거래처코드'].fillna("NA")    
    dfGL['전표번호'] == dfGL['전표번호'].astype('str')
    
    #################################
    #금액처리. 일단 String을 가정하여 바로 숫자처리하되, Error 발생시(아마 이미 int) 바로 fiilna만 적용
    try:
        #dfGL['전표금액'] = dfGL['전표금액'].fillna(0).astype('float64').astype('int64')
        dfGL['전표금액'] = pd.to_numeric(dfGL['전표금액'].fillna(0),downcast='unsigned',errors='coerce')
    except Exception as e:
        print(e)
        dfGL['전표금액'] = dfGL['전표금액'].fillna(0)

    try:
        dfGL['차변금액'] = pd.to_numeric(dfGL['차변금액'].fillna(0),downcast='unsigned',errors='coerce')
    except Exception as e:
        print(e)
        dfGL['차변금액'] = dfGL['차변금액'].fillna(0)
    
    try:
        dfGL['대변금액'] = pd.to_numeric(dfGL['대변금액'].fillna(0),downcast='unsigned',errors='coerce')
    except Exception as e:
        print(e)
        dfGL['대변금액'] = dfGL['대변금액'].fillna(0)
    
    #################################
    #dfGL["Company code"] = dfGL["계정과목코드"].apply(str) + "_" + dfGL["계정과목명"].apply(str) # Company Code # Additional에서 옮겨옴
    # TO연도CYPY = 0b1 << 0
    # TO회계월FR전기일자yyyy_mm_dd = 0b1 << 2
    # TO회계월FR전기일자yyyy_mm = 0b1 << 3
    # TO회계월FR전기일자yyyymm = 0b1 << 4
    # TO대변금액FR대변금액MINUS = 0b1 << 5
    # TO차변금액FR전표금액 = 0b1 << 6
    # TO차대금액FR전표금액 = 0b1 << 7
    # TO전표금액FR차대금액 = 0b1 << 8

    if Const.TO연도CYPY & Flag:        
        dfGL['연도'] = str(year)
        print("연도를 ",str(year),"로 조정")

    #Improve : 전처리 단계에서 아예 일자를 yyyy-mm-dd로 조정하고, 회계월까지 생성함
    print(dfGL.head()) #참고. 231109
    dateFormat = input("날짜형식 입력하세요(예시. %Y-%m-%d or %Y%m%d 등..)>>")
    dfGL['전기일자'] = pd.to_datetime(dfGL['전기일자'],format=dateFormat, errors='coerce') #전기일자부터 %Y-%m-%d로 세팅
    ### DEBUG : with 'coerce'
    dateTmp = pd.to_datetime('2022-01-01', format='%Y-%m-%d')
    dfGL['전기일자'] = dfGL['전기일자'].fillna(dateTmp)
    ###
    dfGL['회계월'] = dfGL['전기일자'].apply(lambda x:x.month)
    
    # if 1 == 0:
    #     if Const.TO회계월FR전기일자yyyy_mm_dd & Flag:
    #         dfGL["회계월"] = dfGL["전기일자"].astype('str').apply(lambda x: x[5:7])  #2023-01-01        
    #         print("회계월 from 전기일자 yyyy-mm-dd")
    #     if Const.TO회계월FR전기일자yyyy_mm & Flag:
    #         dfGL["회계월"] = dfGL['전기일자'].apply(lambda x: x[5:]).astype('int') #2023-06
    #         print("회계월 from 전기일자 yyyy-mm")
    #     if Const.TO회계월FR전기일자yyyymm & Flag:
    #         dfGL["회계월"] = dfGL["전기일자"].astype('int64').astype('str').apply(lambda x: x[4:6]).astype('int64') # 회계월 가공 :202306
    #         print("회계월 from 전기일자 yyyymm")        
    #     if Const.TO회계월FR전기일자yyyy_mm_ddDOT & Flag:        
    #         dfGL['전기일자'] = dfGL['전기일자'].apply(lambda x:x.replace(".","-")) #먼저 .을 -로 바꾼다.
    #         dfGL["회계월"] = dfGL["전기일자"].apply(str).apply(lambda x: x[5:7])  #2023-01-01        
    #         print("회계월 from 전기일자 yyyy.mm.dd")
    #     if Const.TO회계월FR전기일자yyyymmdd & Flag:  #20230101       
    #         dfGL['전기일자'] = dfGL['전기일자'].apply(lambda x:x.replace(".","-")) #먼저 .을 -로 바꾼다.
    #         dfGL["회계월"] = dfGL["전기일자"].apply(str).apply(lambda x: x[5:7])  #2023-01-01        
    #         print("회계월 from 전기일자 yyyy.mm.dd")    
    
    #dfGL['전기일자'] = dfGL['전기일자'].astype(int).astype('str')

    if Const.TO대변금액FR대변금액MINUS & Flag:
        dfGL["대변금액"] = dfGL["대변금액"] * -1
        print("대변금액을 (-)로 조정")

    if Const.TO차대금액FR전표금액 & Flag:
        dfGL["차변금액"] = np.where(dfGL["전표금액"] > 0, dfGL["전표금액"], 0)
        dfGL["대변금액"] = np.where(dfGL["전표금액"] < 0, abs(dfGL["전표금액"]), 0)   
        print("전표금액에서 차변금액/대변금액 생성")

    if Const.TO전표금액FR차대금액 & Flag:
        dfGL["전표금액"] = pd.to_numeric(dfGL["차변금액"],errors='coerce') + pd.to_numeric(dfGL["대변금액"], errors='coerce') #DEBUG : 대변이 (-)여야 정상적용됨
        #DEBUG 231101 : 오류처리를 위해 to_numeric 적용
        print("차변/대변금액에서 전표금액 생성")
    
    return dfGL
    #return dfGL #처리 후 반환. Call by Obj. Ref.이므로 반드시 리턴할 필요는 없으나 수기가공 고려하여

##################################################################

@ErrRetry
def ConcatCYPY(dfGLCY:pd.DataFrame, dfGLPY:pd.DataFrame) -> pd.DataFrame:
## c. Concatenate
    dfGL = pd.DataFrame()
    dfGL = pd.concat([dfGLCY, dfGLPY], axis=0)
    print("행수검증>>")
    print(dfGL.shape[0] == dfGLCY.shape[0] + dfGLPY.shape[0])
    return dfGL

###################################
#검증식 : 차대
def dfGLValidate(dfGL:pd.DataFrame):
    print("전체전표의 합계액 (TOBE 0): ", dfGL['전표금액'].sum())

###################################
## 4) FSLine 추가
@ErrRetry
def AddFSLineCode(dfGL:pd.DataFrame, tgtdir:str)->pd.DataFrame:

    #계정과목 매핑을 읽는다.
    filenameAcctMap = "acctMAP.xlsx"
    #tgtdir = myfd.askdirectory()
    filenameAcctMap = glob.glob(tgtdir+"/"+filenameAcctMap)
    filenameAcctMap = filenameAcctMap[0]
    dfAcc = pd.read_excel(filenameAcctMap)

    dfAcc["DetailCode"] = dfAcc["DetailCode"].astype('str')
    #dfGL["계정과목코드"] = dfGL["계정과목코드"].astype(str)
    try:
        #dfGL["계정과목코드"] = dfGL["계정과목코드"].fillna(0).astype('float64').astype('int64').astype('str') #DEBUG 231105 0150
        dfGL["계정과목코드"] = pd.to_numeric(dfGL["계정과목코드"].fillna(0), downcast='unsigned', errors='coerce').astype('float64').astype('int64').astype('str') #231110 2300 modify
    except Exception as e:
        print(e)        
        dfGL["계정과목코드"] = dfGL["계정과목코드"].fillna(0).astype('str')

    dfGLJoin = dfGL.merge(dfAcc,how='left',left_on='계정과목코드',right_on='DetailCode')

    # NA Test
    print("FSLine Code 추가에 따른 GL 계정과목 완전성 체크를 실시합니다.")
    if not dfGLJoin["FSName"].isna().any(): #False여야 함
        print("완전성체크 PASS.")
    else:
        print("FAIL. 추출합니다. (val_GL_coa_완전성체크.csv")
        dfGLJoin[dfGLJoin["FSName"].isna()]['계정과목코드'].drop_duplicates().to_csv("val_GL_coa_완전성체크.csv", index=None)
    
    return dfGLJoin
##################################################################

#필수실행
def AdditionalCleansing(dfGL:pd.DataFrame, Flag:bin) -> pd.DataFrame:    

    print("Case by Case 추가 클렌징 시작.")    
    
    if Const.TO계정과목명FRDetailName & Flag:    
        dfGL['계정과목명'] = dfGL['DetailName'] #FOR 계정과목명이 GL 원파일에 없는 경우. 사후적으로 붙여줌
        print("계정과목명 추가함")
    
    print("Case by Case 추가 클렌징 종료.")    

    dfGL["Company code"] = dfGL["계정과목코드"].apply(str) + "_" + dfGL["계정과목명"].apply(str) # 마지막으로 이동
    print("Company code가 추가되었습니다. 계정과목코드+계정과목명")

    return dfGL

# #검증식 : TB vs GL Recon
# # GL LOAD

# DEPRECATED => 외부로 분리
# @ErrRetry
# def DoTBGLRecon(dfGL:pd.DataFrame):

#     print("TB/GL Recon 기초자료를 추출합니다.")    
#     dfGL.pivot_table(index=["계정과목코드","계정과목명"],columns="연도",values="전표금액",aggfunc='sum').to_excel("검증_GL.xlsx")
#     print("검증_GL.xlsx 추출완료\n")

#     # TB LOAD from imported
#     #tbFile = myfd.askopenfilename("SELECT dfTB") #PHW
#     tbFile = './imported/dfTB.tsv'
#     pd.read_csv(tbFile, encoding="utf-8-sig", sep="\t").to_excel("검증_TB.xlsx")
#     print("검증_TB.xlsx 추출완료\n")

#     print("Recon은 별도로 진행하세요.")

# # -> 이후 별도 엑셀에서 TB GL Recon한다.
##################################################################
# 추출

def Object2String(df:pd.DataFrame) -> None:
    #df를 받아서, object인 columns을 모두 string으로 변경해줌
    for column in df.columns:    
         if df[column].dtype == 'object':
              df[column] = df[column].astype(str) # Call by Object Refenece이므로 Return 불필요


@ErrRetry
def ExportDF(dfGL:pd.DataFrame):

    if not os.path.exists("./imported"):
        os.makedirs("./imported")        
    # EXPORT
    print("dfGL을 생성합니다.")
    if input("Export to Parquet? Press Y(만약 Err시 tsv로) >>") == 'Y':
        Object2String(dfGL)
        fileName = "./imported/dfGL"+Const.strCYPY+".parquet"        
        dfGL.to_parquet(fileName)
    else:       
        fileName =  "./imported/dfGL"+Const.strCYPY+".tsv"
        dfGL.to_csv(fileName, sep="\t", index=None) #Converted to decorator
    
    print("dfGL을 생성 완료합니다.",fileName)
    #dfGLPY['전표금액'].sum()
    #dfGLPY.groupby('계정과목코드').sum()

def ManualPreprocess(dfGL:pd.DataFrame) -> pd.DataFrame:
    df = dfGL

    while True:
        tmp = input("GL 추가적인 가공이 필요하면 디버깅에서 조정하세요. (객체 df) 아니면 0 입력>>")
        if tmp == '0' :
            print("계속 진행합니다.")
            return df

class Run:    
    @classmethod
    def Run(cls):

        print("GL Processing START:")
        tgtdir = MoveFolder()        
        SetGlobal()
        print("START:",Const.strCYPY)
        df = ImportGL.ImportGLWraper()
        dfGL = AutoMap(df, tgtdir)        
        Flag = ReadFlag(tgtdir)
        dfGL = UserDefinedProc(dfGL, Const.strCYPY, Flag)
        #dfGL = ManualPreprocess(dfGL)
        #dfGLCY = dfGL
        print(Const.strCYPY," Done")

        # ## PY : 분리되어 있는 경우에 수행한다.
        # if input("PY 추가진행한다면 Y>") == 'Y':
        #     df = ImportGL.ImportGLWraper()            
        #     dfGL = AutoMap(df, tgtdir)
        #     dfGL = UserDefinedProc(dfGL, 'PY', Flag) #Flag는 CY와 동일하다고 봄
        #     #dfGL = ManualPreprocess(dfGL)
        #     dfGLPY = dfGL
        #     print("PY Done")
        #     dfGL = ConcatCYPY(dfGLCY, dfGLPY)
        # else:
        #     dfGL = dfGLCY
        
        dfGLValidate(dfGL)

        dfGLJoin = AddFSLineCode(dfGL,tgtdir) #문제의 코드. 될지 모르겠음. 터지면 계정과목만 잘라서 붙이거나 수행불가

        # MEMORY
        del dfGL
        # del dfGLCY
        # del dfGLPY
        dfGL = pd.DataFrame()
        # dfGLCY = pd.DataFrame()
        # dfGLPY = pd.DataFrame()
        gc.collect()
        # MEMORY
        #dfGL = dfGLJoin

        dfGLJoin = AdditionalCleansing(dfGLJoin, Flag) #필요한 경우 추가 정의하여 사용

        #최종단계
        #DoTBGLRecon(dfGL)

        ExportDF(dfGLJoin)
        print("GL Processing END..:")

def RunPreGL():
    Run.Run()

if __name__=="__main__":
    Run.Run()

