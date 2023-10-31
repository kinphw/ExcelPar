##################################################################
# Excel PAR
# 전처리_G/L
# v 0.0.1 DD 231024
# v 0.0.2 DD 231024 GL 엑셀 순환인식부 추가
# v 0.0.3 DD 231024 코드로 구현
##################################################################

#전역부
from ExcelPar.mylib import myFileDialog as myfd
import pandas as pd
import clipboard
import glob
import os
import numpy as np
import openpyxl

class Const:
    #FOR 비트연산을 위한 상수
    TO연도CYPY = 0b1 << 0
    TO회계월FR전기일자yyyy_mm_dd = 0b1 << 2 #2023-01-01
    TO회계월FR전기일자yyyy_mm = 0b1 << 3
    TO회계월FR전기일자yyyymm = 0b1 << 4 #yyyymmdd도 사용
    TO회계월FR전기일자yyyy_mm_ddDOT = 0b1 << 6 #2023.01.01    

    TO대변금액FR대변금액MINUS = 0b1 << 5    
    TO차대금액FR전표금액 = 0b1 << 7
    TO전표금액FR차대금액 = 0b1 << 8

    TO계정과목명FRDetailName = 0b1 << 11

#변수생성부

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

def TempDF(Flag:bool=True, df:pd.DataFrame = None) -> pd.DataFrame: #True면 저장, False면 불러오기

    #임시저장하는 파일
    #Filename = "temp.pickle"
    print("load한 dataframe의 임시저장 여부를 확인합니다.")

    if Flag:
        if not input("임시저장한다면 Y>>") == "Y":
            print("임시저장하지 않습니다")
            return            
        Filename = input("임시저장합니다. 저장할파일명 입력하세요>")
        df.to_pickle(Filename)
        print("임시저장완료.",Filename)
        
    else:
        if not input("임시로드한다면 Y>>") == "Y":
            print("임시로드하지 않습니다")
            return            
        Filename = input("임시저장 파일을 로드합니다. 로드할파일명>")
        df = pd.read_pickle(Filename)        
        print("임시파일 로드완료")
    return df
        

def ImportGL(bExcel:bool=True, bMultisheet:bool=False)->pd.DataFrame:
    
    filename = myfd.askopenfilename("Select GL")
    
    ## STE1. CY
    #필요한 부분 활성화
    #df = pd.read_csv(myfd.askopenfilename(),sep="\t")#, sheet_name="GL")
    # IF EXCEL    
    #df = pd.read_excel(myfd.askopenfilename(), sheet_name='GL(23)')
    
    #파일을 읽어서 시트 수를 찾아낸다

    if not bExcel: #엑셀이 아닌 경우
        print("tsv를 읽습니다.")
        df = pd.read_csv(filename, sep="\t")
    
    else: #엑셀인 경우

        if bMultisheet: #복수시트인 경우
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
        else:
            print("단일 파일을 추출합니다.")
            df = pd.read_excel(filename)

    TempDF(True, df)

    print("데이터프레임을 반환합니다.")
    return df


def ImportGLWraper() -> pd.DataFrame:
    #1. 읽을것인지, 백업파일을 불러올 것인지        

    #1-1. 읽는다.
    if input("GL을 Import합니까? Y,(or Load Temp)>") == 'Y':        
        
        bExcel = input("엑셀이면 Y>") == 'Y'
        if bExcel:            
            bMultisheet = input("엑셀이 복수 시트면 Y>") == 'Y'            
            df = ImportGL(bExcel, bMultisheet)                
        
        else:   #엑셀이 아닌 TSV
            print("tsv를 상정합니다.")
            df = ImportGL(bExcel)
            pass
    
    #1-2. 백업을 불러온다.
    else:                
        df = TempDF(False) #불러온 결과가 df
        

    return df    

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


def UserDefinedProc(dfGL, year:str = 'CY', Flag:bin = 0b0) -> pd.DataFrame:    
    # 수기처리 => 회사에 따라 적절히 변형하여 적용한다.    
    dfGL = dfGL

    #공통처리
    dfGL['계정과목코드'].astype(str)
    dfGL['거래처코드'] = dfGL['거래처코드'].fillna("NA")
    dfGL['전표금액'] = dfGL['전표금액'].fillna(0)
    dfGL['차변금액'] = dfGL['차변금액'].fillna(0)
    dfGL['대변금액'] = dfGL['대변금액'].fillna(0)
    dfGL["Company code"] = dfGL["계정과목코드"].apply(str) + "_" + dfGL["계정과목명"].apply(str) # Company Code # Additional에서 옮겨옴
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
    
    if Const.TO회계월FR전기일자yyyy_mm_dd & Flag:
        dfGL["회계월"] = dfGL["전기일자"].apply('str').apply(lambda x: x[5:7])  #2023-01-01        
        print("회계월 from 전기일자 yyyy-mm-dd")
    if Const.TO회계월FR전기일자yyyy_mm & Flag:
        dfGL["회계월"] = dfGL['전기일자'].apply(lambda x: x[5:]).astype('int') #2023-06
        print("회계월 from 전기일자 yyyy-mm")
    if Const.TO회계월FR전기일자yyyymm & Flag:
        dfGL["회계월"] = dfGL["전기일자"].astype('int64').astype('str').apply(lambda x: x[4:6]).astype('int64') # 회계월 가공 :202306
        print("회계월 from 전기일자 yyyymm")        
    if Const.TO회계월FR전기일자yyyy_mm_ddDOT & Flag:        
        dfGL['전기일자'] = dfGL['전기일자'].apply(lambda x:x.replace(".","-")) #먼저 .을 -로 바꾼다.
        dfGL["회계월"] = dfGL["전기일자"].apply(str).apply(lambda x: x[5:7])  #2023-01-01        
        print("회계월 from 전기일자 yyyy.mm.dd")

    #dfGL['전기일자'] = dfGL['전기일자'].astype(int).astype('str')

    if Const.TO대변금액FR대변금액MINUS & Flag:
        dfGL["대변금액"] = dfGL["대변금액"] * -1
        print("대변금액을 (-)로 조정")

    if Const.TO차대금액FR전표금액 & Flag:
        dfGL["차변금액"] = np.where(dfGL["전표금액"] > 0, dfGL["전표금액"], 0)
        dfGL["대변금액"] = np.where(dfGL["전표금액"] < 0, abs(dfGL["전표금액"]), 0)   
        print("전표금액에서 차변금액/대변금액 생성")

    if Const.TO전표금액FR차대금액 & Flag:
        dfGL["전표금액"] = dfGL["차변금액"] + dfGL["대변금액"] #DEBUG
        print("차변/대변금액에서 전표금액 생성")
    
    return dfGL
    #return dfGL #처리 후 반환. Call by Obj. Ref.이므로 반드시 리턴할 필요는 없으나 수기가공 고려하여


#dfGL["전표금액"] = dfGL["차변금액"] - dfGL["대변금액"] # 전표금액 = 차변잔액 - 대변잔액
# import numpy as np
# dfGL["연도"] = dfGL["전기일자"].apply(lambda x:x.year)
# dfGL["연도"] = np.where(dfGL["연도"]==2023,"CY","PY")

#계정과목명 붙여야함
##########################################################################
#dfGL.shape[0]

##################################################################

##################################################################

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
def AddFSLineCode(dfGL:pd.DataFrame, tgtdir:str)->pd.DataFrame:

    #계정과목 매핑을 읽는다.
    filenameAcctMap = "acctMAP.xlsx"
    #tgtdir = myfd.askdirectory()
    filenameAcctMap = glob.glob(tgtdir+"/"+filenameAcctMap)
    filenameAcctMap = filenameAcctMap[0]
    dfAcc = pd.read_excel(filenameAcctMap)

    dfAcc["DetailCode"] = dfAcc["DetailCode"].astype('str')
    #dfGL["계정과목코드"] = dfGL["계정과목코드"].astype(str)
    dfGL["계정과목코드"] = dfGL["계정과목코드"].astype('float64').astype('int64').astype('str')

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
    return dfGL

# #검증식 : TB vs GL Recon
# # GL LOAD

def DoTBGLRecon(dfGL:pd.DataFrame):

    print("TB/GL Recon 기초자료를 추출합니다.")    
    dfGL.pivot_table(index=["계정과목코드","계정과목명"],columns="연도",values="전표금액",aggfunc='sum').to_excel("검증_GL.xlsx")
    print("검증_GL.xlsx 추출완료\n")

    # TB LOAD from imported
    tbFile = myfd.askopenfilename("SELECT dfTB") #PHW
    pd.read_csv(tbFile, encoding="utf-8-sig", sep="\t").to_excel("검증_TB.xlsx")
    print("검증_TB.xlsx 추출완료\n")

    print("Recon은 별도로 진행하세요.")

# -> 이후 별도 엑셀에서 TB GL Recon한다.
##################################################################
# 추출

def ExportDF(dfGL:pd.DataFrame):

    if not os.path.exists("./imported"):
        os.makedirs("./imported")        
    # EXPORT    
    print("dfGL을 생성합니다.")    
    while True:
        try:    
            dfGL.to_csv("./imported/dfGL.tsv", sep="\t", index=None)
            break
        except Exception as e:
            print(e)
            print("파일이 열려있는것 같습니다.. 파일을 삭제하시고 엔터를 누르세요.")
            input(">>")
    print("dfGL을 생성 완료합니다.")
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
        print("Import CY")
        df = ImportGLWraper()
        dfGL = AutoMap(df, tgtdir)

        #Flag = 0b0 #상수. 필요한 경우 조정하여 사용
        Flag = ReadFlag(tgtdir)
        dfGL = UserDefinedProc(dfGL, 'CY', Flag)
        #dfGL = ManualPreprocess(dfGL)

        dfGLCY = dfGL
        print("CY Done")

        ## PY : 분리되어 있는 경우에 수행한다.
        if input("PY 추가진행한다면 Y>") == 'Y':
            #df = ImportGL(True)
            df = ImportGLWraper() #DEBUG : 231101
            dfGL = AutoMap(df, tgtdir)
            dfGL = UserDefinedProc(dfGL, 'PY', Flag) #Flag는 CY와 동일하다고 봄
            #dfGL = ManualPreprocess(dfGL)
            dfGLPY = dfGL
            print("PY Done")
            dfGL = ConcatCYPY(dfGLCY, dfGLPY)
        else:
            dfGL = dfGLCY
        
        dfGLValidate(dfGL)

        dfGLJoin = AddFSLineCode(dfGL,tgtdir)
        dfGL = dfGLJoin

        dfGL = AdditionalCleansing(dfGL, Flag) #필요한 경우 추가 정의하여 사용

        #최종단계
        DoTBGLRecon(dfGL)

        ExportDF(dfGL)
        print("GL Processing END..:")

def RunPreGL():
    Run.Run()

if __name__=="__main__":
    Run.Run()

# #FOR 비트연산
# TO연도CYPY = 0b1 << 0
# TO회계월FR전기일자yyyy_mm_dd = 0b1 << 2
# TO회계월FR전기일자yyyy_mm = 0b1 << 3
# TO회계월FR전기일자yyyymm = 0b1 << 4
# TO대변금액FR대변금액MINUS = 0b1 << 5
# TO차변금액FR전표금액 = 0b1 << 6
# TO차대금액FR전표금액 = 0b1 << 7
# TO전표금액FR차대금액 = 0b1 << 8