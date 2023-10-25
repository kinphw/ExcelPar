###
# 에프알엘코리아
###

#테이블 설계는 LEAD Excel파일에서

###
# 3. CSV 2 DF 2 SQL
###

# 전역선언부
import pandas as pd
import tqdm
import sys #매개변수 읽기위해
from mylib import myFileDialog as myfd #주. 라이브러리 소재 경로가 아니라 호출하는 코드 경로 기준임.

## PART1. txt 2 DF

def txt2df(fileName, delimiter : chr = ",", headerBool : bool = True, chunksize = 1000000) -> pd.DataFrame :

    # 1) Dataframe으로 Import
    #path = "01re.txt"
    path = fileName
    df = pd.DataFrame()    
    totalChunk = 0
    #chunksize = 1000000
    #for chunk in pd.read_csv(path, delimiter="|", low_memory=False, chunksize=chunksize):   #이번엔 헤더가 있음 
    if headerBool: #헤더가 있다고 받았으면
        headerArg = 0
    else:
        headerArg = None

    for chunk in pd.read_csv(path, delimiter=delimiter, low_memory=False, header = headerArg, chunksize=chunksize):   #delimiter
        df = pd.concat([df,chunk])
        totalChunk = totalChunk + chunk.shape[0]
        print(totalChunk, "행 Done")
    #이 시점에 dfF 전처리
    df = preproc(df)

    return df

def preproc(df : pd.DataFrame) -> pd.DataFrame:
    #전처리하는 함수 : for FRL
    tmp = df.columns.to_list()
    tmplist = []
    for i in tmp:
        #print(i)
        try:
            tmplist.append(i.strip())
        except:
            tmplist.append(i)
    df.columns = tmplist

    # 숫자로 만들기 - 거래통화
    df['Amount'] = df['Amount'].str.replace(",","")
    df['Amount'] = pd.to_numeric(df['Amount'])

    # 숫자로 만들기 - 표시통화
    df['Loc.curr.amount'] = df['Loc.curr.amount'].str.replace(",","")
    df['Loc.curr.amount'] = pd.to_numeric(df['Loc.curr.amount'])

    return df

## PART2. DF 2 DB

def df2db(dfF : pd.DataFrame, dbName : str, tableName : str = "", argv1=1000000):
    
    # 2) DF 2 MySQL - 전처리    

    engine = dbCon('mysql', dbName) #설정된 engine객체 반환

    # 7) DF 2 MySQL - 작업

    #tableName = input("입력할 테이블명(DBMS)>>")
    #tableName = tableName
    #tableName = 'hs01'

    tgtNo = dfF.shape[0] #행수 기록d

    #pbar
    pbar = tqdm.tqdm(total=tgtNo , desc="작업대상행수")

    startNo = 0
    endFlag = True
    resultInsertAcc = 0
    #chunksize = 50000
    chunksize = argv1 #받은 인자 연결 230921

    while endFlag:        
        endNo = startNo + chunksize    
        if (endNo > tgtNo):
            endNo = tgtNo
            endFlag = False #종료시키기
        df1 = dfF.iloc[startNo:endNo,:]   #df1 : 입력할 chunk    
        tmp = pbar.update(endNo-startNo)
        resultInsert = df1.to_sql(name=tableName, con=engine, index=False, if_exists="append") #index 제거, append로 해야 계속 넣을 수 있음
        #print(startNo,"/",endNo,"/ Insert 완료")
        startNo = startNo + chunksize
        resultInsertAcc = resultInsertAcc + resultInsert
    pbar.close()    

def dbCon(flag, dbName : str = 'frl'):

    import pymysql
    import sqlalchemy
    pymysql.install_as_MySQLdb()
    import MySQLdb    
    import pyodbc

    if flag=="mysql":
        mySQL_ID = "root"
        mySQL_PW = "genius"
        #mySQL_DB = "frl" #호텔신라
        mySQL_DB = dbName #호텔신라
        #MySQL에 연결하는 경우
        engine = sqlalchemy.create_engine("mysql+mysqldb://"+mySQL_ID+":"+mySQL_PW+"@127.0.0.1/"+mySQL_DB, encoding='utf-8')

    elif flag=="mssql":

        #MSSQL에 연결하는 경우
        engine = sqlalchemy.create_engine("mssql+pyodbc://SA:qkrguddnjs1!@mymssql")
        #engine.connect()
    
    engine.connect()
    print("DB에 연결되었습니다.")
    return engine #engine을 반환

#230921
def getArgv1():
    if len(sys.argv) < 2:
        print("DF2SQL ChunkSize를 설정하지 않았으므로 기본값인 100,000으로 설정합니다.")
        sys.argv.append(100000)
    else:
        print("DF2SQL ChunkSize : ", sys.argv[1])

if(__name__=="__main__"):

    getArgv1() #매개변수 없으면 추가

    print("Start:")
    #fileName = input("입력할 txt명>>")
    fileName = myfd.askopenfilename()
    chunksize = int(sys.argv[1])
    df = txt2df(fileName, "\t", True, chunksize)
    df2db(df, 'frl', chunksize) #두번째 인자 : 연결할 데이터베이스
    print("End:")

