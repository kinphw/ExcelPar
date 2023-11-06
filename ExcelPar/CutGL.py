###
# GL을 읽어서 일부만 추출하는 Code FOR SAP(ex. 현대삼호중공업)
# v0.0.1 DD231106
###

from ExcelPar.mylib import myFileDialog as myfd
import pandas as pd
import time

#방법3 : dask 활용
import dask.dataframe as dd
from dask.diagnostics import ProgressBar
from dask.distributed import Client

#1. 파일 지정

path = myfd.askopenfilename()

#1안. read by pandas directly
#2안. read by pandas with chunk
#3안. read by dask => 3안

# #방법1 : 직접 읽는다

# startTime = time.time()
# df = pd.read_csv(path, sep="|", encoding='cp949')
# totalTime = time.time() - startTime #103초

# #방법2 : Chunksize
# chunksize = 10 ** 6 #100만
# reader = pd.read_csv(path, sep="|", encoding='cp949', chunksize=chunksize)
# startTime = time.time()
# for cnt, chunk in enumerate(reader):
#     df = pd.concat([df, chunk])
# totalTime = time.time() - startTime #93초
# print(totalTime)

# DASK Worker 늘리지 않는다.
# cluster = Client(n_workers=4) #두뇌풀가동
# pbar = ProgressBar()
# pbar.register()

# 애초에 필요한 Column만 추출하는 게 효율적이다.

reader = dd.read_csv(path, sep="|", encoding='cp949', dtype={'  BusinessArea ':'object'})

# 컬럼변경. Trim해준다.
oldColumn = reader.columns
newColumn = reader.columns.map(str.strip)
di = dict(zip(oldColumn, newColumn))
reader = reader.rename(columns=di)

# csv로부터 dask.df로 추출한다. # 50초 소요

startTime = time.time()
df = reader[['전표번호','전기일','현지통화금액','계정코드','고객','전표아이템텍스트']]
dfComputed = df.compute()
totalTime = time.time() - startTime #103초
print(totalTime)

df = dfComputed
df.shape

# SAP 전표금액 가공부

def ColumnStrToInt(df:pd.DataFrame, ColumnName:str) -> None:    

    #숫자 뒤에 -를 붙여서 음수가 추출된 경우 앞으로 붙여주는 함수
    def InvertMinus(tgt:str) -> str:
        if len(tgt) <= 1 : return tgt #1글자거나(-) 1글자보다 적으면('') 그냥 바로 반환
        if tgt[-1] == '-':
            tgt = tgt.replace('-','')
            tgt = "-" + tgt
            return tgt
        return tgt

    try:
        #df[ColumnName] = df[ColumnName].fillna(0).apply(str.strip).apply(lambda x : x.replace(',','') ).apply(lambda x:x.replace('','0')).astype('float64').astype('int64')
        #df[ColumnName] = pd.to_numeric(df[ColumnName].fillna(0).astype('str').apply(str.strip).apply(lambda x : x.replace(',','')).apply(InvertMinus).replace( '[,)]','', regex=True).replace('[(]','-',regex=True))
        tmp = df[ColumnName].fillna(0) #일단 공백을 죽이고
        tmp = tmp.astype('str') #문자로 만든 다음
        tmp = tmp.apply(str.strip) #trim처리하고
        tmp = tmp.apply(lambda x: x.replace(',','')) #쉼표는 없앤다.
        tmp = tmp.replace('^-$', '0', regex=True) #그냥 오직 -는 0으로 바꿔준다.

        # 이후 로직은 위 아래 중 택일해서 돌려아함
        #tmp = tmp.apply(InvertMinus) 
        Flag = input("음수표시가 ()면 1, 마지막글자-면 2, 이외엔 그냥 엔터>>")
        match Flag:
            case '1':
                tmp = tmp.replace( '[,)]','', regex=True).replace('[(]','-',regex=True) #()를 마이너스로 바꿔주는 함수        
            case '2':
                tmp = tmp.apply(InvertMinus)
            case _:
                print("선택하지 않았습니다.")                

        tmp = pd.to_numeric(tmp)
        
        df[ColumnName] = tmp

    except Exception as e:
        print(e)        

ColumnStrToInt(df, '현지통화금액')

## 전표금액 가공부2. SAP은 100을 곱해줘야 한다.

df['현지통화금액'] = df['현지통화금액'] * 100

## RECON 추출부

res = df.groupby('계정코드')['현지통화금액'].sum()
res.to_excel("검증_GLCY.xlsx")

df['현지통화금액'].sum() # 차대 Test

#PARQUET로 저장

# 전처리
def Object2String(df:pd.DataFrame) -> None:
    #df를 받아서, object인 columns을 모두 string으로 변경해줌
    for column in df.columns:    
         if df[column].dtype == 'object':
              df[column] = df[column].astype(str) # Call by Object Refenece이므로 Return 불필요

Object2String(df)

# 저장
filename = "rawGLCY.parquet"
df.to_parquet(filename) #5Gb to 300Mb