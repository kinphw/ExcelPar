################################
#GL을 READ하여 금액 전처리후 PARQUET로 저장
# v0.0.1 DD 231104
################################

# 추후 프로그램으로 구현

import os
import glob

import pandas as pd
import dask.dataframe as dd
from dask.diagnostics import ProgressBar
from dask.distributed import Client

from ExcelPar.mylib import myFileDialog as myfd


# Client Setting
cluster = Client(n_workers=4) #두뇌풀가동
pbar = ProgressBar()
pbar.register()

## 필요한 1회용 함수 선언부

#숫자 뒤에 -를 붙여서 음수가 추출된 경우 앞으로 붙여주는 함수
def InvertMinus(tgt:str) -> str:
        if len(tgt) <= 1 : return tgt #1글자거나(-) 1글자보다 적으면('') 그냥 바로 반환
        if tgt[-1] == '-':
            tgt = tgt.replace('-','')
            tgt = "-" + tgt
            return tgt
        return tgt

def ColumnStrToInt(df:pd.DataFrame, ColumnName:str) -> None:    
    try:
        #df[ColumnName] = df[ColumnName].fillna(0).apply(str.strip).apply(lambda x : x.replace(',','') ).apply(lambda x:x.replace('','0')).astype('float64').astype('int64')
        #df[ColumnName] = pd.to_numeric(df[ColumnName].fillna(0).astype('str').apply(str.strip).apply(lambda x : x.replace(',','')).apply(InvertMinus).replace( '[,)]','', regex=True).replace('[(]','-',regex=True))
        tmp = df[ColumnName].fillna(0)
        tmp = tmp.astype('str')
        tmp = tmp.apply(str.strip)
        tmp = tmp.apply(lambda x: x.replace(',',''))

        # 위 아래 중 택일해서 돌려아함
        #tmp = tmp.apply(InvertMinus) 

        tmp = tmp.replace('^-$', '0', regex=True) #그냥 -를 0으로 바꿔준다.
        tmp = tmp.replace( '[,)]','', regex=True).replace('[(]','-',regex=True) #()를 마이너스로 바꿔주는 함수        

        tmp = pd.to_numeric(tmp)
        
        df[ColumnName] = tmp

    except Exception as e:
        print(e)        

# 경로체크
path = myfd.askdirectory()

list = glob.glob(path+"/"+"*.tsv")
file = list[0]
#읽기
dfCon = pd.DataFrame()

for file in list:
    print(file , "작업 개시")
    df = pd.read_csv(file, sep="\t", encoding='cp949', quotechar='"')
    
    #COLUMNS TRIM
    df.columns = df.columns.map(str.strip)

    #전처리 => 핵심로직
    ColumnStrToInt(df,'차변원화')
    ColumnStrToInt(df,'대변원화')

    df['차변원화'].sum()
    df['대변원화'].sum()

    print(file , "전처리 완료")

    #파케이 저장을 위해 object to string

    for column in df.columns:
         #print(column)
         if df[column].dtype == 'object':
              df[column] = df[column].astype(str)

    
    #새로 저장(파케이)
    #filenameNew = file.replace(".DAT", ".tsv")
    filenameNewParquet = file.replace(".tsv", ".parquet")
    df.to_parquet(filenameNewParquet)
    #df.to_csv(filenameNew,sep="\t",index=None, encoding='utf-8')
    
    dfCon = pd.concat([dfCon,df])

    print(filenameNewParquet, "작업 완료.")

    #기존파일 삭제 => 삭제는 어렵지 않으니 따로 삭제
    #os.remove(file)
    #print(file , "삭제 완료")

################################
# 합쳐서 일괄 파케이로 저장
################################

#전처리.. 짤라버린다
df1 = pd.read_parquet(list[0])
df2 = pd.read_parquet(list[1])
df2 = df2.drop(columns=df2.columns[-2:])

df2.to_parquet(list[1])

# 검증부
# LOAD
list = glob.glob(path+"/"+"*.parquet")
df = dd.read_parquet(path+"/"+"*.parquet")

# VALIDATE
df['차변원화'].sum().compute()
df['대변원화'].sum().compute()

df.compute().to_parquet("GLCY.parquet")


