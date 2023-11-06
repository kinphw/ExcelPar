# py 파일이 있는 폴더 안의 모든 txt를 읽어서 6개 컬럼을 추출해서 일단 Dataframe으로 읽는 코드 FOR KIA
# and Save as Parquet
# v0.0.1 DD 231106

# 전역부
import pandas as pd
import dask.dataframe as dd
import time

fileName = "202207.parquet"

# 읽고

df = dd.read_csv("*.txt", sep='|', encoding='cp949', header=None,
                 dtype={10: 'object',
       17: 'object',
       34: 'object',
       35: 'object'})

# 헤더설정하고

dfHeader = pd.read_csv("BKPF_BSEG_202207_Header.tsv", encoding='cp949', sep='|', header = None)
headerTmp = list(map(str, dfHeader.loc[0,:].to_list()))
headerNew = list(map(str.strip, headerTmp))

headerOld = df.columns
dictHeader = dict(zip(headerOld, headerNew))

df = df.rename(columns=dictHeader)

# Slicing해서 DF로 읽고

tgtColumns = ['전표헤더','전기일','차대변지시자','현지통화금액','계정코드','품목텍스트']

startTime = time.time()
dfComputed = df[tgtColumns].compute()
time.time() - startTime

# 저장
dfComputed.to_parquet(fileName)