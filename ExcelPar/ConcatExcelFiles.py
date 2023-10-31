#경로 내 엑셀파일을 합치는 코드

#전역부
import ExcelPar.mylib.myFileDialog as myfd
import os
import glob
import pandas as pd
import time
import gc
import openpyxl

#경로설정 및 리스트 추출
path = myfd.askdirectory()
list = glob.glob(path + "/" + "*.xls*")

#리스트갯수 : 검증
len(list)

#합산
dfCon = pd.DataFrame()

startTime = time.time()
for i in list:    
    nowTime = time.time()
    print(i," 시작:")
    df = pd.read_excel(i)
    df.columns = [i.strip() for i in df.columns] #추가코드. strip (TRIM)
    dfCon = pd.concat([dfCon,df])
    print(i," 종료:")
    print("소요시간 : ",time.time() - nowTime)

print("총 소요시간: ",time.time()-startTime)

#####################################

# 파일 내 시트순환시 사용

#대상파일 Read
filename = myfd.askopenfilename()

#파일을 읽어서 시트 수를 찾아낸다
wb = openpyxl.load_workbook(filename)
sheetCount = len(wb.sheetnames)

dfCon = pd.DataFrame()

for i in range(sheetCount):
    df = pd.read_excel(filename,sheet_name=i) #Should be i
    dfCon = pd.concat([dfCon, df])

#####################################

#임시저장 : 일단 RAW를 저장해놔야 처리하기가 용이함
dfCon.to_pickle("2023RAW.pickle") #전혀 미가공

#임시파일 LOAD
dfCon = pd.read_pickle(myfd.askopenfilename())

#계정과목 합산 검증 For sense check
acctColumn = 'ACCT_CD'
amountColumn = 'ACCT_AM'
dfTmp = dfCon[[acctColumn,amountColumn]]
dfTmp.loc[:, amountColumn] = pd.to_numeric(dfCon[amountColumn], errors='coerce').fillna(0) #coerce
dfTmp.groupby(acctColumn).sum().to_excel('SenseCheck_GL_TB_RECON.xlsx')

#추출부
filename = "rawGL_CY.tsv" #가공해서 사용
dfCon.to_csv(filename,sep="\t",index=None)

#GARBAGE COLLECT
del dfCon
gc.collect()
dfCon = pd.DataFrame()




