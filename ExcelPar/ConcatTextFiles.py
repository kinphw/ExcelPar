#전역부
import ExcelPar.mylib.myFileDialog as myfd
import os
import glob
import pandas as pd
import time
import gc
import openpyxl

###########################################
#경로 내 Text파일을 합치는 코드 by Pandas

#경로설정 및 리스트 추출
path = myfd.askdirectory()
list = glob.glob(path + "/" + "*.tsv")

#리스트갯수 : 검증
len(list)

#합산
dfCon = pd.DataFrame()

startTime = time.time()
for i in list:    
    nowTime = time.time()
    print(i," 시작:")
    #df = pd.read_excel(i)
    df = pd.read_csv(i,sep="\t", encoding="cp949")
    df.columns = [i.strip() for i in df.columns] #추가코드. strip (TRIM) 왜곡될지도..
    dfCon = pd.concat([dfCon,df])
    print(i," 종료:")
    print("소요시간 : ",time.time() - nowTime)

print("총 소요시간: ",time.time()-startTime)

dfCon.to_csv("rawGLCY.tsv", sep='\t', index=False)

