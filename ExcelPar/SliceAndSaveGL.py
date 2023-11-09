########################################################
# py 파일이 있는 폴더 안의 모든 txt를 읽어서,
# 특정 컬럼을 추출해서 일단 Dataframe으로 읽고,
# and Save as Parquet
# v0.0.2 DD 231107
# SaveAsGL is outdated.
########################################################

# 전역부
import os

import pandas as pd
import dask.dataframe as dd
import dask
import time
from dask.diagnostics import ProgressBar
try:
       from ExcelPar.mylib import myFileDialog as myfd
except Exception as e:
       print(e)
       from mylib import myFileDialog as myfd

class gc():
      path = ""
      fileTgt = ""
      fileName = ""     
      tgtColumns = [] #타겟 컬럼명

def SetGlobal():
       
       path = myfd.askopenfilename()
       
       #gc.path = Win2TPy(input("배치돌릴 폴더명>>"))
       #gc.path = Win2TPy(r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_금호석유화학\00 PBC\금호석유화학_2023 원장 가공")
       gc.path = os.path.dirname(path)

       #gc.fileTgt = input("배치돌릴 파일. ex. *.txt>>")
       #gc.fileTgt = '2023 통합 총계정원장_v2.tsv'
       gc.fileTgt = path

       gc.fileName = input("저장할 파일명>>")
       #gc.fileName = 'rawGLPY.parquet'

       gc.tgtColumns = ['전표번호','전기일자','차변(S)/대변(H)','현지통화금액','계정','계정명','고객명','항목텍스트'] #HARDCODING

def Import() -> dd.DataFrame:      
       #Import with dd
       #df = dd.read_csv(gc.path+"/"+gc.fileTgt, sep='\t', encoding='utf-8',dtype='str')
       df = dd.read_csv(gc.fileTgt, sep='\t', encoding='utf-8',dtype='str')
       return df

def Win2TPy(pathOld:str) -> str:
       return pathOld.replace("\\", "/")      

def Export(df : dd.DataFrame):
       pbar = ProgressBar()
       pbar.register()

       print("begin computing")
       dfComputed = df[gc.tgtColumns].compute()
       print("success. begin exporting")
       dfComputed.to_parquet(gc.fileName)
       print("Done. Check the exported file:" + gc.fileName)

def TrimHeader(df : dd.DataFrame) -> dd.DataFrame:
       headerOld = df.columns
       headerNew = list(map(str.strip, headerOld))       
       
       dictHeader = dict(zip(headerOld, headerNew))
       df = df.rename(columns=dictHeader)

       return df


def RunSliceAndSaveGL():
       SetGlobal() #변수설정
       df = Import()
       df = TrimHeader(df) #DEBUG
       Export(df)

if __name__=="__main__":
       RunSliceAndSaveGL()
    
