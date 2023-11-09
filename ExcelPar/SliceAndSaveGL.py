########################################################
# py 파일이 있는 폴더 안의 모든 txt를 읽어서,
# 특정 컬럼을 추출해서 일단 Dataframe으로 읽고, 
# and Save as Parquet
# v0.0.3 DD 231109
# SaveAsGL is outdated.
# v0.0.3 DD 231109 : #Slicing은 선택적으로 구현한다.
########################################################

# 전역부
import os
import glob
import csv

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
      rootPath = ""
      ext = ""
      fileTgt = "" #경로+확장자
      #fileName = ""  #저장할 파일명     
      tgtColumns = [] 
      bSlicing = True #Init = True
      sep = ""
      encoding = ""

def SetGlobal():

       print("변수 지정:")
       gc.rootPath = Win2TPy(os.getcwd()) #원래위치 지정              
       gc.path = myfd.askdirectory("추출할 파일들이 있는 폴더를 지정하세요")       
       #gc.path = os.path.dirname(path)

       gc.ext = input("배치돌릴 파일. ex. *.txt>>")
       gc.fileTgt = gc.path + "/" + gc.ext     
       
       gc.sep = input("Seperator(기본값: 탭)>>") or '\t'

       #gc.tgtColumns = ['전표번호','전기일자','차변(S)/대변(H)','현지통화금액','계정','계정명','고객명','항목텍스트'] #HARDCODING
       if (input("Slicing?") == 'Y'):
              gc.tgtColumns = pd.read_excel(gc.rootPath+"/dtype.xlsx",sheet_name='column', header=None)[0].to_list()
              gc.bSlicing = True
       else:
              gc.tgtColumns = []
              gc.bSlicing = False

def Import() -> dd.DataFrame:      
       #Import with dd
       #df = dd.read_csv(gc.path+"/"+gc.fileTgt, sep='\t', encoding='utf-8',dtype='str')
       gc.encoding = input("인코딩(기본값 : cp949)>>") or 'cp949'

       flag = input("USE DASK? (or pandas.) >>") or 'Y'
       match flag:
              case 'Y': #DASK
                     df = dd.read_csv(gc.fileTgt, sep=gc.sep, encoding=gc.encoding,dtype='str')
                     dfComputed = Compute(df)
                     return dfComputed
              case _: #PANDAS
                     dfCon = ImportFilesByPandas()
                     return dfCon       

def Win2TPy(pathOld:str) -> str:
       return pathOld.replace("\\", "/")      

########################
# READ부

## DD로 일괄 읽는 메서드
def Compute(df : dd.DataFrame) -> pd.DataFrame:
       pbar = ProgressBar()
       pbar.register()

       print("begin computing")

       if gc.bSlicing: #TRUE면
              dfComputed = df[gc.tgtColumns].compute() # 컬럼 슬라이싱
       else:
              dfComputed = df.compute() # Not slicing
       
       print("success. begin exporting")

       return dfComputed

def ImportFilesByPandas(): #bSlicing 사용
       fileList = glob.glob(gc.fileTgt)       
       dfCon = pd.DataFrame()
       for i in fileList:
              print(i)
              dfCon = pd.concat([dfCon,pd.read_csv(i, sep=gc.sep, dtype='str', engine='python', quoting=csv.QUOTE_NONE, encoding=gc.encoding)])
       if gc.bSlicing:
              return dfCon[gc.tgtColumns]
       else:
              return dfCon

def Export(dfComputed : pd.DataFrame):             
       fileName = input("저장할 파일명>>")
       dfComputed.to_parquet(fileName)
       print("Done. Check the exported file:" + fileName)       

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
    
