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

# to save parquet. 필요시 사용
def Object2String(df:pd.DataFrame) -> None:
    #df를 받아서, object인 columns을 모두 string으로 변경해줌
    for column in df.columns:    
         if df[column].dtype == 'object':
              df[column] = df[column].astype(str) # Call by Object Refenece이므로 Return 불필요

# SAP 전표금액 가공함수. 필요시 사용
class ForSAP:
      
       def InvertMinus(self, tgt:str) -> str:
              if len(tgt) <= 1 : return tgt #1글자거나(-) 1글자보다 적으면('') 그냥 바로 반환
              if tgt[-1] == '-':
                     tgt = tgt.replace('-','')
                     tgt = "-" + tgt
                     return tgt
              return tgt

       def ColumnStrToInt(self, df:pd.DataFrame, ColumnName:str) -> None:    

#숫자 뒤에 -를 붙여서 음수가 추출된 경우 앞으로 붙여주는 함수
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
                                   tmp = tmp.apply(self.InvertMinus)
                            case _:
                                   print("선택하지 않았습니다.")                

                     tmp = pd.to_numeric(tmp, errors='coerce') #에러는 0으로..
                     df[ColumnName] = tmp

              except Exception as e:
                     print(e)        

       ## 전표금액 가공부2. SAP은 100을 곱해줘야 한다.
       def Multiple100(self):
              df['현지통화금액'] = df['현지통화금액'] * 100

########################################################################################################################################################################

def RunSliceAndSaveGL():
       SetGlobal() #변수설정
       df = Import()
       df = TrimHeader(df) #DEBUG
       Export(df)

if __name__=="__main__":
       RunSliceAndSaveGL()
    
