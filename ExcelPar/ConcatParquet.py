import pandas as pd
import dask.dataframe as dd
from dask.diagnostics import ProgressBar

import ExcelPar.mylib.myFileDialog as myfd
from ExcelPar.concatText import runConcatText

def __ConcatParquet():    
    flag = input("확장자(parquet OR text) (기본값 parquet)>>") or 'parquet'
    match flag:
        case 'parquet':
            path = myfd.askdirectory("selects Parquet 폴더")
            target = path + "/" + "*.parquet"
            print("합칩니다.")
            ProgressBar().register()
            df:pd.DataFrame = dd.read_parquet(target).compute()
            print("추출합니다.")
            fileNameToExport = input("파일명?>>")
            df.to_parquet(path+"/"+fileNameToExport)
            print("완료")
        case 'text':
            runConcatText() #concatText 모듈에서 호출
        case _:
            'FAIL.'

def RunConcatParquet():
    __ConcatParquet()

if __name__=='__main__':
    RunConcatParquet()

