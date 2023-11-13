# v0.0.2 DD231110 dask로 포팅
import os

import pandas as pd
import dask.dataframe as dd
from dask.diagnostics import ProgressBar

try:
    from ExcelPar.mylib.ErrRetry import ErrRetry
except Exception as e:
    print(e)    
    from mylib.ErrRetry import ErrRetry
try:
    from ExcelPar.mylib import myFileDialog as myfd
except Exception as e:
    print(e)
    from mylib import myFileDialog as myfd    

class TBGLRecon:

    @classmethod
    def TBGLReconWrapper(cls):

        #pathWork = myfd.askdirectory("Select WORK Folder") #NOT USE ACCTMAP
        #os.chdir(pathWork)

        #Select Folder / file
        path = myfd.askdirectory("GL 파일들이 있는 폴더를 선택하세요")
        file = input("파일형식(*.parquet 등)>>")
        tgtPathFile = path + '/' + file
        
        #Load dask
        ext = os.path.splitext(tgtPathFile)[1]
        match ext:
            case '.parquet':
                df = dd.read_parquet(tgtPathFile)
                ProgressBar().register()
            case '.tsv':
                print("tsv")

        cls.TBGLRecon(df) #Inject dd

    @classmethod
    @ErrRetry    
    def TBGLRecon(cls, dfGL:dd.DataFrame): # adapter pattern #DD

        print("TB/GL Recon 기초자료를 추출합니다.")    
        #dfGL.pivot_table(index=["계정과목코드","계정과목명"],columns="연도",values="전표금액",aggfunc='sum').to_excel("검증_GL.xlsx")
        #dfGL.pivot_table(index=["계정과목코드","계정과목명"],columns="연도",values="전표금액",aggfunc='sum').compute().to_excel("검증_GL.xlsx")
        #dfGL = dfGL.categorize(columns=['연도'])
        #dfTmp = dfGL.pivot_table(index="계정과목코드",columns="연도",values="전표금액",aggfunc='sum').compute() # FOR DASK PIVOT
        dfTmp:pd.DataFrame = dfGL.groupby(['계정과목코드','계정과목명','연도'])['전표금액'].sum().compute() # GROUPBY AND UNSTACK
        dfTmp = dfTmp.unstack('연도')
        dfTmp.to_excel("검증_GL.xlsx")        
        print("검증_GL.xlsx 추출완료\n")

        # TB LOAD from imported
        tbFile = myfd.askopenfilename("SELECT dfTB.tsv") #PHW
        #tbFile = './imported/dfTB.tsv'
        pd.read_csv(tbFile, encoding="utf-8-sig", sep="\t").to_excel("검증_TB.xlsx")
        print("검증_TB.xlsx 추출완료\n")

        print("Recon은 별도로 진행하세요.")

    # -> 이후 별도 엑셀에서 TB GL Recon한다.

def RunTBGLRecon():
    TBGLRecon.TBGLReconWrapper()