import pandas as pd
import os

import dask.dataframe as dd

from ExcelPar.mylib import myFileDialog as myfd
from ExcelPar.mylib.ErrRetry import ErrRetry

class ImportGL :
    @classmethod
    @ErrRetry
    def ImportGL(cls) -> dd.DataFrame:

        flag = input("Select multi-gl files with dask?(Y) >>")
        if flag == 'Y':
            glPath = myfd.askdirectory("Select GL Folder to read *.parquet") #PHW            
            glTgt = glPath + "/*.parquet"
            gl = dd.read_parquet(glTgt) #return gl
        else:
            glFile = myfd.askopenfilename("Select GL File") #PHW
            #glFile = './imported/dfGL.parquet'

            ext = os.path.splitext(glFile)
            match ext[1]:
                case '.tsv':
                    gl = pd.read_csv(glFile, encoding="utf-8-sig", sep="\t", low_memory=False)
                case '.parquet':
                    gl = pd.read_parquet(glFile)
                case _:
                    print("TSV 또는 Parquet를 선택하세요")
        
        return gl