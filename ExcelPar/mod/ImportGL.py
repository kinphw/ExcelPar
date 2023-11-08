import pandas as pd
import os

from ExcelPar.mylib import myFileDialog as myfd
from ExcelPar.mylib.ErrRetry import ErrRetry

class ImportGL:
    @classmethod
    @ErrRetry
    def ImportGL(cls):
        glFile = myfd.askopenfilename("Select GL") #PHW
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