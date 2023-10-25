import pandas as pd
from ExcelPar.mylib import myFileDialog as myfd

class ImportGL:
    @classmethod
    def ImportGL(cls):                
        glFile = myfd.askopenfilename() #PHW
        gl = pd.read_csv(glFile, encoding="utf-8-sig", sep="\t", low_memory=False)
        return gl