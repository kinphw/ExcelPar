import pandas as pd
import numpy as np

from ExcelPar.mod.SetGlobal import SetGlobal

class PreprocessGL:
    @classmethod
    def PreprocessGL(cls, gl : pd.DataFrame) -> pd.DataFrame:
        gl['전표번호'] = gl['전표번호'].astype(str) #전표번호 2 String
        gl['전기일자'] = pd.to_datetime(gl['전기일자'],format="%Y-%m-%d") #전기일자 2 Datetime
        gl['계정과목코드'] = gl['계정과목코드'].astype(str) #계정과목코드 2 String   
        gl['거래처코드'] = gl['거래처코드'].fillna('NAN') #DEBUG 231026

        if SetGlobal.Level == 'Detail': #Detail일때만 수행
            print("Detail 수준 GL을 전처리 - Company code를 계정과목코드로 지정합니다.")
            gl["Company code"] = gl["계정과목코드"].apply(str) + "_" + gl["계정과목명"].apply(str) # Company Code    

        print("GL 전처리 Done")

        return gl