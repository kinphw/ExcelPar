import pandas as pd
import numpy as np

class ChangeTBGL:
    @classmethod
    def ChangeTBGL(cls, gl:pd.DataFrame, tb:pd.DataFrame) -> pd.DataFrame: #return tb #-> list[pd.DataFrame, pd.DataFrame]: 
        
        
        #1. GL [Company Code]를 FS Line으로 변경한다.        
        gl['Company code'] = gl['FSCode'].fillna(0).astype(float).astype(int).astype(str) + "_" + gl['FSName'].astype(str) #DEBUG 231025
        gl['계정과목코드'] = gl['FSCode'].fillna(0).astype(float).astype(int).astype(str) #DEBUG 231025
        gl['계정과목명'] = gl['FSName']

        #Call by object reference이므로 df조작(gl)은 반영됨
        print("TB와 GL의 KEY가 FS Line으로 변경되었습니다.")        
        
        #2. TB를 FS Line 기준으로 재합산한다.    
        tbFS = tb.groupby(['FSCode', 'FSName'])[['당기말','전년동기말','전전기말','전기말']].sum()    
        dfTmp = tb[['FSCode','FSName','BSPL']].drop_duplicates() #계정과목 매핑 테이블을 스스로 생성    
        tbFS = tbFS.merge(dfTmp,how='left', on='FSCode') #BSPL까지 붙임            

        tbFS["CY"] = tbFS["당기말"]
        tbFS["PY"] = np.where(tbFS["BSPL"] == "BS", tbFS["전기말"], tbFS["전년동기말"]) #조건식으로 브로드캐스팅
        tbFS["PY1"] = tbFS["전전기말"]

        tbFS['계정과목코드'] = tbFS['FSCode'].fillna(0).astype(float).astype(int).astype(str)
        tbFS['계정과목명'] = tbFS['FSName']
        tbFS["Company code"] = tbFS["계정과목코드"].apply(str) + "_" + tbFS["계정과목명"].apply(str) # Company Code

        tbFS['T1'] = 'T1'
        tbFS['T2'] = 'T2'
        tbFS['T3'] = 'T3'
        tbFS['T4'] = 'T4'
        tbFS['통제활동의존'] = 'CR'
        tbFS['위험수준'] = 'RL'

        tb = tbFS
        return tb       
