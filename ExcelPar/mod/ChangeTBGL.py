import pandas as pd
import numpy as np

from ExcelPar.mod.SetGlobal import SetGlobal

class ChangeTBGL:
    @classmethod
    def ChangeTBGL(cls, tb:pd.DataFrame) -> pd.DataFrame: #return tb #-> list[pd.DataFrame, pd.DataFrame]: 
        
        SetGlobal.Level = "FSLine" #Set class var                

        #Call by object reference이므로 df조작(gl)은 반영됨
        print("TB를 FS Line 기준으로 Reshape합니다.")        
        
        #2. TB를 FS Line 기준으로 재합산한다.    
        tbFS = tb.groupby(['FSCode', 'FSName'])[['당기말','전년동기말','전전기말','전기말']].sum()    
        dfTmp = tb[['FSCode','FSName','BSPL']].drop_duplicates() #계정과목 매핑 테이블을 스스로 생성    
        tbFS = tbFS.merge(dfTmp,how='left', on='FSCode') #BSPL까지 붙임            

        tbFS["CY"] = tbFS["당기말"]
        tbFS["PY"] = np.where(tbFS["BSPL"] == "BS", tbFS["전기말"], tbFS["전년동기말"]) #조건식으로 브로드캐스팅
        tbFS["PY1"] = tbFS["전전기말"]

        try:
            tbFS['계정과목코드'] = tbFS['FSCode'].fillna(0).astype(float).astype(int).astype(str)
        except Exception as e:
            print(e)
            tbFS['계정과목코드'] = tbFS['FSCode'].fillna(0).astype(str)
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
