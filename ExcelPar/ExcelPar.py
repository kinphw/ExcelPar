#####################################################################
# Excel PAR - 분석부
# v0.0.7 DD 231026
#####################################################################

import pandas as pd
import dask.dataframe as dd

# Step1 - Detail

### 1. 전역변수 설정
from ExcelPar.mod.SetGlobal import SetGlobal
# ./mod/SetGlobal.py

### 2. Import_GL
from ExcelPar.mod.ImportGL import ImportGL
# ./mod/ImportGL.py

### 3. PreProcess_GL
from ExcelPar.mod.PreprocessGL import PreprocessGL
# ./mod/PreProcessGL.py

### 4. Import TB
from ExcelPar.mod.ImportTB import ImportTB
# ./mod/ImportTB.py

### 5. PreprocessTB
from ExcelPar.mod.PreprocessTB import PreprocessTB
# ./mod/PreprocessTB.py

### 6. 분석적검토 자동문구 생성
from ExcelPar.mod.AnalyzeAccounts import AnalyzeAccounts
# ./mod/AnalyzeAccounts.py

### 7. 월별증감/누적월별증감을 계산하여 파일로 추출하는 부
from ExcelPar.mod.SummarizeMonthlyVarAmount import SummarizeMonthlyVarAmount
# ./mod/SummarizeMonthlyVarAmount.py

### 8. Lead 생성
from ExcelPar.mod.CreateLeadReport import CreateLeadReport
# ./mod/CreateLeadReport.py

### 후처리
from ExcelPar.mod.PostProcess import Postprocess

###2차 분석을 위해 TBGL을 FS Line으로
from ExcelPar.mod.ChangeTBGL import ChangeTBGL
########################################################

# Step2 - FSLine

### 9.
# ./mod/ChangeTBGL.py

#####################################################################

#####################################################################
class ExcelPar:

    @classmethod
    def Handler(cls):

        #HANDLER
        print("###EXCEL PAR BEGIN:")

        print("\n###Phase1 : Detail")
        
        print("\n#1. 기초정보를 설정합니다.")             
        SetGlobal.SetGlobal()

        print("\n#2-1. 전처리된 GL을 Import합니다.")        
        gl:dd.DataFrame|pd.DataFrame = ImportGL.ImportGL() #import하여 gl 객체 선언(dd)

        #print("\n#2-2. GL을 전처리합니다.")        
        #gl = PreprocessGL.PreprocessGL(gl) #gl 객체 가공함

        print("\n#3-1. TB를 Import합니다.")        
        tb = ImportTB.ImportTB() #import하여 tb 객체 선언

        print("\n#3-2. 분석대상 계정과목을 인식합니다.")        
        tb = PreprocessTB.PreprocessTB(tb) #목표 변경. 대부분 PreTB로 내림

        print("\n#4. 계정별 분석을 실시하고 계정별 분석보고서를 생성합니다.")        
        AnalyzeAccounts.AnalyzeAccounts(gl) #gl을 넣어서 STATIC변수(분석문구)를 업데이트한다.    

        print("\n#5. 계정별 월별/누적월별증감액 보고서를 생성합니다.")        
        tb_월별 = SummarizeMonthlyVarAmount.SummarizeMonthlyVarAmount(gl,tb) #gl과 tb를 넣어 STATIC WORK, tb_월별을 반환

        print("\n#6. 총괄 분석보고서를 생성합니다.")            
        LeadFileName = CreateLeadReport.CreateLeadReport(tb, tb_월별)

        print("\n#7. 후처리 후 파일을 정리합니다.")            
        Postprocess.Postprocess(LeadFileName)

        if input("FS Line 분석을 추가 실시합니까? 실시하면 Y>>") == 'Y':

            print("\n\n###Phase2 : FS Line")

            print("\n#0. GL과 TB의 계정과목을 FS Line으로 대체합니다.")                
            tb = ChangeTBGL.ChangeTBGL(tb) #tb는 반환하여 재설정

            print("\n#3-2. TB를 전처리합니다.")
            PreprocessTB.PreprocessTB(tb)    

            print("\n#4. 계정별 분석을 실시하고 계정별 분석보고서를 생성합니다.")
            AnalyzeAccounts.AnalyzeAccounts(gl)

            print("\n#5. 계정별 월별/누적월별증감액 보고서를 생성합니다.")
            tb_월별 = SummarizeMonthlyVarAmount.SummarizeMonthlyVarAmount(gl,tb)    

            print("\n#6. 총괄 분석보고서를 생성합니다.")    
            LeadFileName = CreateLeadReport.CreateLeadReport(tb, tb_월별)

            print("\n#7. 후처리 후 파일을 정리합니다.")    
            Postprocess.Postprocess(LeadFileName)

        print("###EXCEL PAR END:...")        

def RunEP():
    ExcelPar.Handler()

if(__name__=="__main__"):
    pass    