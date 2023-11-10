import pandas as pd
import numpy as np

from ExcelPar.mod.SetGlobal import SetGlobal

class PreprocessTB:
    @classmethod
    def PreprocessTB(cls, tb:pd.DataFrame) -> pd.DataFrame:
        
        tb['계정과목코드'] = tb['계정과목코드'].astype(str)

        #증감금액 = CY와 PY의 차이임. PY 전처리는 별도 코드에서 수행
        tb["증감금액"] = tb["CY"] - tb["PY"]
        
        #증감비율
        # CY,PY 모두 0이면 0
        # CY가 0이 아니고 PY는 0이면 1
        # 이외에는 변동 / PY

        tb["증감비율"] = np.where((tb["CY"] == 0) & (tb["PY"] == 0), 0,
                            np.where((tb["CY"] != 0) & (tb["PY"] == 0), 1, tb["증감금액"]/tb["PY"]))
        ###########################################################################################

        # 통제위험 & 위험평가 결과에 따라 Threshold 달리 동적으로 적용하는 코드, 현재는 미사용

        # conditions = [
        #     (tb["통제활동의존"] == "NotRelyingOnControls") & (tb["위험수준"] == "LowerRisk"),
        #     (tb["통제활동의존"] == "NotRelyingOnControls") & (tb["위험수준"] == "HigherRisk"),
        #     (tb["통제활동의존"] == "NotRelyingOnControls") & (tb["위험수준"] == "SignificantRisk"),
        #     (tb["통제활동의존"] == "RelyingOnControls") & (tb["위험수준"] == "HigherRisk"),
        #     (tb["통제활동의존"] == "RelyingOnControls") & (tb["위험수준"] == "SignificantRisk")
        # ]

        # choices = [
        #     np.minimum(tb["CY"].abs() * 0.22, PM * 0.65),
        #     np.minimum(tb["CY"].abs() * 0.15, PM * 0.45),
        #     np.minimum(tb["CY"].abs() * 0.35, PM * 0.95),
        #     np.minimum(tb["CY"].abs() * 0.25, PM * 0.9),
        #     np.minimum(tb["CY"].abs() * 0.2, PM * 0.5)
        # ]

        # tb["Threshold"] = np.select(conditions, choices, default=0)
        ###########################################################################################

        #Threshold = Min[당기잔액*20%, PM*50%]
        tb["Threshold"] = np.minimum(
            tb['CY'].abs() * 0.2,
            SetGlobal.PM * 0.5)

        # 변동액 <= CTT : X
        # 변동액 >= Threshold 또는 변동비율 > 20% : O

        conditions = [
            tb["증감금액"].abs() <= SetGlobal.De_minimis,
            (tb["증감금액"].abs() >= tb["Threshold"]) | (tb["증감비율"].abs() >= SetGlobal.diff_비율) #오타수정
        ]

        choices = ["X", "O"]

        tb['분석대상'] = np.select(conditions, choices, default="X")

        tb_분석대상 = tb[tb["분석대상"] == "O"]
        
        if SetGlobal.Level == 'Detail':
            SetGlobal.분석계정과목 = tb_분석대상["Company code"].unique()    #Detail이면 기존 Logic     
        elif SetGlobal.Level == 'FSLine':
            tmp1 = pd.to_numeric(tb_분석대상["FSCode"].fillna(0),errors='coerce',downcast='unsigned').astype('str')
            tmp2 = tb_분석대상["FSName"].fillna(0).astype('str')
            tb_분석대상["New"] = tmp1+"_"+tmp2
            SetGlobal.분석계정과목 = tb_분석대상["New"].unique()        #FSLine이면 그냥 계정과목코드로
        else:
            pass

        ## 대분류 증감요인

        tb["T1_증감금액"] = tb.groupby(["T1"])["증감금액"].transform("sum")
        tb["T2_증감금액"] = tb.groupby(["T2"])["증감금액"].transform("sum")
        tb["T1_설명비율"] = np.where((tb["T2_증감금액"] == 0) & (tb["T1_증감금액"] == 0), 0,
                            np.where((tb["T2_증감금액"] != 0) & (tb["T1_증감금액"] == 0), 1, tb["T2_증감금액"]/tb["T1_증감금액"]))

        tb["T2_설명비율"] = np.where((tb["증감금액"] == 0) & (tb["T2_증감금액"] == 0), 0,
                            np.where((tb["증감금액"] != 0) & (tb["T2_증감금액"] == 0), 1, tb["증감금액"]/tb["T2_증감금액"]))
        
        return tb
