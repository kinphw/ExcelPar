######################################################################
# BKPF BSEG join for 신영 => raw를 만드는 과정. 이후에는 공통 전처리 메서드 적용
######################################################################

# FOR 인터프리터
# v0.0.1 DD 231024

from mylib import myFileDialog as myfd
import pandas as pd
import openpyxl
import clipboard
import os

##################################################################
#0. 편의를 위한 폴더 이동
tgtdir = myfd.askdirectory()
os.chdir(tgtdir)

##################################################################
# 1. read BKPF

# 엑셀순환인식 후 합친다.

#대상파일 Read
filename = myfd.askopenfilename()

#파일을 읽어서 시트 수를 찾아낸다
wb = openpyxl.load_workbook(filename)
sheetCount = len(wb.sheetnames)

dfBKPF = pd.DataFrame()

for i in range(sheetCount):
    df = pd.read_excel(filename,sheet_name=i) #Should be i
    dfBKPF = pd.concat([dfBKPF, df])

#중복제거
dfBKPF.drop_duplicates().shape[0]
dfBKPF = dfBKPF.drop_duplicates()

# 2. read BSEG

# BSEG
filename = myfd.askopenfilename()
dfBSEG = pd.read_csv(filename, sep="|", encoding="cp949")

# (선택) 컬럼추출
clipboard.copy(','.join(dfBKPF.columns.to_list()))
clipboard.copy(','.join(dfBSEG.columns.to_list()))

################################
#JOIN
dfJoin = dfBSEG.merge(right = dfBKPF[['BELNR','BLART','BUDAT','BKTXT']],how='left',on='BELNR')
dfJoin.shape[0] == dfBSEG.shape[0] #True
dfGL = dfJoin

#EXPORT - 단순 JOIN
dfJoin.to_csv("./raw/rawGL.txt", sep="\t", index=None)
################################
#IMPORT 
import os
tmp = myfd.askdirectory()
os.chdir(tmp)

import pandas as pd
dfGL=pd.read_csv("./raw/rawGL.txt",sep="\t")



################################
# 계정과목 추가 - 해당 폴더
dfAcct=pd.read_excel(myfd.askopenfilename())
dfAcct=dfAcct[["DetailCode","DetailName"]]
dfAcct=dfAcct.rename(columns={'DetailCode':'HKONT',
                'DetailName':'계정과목명'                
            })
dfGLMerged = dfGL.merge(right=dfAcct,how='left',on='HKONT')
dfGL.shape[0] == dfGLMerged.shape[0] #True


del dfGL
dfGL = dfGLMerged

# 차대에 따라 방향을 바꿔준다
import numpy as np

dfGL['DMBTR'] = np.where(dfGL['SHKZG'] == 'S', dfGL['DMBTR'], dfGL['DMBTR']*-1)
dfGL['WRBTR'] = np.where(dfGL['SHKZG'] == 'S', dfGL['WRBTR'], dfGL['WRBTR']*-1)

# 기간 외 전표 삭제

startDate = 20220401
endDate = 20230331

dfGL[~(dfGL['BUDAT'] < startDate)]
dfGL[dfGL['BUDAT'] < startDate]

dfGL[~dfGL['BUDAT'] > endDate]
dfGL[dfGL['BUDAT'] > endDate]
dfGLtmp = dfGL[(dfGL['BUDAT'] >= startDate) & (dfGL['BUDAT'] <= endDate)]

dfGLtmp.shape[0] == dfGL.shape[0]

dfGL = dfGLtmp

# * 100 # 소수점이 있으니 astype(int) 하면 안됨
dfGL['DMBTR'] = dfGL['DMBTR'] * 100
dfGL['WRBTR'] = dfGL['WRBTR'] * 100

#숫자로
#dfGL["HKONT"] = dfGL["HKONT"].astype(float).astype(int).astype(str)
dfGL["HKONT"] = dfGL["HKONT"].astype(str)

## 4) FSLine 추가

from mylib import myFileDialog as myfd
import pandas as pd

##################################################################
#계정과목 매핑을 읽는다.
filenameAcctMap = "acctMAP.xlsx"
#tgtdir = myfd.askdirectory()
filenameAcctMap = glob.glob(tgtdir+"/"+filenameAcctMap)
filenameAcctMap = filenameAcctMap[0]

dfAcc = pd.read_excel(filenameAcctMap)

dfAcc["DetailCode"] = dfAcc["DetailCode"].astype(str)
#dfGL["계정과목코드"] = dfGL["계정과목코드"].astype(str)

dfGL["계정과목코드"] = dfGL["계정과목코드"].astype(float).astype(int).astype(str)

dfGLJoin = dfGL.merge(dfAcc,how='left',left_on='계정과목코드',right_on='DetailCode')

# NA Test
dfGLJoin[dfGLJoin["FSName"].isna()] #Empty여야 함
dfGLJoin[dfGLJoin["FSName"].isna()]['계정과목코드'].drop_duplicates().to_csv("a.csv")

dfGL['계정과목명'] = dfGL['DetailName'] # FOR 신영
##################################################################


#EXPORT_FIN
dfGL.to_csv("./raw/rawGL_PY.txt", sep="\t", index=None)


