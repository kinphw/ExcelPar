#전역부
try :
    import ExcelPar.mylib.myFileDialog as myfd
except Exception as e:
    print(e)
    import mylib.myFileDialog as myfd
import os
import glob
import pandas as pd
import time
###########################################
#경로 내 Text파일을 합치는 코드 by Pandas

class gc:
    path = ""
    ext = "*.tsv"

def SetGlobal():
#경로설정 및 리스트 추출
    print("폴더를 선택합니다.")
    gc.path = myfd.askdirectory()
    gc.list = glob.glob(gc.path + "/" + gc.ext)

    #리스트갯수 : 검증
    #len(list)
def Concat() -> pd.DataFrame:

#합산
    dfCon = pd.DataFrame()

    startTime = time.time()
    for i in gc.list:    
        nowTime = time.time()
        print(i," 시작:")
        #df = pd.read_excel(i)
        df = pd.read_csv(i,sep="\t", encoding="cp949", quotechar='"')
        df.columns = [i.strip() for i in df.columns] #추가코드. strip (TRIM) 왜곡될지도..
        dfCon = pd.concat([dfCon,df])
        print(i," 종료:")
        print("소요시간 : ",time.time() - nowTime)

    print("총 소요시간: ",time.time()-startTime)

    return dfCon

def Object2String(df:pd.DataFrame) -> None:
    #df를 받아서, object인 columns을 모두 string으로 변경해줌
    for column in df.columns:    
        if df[column].dtype == 'object':
            df[column] = df[column].astype(str) # Call by Object Refenece이므로 Return 불필요

def Export(dfCon:pd.DataFrame):

    fileName = input("input file name to save>>")    
    tmp = input("parquet이면 Y, 이외엔 tsv>>")    

    match tmp:
        case 'Y':
            Object2String(dfCon)
            dfCon.to_parquet(fileName)
        case _:
            dfCon.to_csv(fileName, sep='\t', index=False)
    print("End")    

def RunConcatTextFiles():
    SetGlobal()
    dfCon = Concat()    
    Export(dfCon)

if __name__=="__main__":
    RunConcatTextFiles()




