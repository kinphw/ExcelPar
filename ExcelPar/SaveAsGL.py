########################################################
#다수의 Text GL을 READ하여 금액 전처리후 PARQUET로 저장 - Multifile support
# v0.0.1 DD 231104
########################################################

# 전역부

import os
import glob

import pandas as pd
import dask.dataframe as dd
from dask.diagnostics import ProgressBar
from dask.distributed import Client

try:
    from ExcelPar.mylib import myFileDialog as myfd
except Exception as e:
    from mylib import myFileDialog as myfd

path = ''

def DaskSet():
    # Client Setting
    cluster = Client(n_workers=4) #두뇌풀가동
    pbar = ProgressBar()
    pbar.register()

## 필요한 1회용 함수 선언부

def SetColumns() -> list:
    #숫자 전처리 필요한 컬럼을 지정하는 곳
    #일단 귀찮으니 수기지정하자. 나중에 입력으로 구현
    listColumn = ['차변','대변']
    return listColumn

def ColumnStrToInt(df:pd.DataFrame, ColumnName:str) -> None:    

    #숫자 뒤에 -를 붙여서 음수가 추출된 경우 앞으로 붙여주는 함수
    def InvertMinus(tgt:str) -> str:
        if len(tgt) <= 1 : return tgt #1글자거나(-) 1글자보다 적으면('') 그냥 바로 반환
        if tgt[-1] == '-':
            tgt = tgt.replace('-','')
            tgt = "-" + tgt
            return tgt
        return tgt

    try:
        #df[ColumnName] = df[ColumnName].fillna(0).apply(str.strip).apply(lambda x : x.replace(',','') ).apply(lambda x:x.replace('','0')).astype('float64').astype('int64')
        #df[ColumnName] = pd.to_numeric(df[ColumnName].fillna(0).astype('str').apply(str.strip).apply(lambda x : x.replace(',','')).apply(InvertMinus).replace( '[,)]','', regex=True).replace('[(]','-',regex=True))
        tmp = df[ColumnName].fillna(0) #일단 공백을 죽이고
        tmp = tmp.astype('str') #문자로 만든 다음
        tmp = tmp.apply(str.strip) #trim처리하고
        tmp = tmp.apply(lambda x: x.replace(',','')) #쉼표는 없앤다.
        tmp = tmp.replace('^-$', '0', regex=True) #그냥 오직 -는 0으로 바꿔준다.

        # 이후 로직은 위 아래 중 택일해서 돌려아함
        #tmp = tmp.apply(InvertMinus) 
        Flag = input("음수표시가 ()면 1, 마지막글자-면 2, 이외엔 그냥 엔터>>")
        match Flag:
            case '1':
                tmp = tmp.replace( '[,)]','', regex=True).replace('[(]','-',regex=True) #()를 마이너스로 바꿔주는 함수        
            case '2':
                tmp = tmp.apply(InvertMinus)
            case _:
                print("선택하지 않았습니다.")                

        tmp = pd.to_numeric(tmp)
        
        df[ColumnName] = tmp

    except Exception as e:
        print(e)        

def ReadGL(listColumn:list):

# 경로체크
    global path
    path = myfd.askdirectory()
    
    #ext = "DAT"
    ext = input("Enter Ext>>")

    list = glob.glob(path+"/"+"*."+ext)
    #file = list[0]
    #읽기
    #dfCon = pd.DataFrame()

    for file in list:
        print(file , "작업 개시")
        df = pd.read_csv(file, sep="\t", encoding='cp949', quotechar='"')
        
        #COLUMNS TRIM
        df.columns = df.columns.map(str.strip)

        #전처리 => 핵심로직
        for i in listColumn:
            print(i)
            ColumnStrToInt(df,i)
            df[i].sum()

        print(file , "전처리 완료")

        #파케이 저장을 위해 object to string

        for column in df.columns:
            #print(column)
            if df[column].dtype == 'object':
                df[column] = df[column].astype(str)

        
        #새로 저장(파케이)
        #filenameNew = file.replace(".DAT", ".tsv")    
        #filenameNewParquet = file.replace(".tsv", ".parquet")
        filenameNewParquet = os.path.splitext(file)[0] + ".parquet"
        df.to_parquet(filenameNewParquet)
        #df.to_csv(filenameNew,sep="\t",index=None, encoding='utf-8')
        
        #dfCon = pd.concat([dfCon,df])

        print(filenameNewParquet, "작업 완료.")

        #기존파일 삭제 => 삭제는 어렵지 않으니 따로 삭제
        #os.remove(file)
        #print(file , "삭제 완료")

################################
# 합쳐서 일괄 파케이로 저장
################################

# #전처리.. 짤라버린다
# df1 = pd.read_parquet(list[0])
# df2 = pd.read_parquet(list[1])
# df2 = df2.drop(columns=df2.columns[-2:])

# df2.to_parquet(list[1])
def ToParquet(listColumn:list):    
    # 일괄통합부 
    # LOAD
    global path
    list = glob.glob(path+"/"+"*.parquet")
    df = dd.read_parquet(path+"/"+"*.parquet")

    # VALIDATE
    for i in listColumn:
        print(i)
        print(df[i].sum().compute())
    # print(df['차변원화'].sum().compute())
    # print(df['대변원화'].sum().compute())

    df.compute().to_parquet(path+"/"+"ConcatenatedGL.parquet")
    print("DONE")

#MAIN
def RunSaveAsGL():
    listColumn = SetColumns()
    ReadGL(listColumn)
    DaskSet()
    ToParquet(listColumn)

if __name__=='__main__':    
    RunSaveAsGL()


