##############################################
# Parquet의 Columns을 수기로 변경해주는 Class
# v0.0.1
# 향후 변경 : 그냥 데이터프레임을 매개변수로 받아서 (Call by object reference) 컬럼명만 바꾸는 쪽으로 간략화
##############################################
import os

import pandas as pd
import ExcelPar.mylib.myFileDialog as myfd

class ChangeColumns:
    
    #Declare Class var
    df:pd.DataFrame
    ColumnsFirst:list
    ColumnsOld:list
    ColumnsNew:list
    ColumnOld:str
    ColumnNew:str

    def __init__(self, df:pd.DataFrame = None):
        print("Created...")
        if df is not None:
            self.df = df
    
    def __del__(self):
        pass

    def Read(self)->pd.DataFrame:
        fileName = myfd.askopenfilename("Select file")
        match os.path.splitext(fileName)[1]: #확장자에 따라
            case '.txt':
                print("TSV_UTF-8")
                self.df = pd.read_csv(fileName, sep='\t')
                self.__InitSet()
            case '.parquet':
                print("Parquet")
                self.df = pd.read_parquet(fileName)
                self.__InitSet()                
            case _:
                print("읽을 수 없었습니다...")                

    def __InitSet(self):
        self.ColumnsFirst = self.df.columns.to_list()
        self.ColumnsOld = self.df.columns.to_list().copy()
        self.ColumnsNew = self.ColumnsOld.copy() #처음에 복사해놓는다.     
        print("Done. Call Change().")   
    
    def Change(self)->pd.DataFrame:        
        while True:
            print("change begin")
            
            print(self.ColumnsNew)
            
            self.ColumnOld = input("바꿀 Column (Exit : 'exit')>>")

            match self.ColumnOld:
                case 'exit':                    
                    print("종료합니다. 추출하고자 한다면 Call Export()")                    
                    break
                case _:
                    self.__Loop()

    def __Loop(self):
            #순환하며 찾는다
            try:
                idx = self.ColumnsOld.index(self.ColumnOld) #인덱스를 찾는다
                self.ColumnNew = input("Found. New Column Name?>>")
                self.ColumnsNew[idx] = self.ColumnNew #ColumnsNew를 바꾼다.
                self.__Rename()
                print("Changed.", self.ColumnOld, "->", self.ColumnNew)
                #print(self.ColumnsNew)
            except Exception as e:
                #오류가 났다는 것은 없다는 것임
                print(e)
                print("No Column Found")

    def Export(self) -> pd.DataFrame:
        print("ColumnName이 변경된 DataFrame을 반환합니다.")
        return self.df

    def __Rename(self):
        #di = dict(zip(self.ColumnsOld, self.ColumnsNew))
        #self.df.rename(columns=di) ## FOR DD CODE
        self.df.columns = self.ColumnsNew
        self.ColumnsOld = self.df.columns.to_list().copy() #RESET
        self.ColumnsNew = self.ColumnsOld.copy() #RESET

