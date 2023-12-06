# 전역선언부
import os
import glob
import tqdm
import ExcelPar.mylib.myFileDialog as myfd

# 함수부

class ConcatText:

    liTgt:list    
    filenameNew:str
    encodingOld:str
    encodingNew:str
    bHeader:bool

    def run(cls): #HANDLER        
        cls.getFiles()
        cls.filenameNew = input("new file name(result.txt)>>") or 'result.txt'                
        cls.encodingOld = input("encodingOld(cp949)>>") or 'cp949'
        cls.encodingNew = input("encodingNew(utf8)>>") or 'utf8'

        flag = input("헤더가 포함되어 있습니까?(기본값 Y). Y인 경우 첫번째 파일만 포함함)>>") or 'Y'
        if flag == 'Y': cls.bHeader = True
        else: cls.bHeader = False

        cls.concat(cls.bHeader)
        print("DONE")

    def getFiles(cls):
        path = myfd.askdirectory("합칠 텍스트파일이 있는 경로명")        
        ext = input("확장자(tsv)>>") or 'tsv'        
        cls.liTgt = glob.glob(path + '/*.'+ext)    #result.txt 생성 앞에 해야함

    def concat(cls, bHeader:bool): #bHeader가 True면 헤더 포함
        # 합산파일 생성
        fileNew = open(cls.filenameNew,'wt', encoding=cls.encodingNew)
        fileNew.close()

        # 파일 하나씩 append        
        fileNew = open(cls.filenameNew,'at', encoding=cls.encodingNew)
        
        pbar = tqdm.tqdm(total=len(cls.liTgt) , desc="순환")

        print("총",len(cls.liTgt),"개")
        
        totalj = 0
        i = 0 #i = File
        for file in cls.liTgt:                                  
            j = 0 #j = Line           
            fileOld = open(file, 'rt', encoding=cls.encodingOld)
            for line in fileOld:
                if bHeader and i!=0 and j==0: #헤더가 True이고, 첫번째 파일이 아닌 다른 파일인 경우, 첫번재 줄이면 생략함
                    j+=1
                    continue 
                tmp = fileNew.write(line)            
                j += 1
            fileOld.close()
            i += 1

            #pbar update
            pbar.desc = file #print 대체
            pbar.update(1)
            
            totalj += j
            print(file,":",totalj)            

        fileNew.close()
        pbar.close()


def runConcatText():
    ConcatText().run()

if(__name__ == "__main__"):
    runConcatText()  