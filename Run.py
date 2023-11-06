#from ExcelPar.ExcelPar import ExcelPar
import ExcelPar as ep
Text = """
1. Import TB
2. Preprocess GL(tsv to parquet) and Concatenate
3. Import GL
4. Run Excel Par
"""
print(Text)

match input(">>"):
    case '1':
        ep.RunPreTB()
    case '2':
        ep.RunSaveAsGL()
    #case '': #읽어서 전처리해서 숫자만 저장. 가장 효율적 코드
        #ep.CutGL.py 
    case '3':
        ep.RunPreGL()
    case '4':
        ep.RunEP()
    case _:
        print("END")