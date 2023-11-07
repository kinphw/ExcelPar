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
        #ep.RunSaveAsGL() #deprecated    
        ep.RunSliceAndSaveGL()        
    case '3':
        ep.RunPreGL()
    case '4':
        ep.RunEP()
    case _:
        print("END")