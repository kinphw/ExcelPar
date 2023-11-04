#from ExcelPar.ExcelPar import ExcelPar
import ExcelPar as ep

match int(input(">>")):
    case 1:
        #Import TB
        ep.RunPreTB()
    case 2:
        #Deal with GL From RAW(tsv) To Parquet
        ep.RunSaveAsGL()
    case 3:
        #Import GL
        ep.RunPreGL()
    case 4:
        #CREATE EP
        ep.RunEP()
    case _:
        print("END")