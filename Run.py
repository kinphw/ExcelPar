#from ExcelPar.ExcelPar import ExcelPar
import ExcelPar as ep

match int(input(">>")):
    case 1:
        #Import TB
        ep.RunPreTB()
    case 2:
        #Import GL
        ep.RunPreGL()
    case 3:
        #CREATE EP
        ep.RunEP()
    case _:
        print("END")