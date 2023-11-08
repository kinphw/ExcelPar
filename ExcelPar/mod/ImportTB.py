import pandas as pd
from ExcelPar.mylib import myFileDialog as myfd

### 2-2. Import TB
class ImportTB:
    @classmethod
    def ImportTB(cls) -> pd.DataFrame:        
        fileName = myfd.askopenfilename("Select TB")
        tb = pd.read_csv(fileName, encoding="utf-8-sig", sep="\t")
        print("DONE")
        return tb