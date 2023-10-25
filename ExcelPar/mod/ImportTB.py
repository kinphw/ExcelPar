import pandas as pd
from ExcelPar.mylib import myFileDialog as myfd

### 2-2. Import TB
class ImportTB:
    @classmethod
    def ImportTB(cls) -> pd.DataFrame:        
        tb = pd.read_csv(myfd.askopenfilename(), encoding="utf-8-sig", sep="\t")
        print("DONE")
        return tb