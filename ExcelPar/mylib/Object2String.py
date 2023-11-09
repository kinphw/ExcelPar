## 별도 함수로 분리

import pandas as pd

# to save parquet. 필요시 사용
def Object2String(df:pd.DataFrame) -> None:
    #df를 받아서, object인 columns을 모두 string으로 변경해줌
    for column in df.columns:    
         if df[column].dtype == 'object':
              df[column] = df[column].astype(str) # Call by Object Refenece이므로 Return 불필요