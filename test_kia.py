import ExcelPar.PreGL_tmp as pg

file11 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\22_0\22_0.parquet".replace("\\","/")
file12 = "./imported/dfGL22_0.parquet"

file21 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\22_1\22_1.parquet".replace("\\","/")
file22 = "./imported/dfGL22_1.parquet"

file31 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\22_2\22_2.parquet".replace("\\","/")
file32 = "./imported/dfGL22_2.parquet"

file41 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\22_3\22_3.parquet".replace("\\","/")
file42 =  "./imported/dfGL22_3.parquet"

file51 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\23_1\23_1.parquet".replace("\\","/") 
file52 = "./imported/dfGL23_1.parquet"

file61 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\23_2\23_2.parquet".replace("\\","/")  
file62 = "./imported/dfGL23_2.parquet"

file71 = r"C:\Users\hyungwopark\OneDrive - Deloitte (O365D)\엑셀파_FY2023\10 Engagement별\231026_기아주식회사\23_3\23_3.parquet".replace("\\","/")   
file72 = "./imported/dfGL23_3.parquet" 

liImport = [file11,file21,file31,file41,file51,file61,file71]
liExport = [file12,file22,file32,file42,file52,file62,file72]
liYear = ['PY','PY','PY','PY','CY','CY','CY']

for i in range(7):
    pg.RunPreGL(liImport[i],liExport[i],liYear[i])

print("END")