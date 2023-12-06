# Excel Par

Excel Par Sourcecode.
v0.1.9

# How to use

import ExcelPar as ep  
ep.run()

# History

v0.0.4, 231019 : 세부계정 월별순증분석 상단/하단 Logic 변경.  
v0.0.5, 231019 : FS Line 추가생성 구현  
v0.0.6, 231023 : 상하단 Logic 재변경_이예린SM (Mean ± 2Std)  
v0.1.2, 231104 : Improve with Parquet  
v0.1.3, 231104 : GL Slicing import 구현  
#0.1.5, 231108 : ADD README  

# README

## 설계도

모듈.메서드()  
실제 호출은 메서드를 바로 호출

1. TB 전처리부 : PreGL.RunPreTB() 
- 분석을 위해 rawTB를 읽어서 분석form으로 전처리 (어댑터 패턴)
- 사전에 지정된 excel form을 통해 필요한 변수를 입력받음

2. GL 전전처리부 (rawGL이 직접 Read할 수 없게 되어 있는 경우)

2-1. Excel 합치기 : ConcatExcelFiles.RunConcateExcelFiles()
- rawGL이 다수의 Excel파일(다수 시트 포함)일 경우 직접 읽어서 합치는 메서드
- pandas 사용
- 일체 전처리 없는 단순 합치기 (Column name은 trim함)

2-2. Text 합치기 : ConcatTextFiles.RunConcateTextFiles()
- rawGL이 다수의 Text파일일 경우 직접 읽어서 합치는 메서드
- pandas 사용
- 일체 전처리 없는 단순 합치기 (Column name은 trim함)

2-3. Text를 Slicing해서 합치기 : SliceAndSaveGL.RunSliceAndSaveGL()
- text를 필요한 컬럼만 slicing해서 합치는 메서드 => 가장 효율적
- dask 사용
- 대상컬럼은 현재 하드코딩되어 있으므로 향후 보완필요

2-4. BKPF+BSEG Join : JoinBKPFBSEG
- BKPF & BSEG를 다른 테이블로 주는 경우 join하는 구문
- 범용 tool이 아니므로 메서드로 구현하지는 않고 소스만 들어있음 (직접 호출 불가)

3. GL 전처리부 : PreGL.RunPreGL()
- 전전처리를 통해 읽을수 있게 된 rawGL을 읽어서 분석form으로 전처리 (어댑터 패턴)
- 사전에 지정된 excel form을 통해 필요한 변수를 입력받음
- FSLine 계정과목을 붙여준 후, TB/GL Recon data 자동 생성함 (실제 recon은 직접 엑셀에 붙여서 수행)

4. Excel Par Main분석부 : ExcelPar.RunEP()
- 전처리 완료된 TB/GL을 읽어서 분석 후 산출물 생성
- 최초코드를 리팩토링함

test