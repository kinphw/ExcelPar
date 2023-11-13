import time

class SetGlobal:
    _instance = None
    def __new__(class_, *args, **kwargs): #생성자 오버라이딩. #class : self
        if not isinstance(class_._instance, class_): #self의 인스턴스가 나랑 같은가? 아니면, 새로 만든다.
            class_._instance = object.__new__(class_, *args, **kwargs)
        return class_._instance        
    
    #Class Var. - First set
    PM = '0'
    De_minimis = '0' #Analyze에서 사용
    diff_비율 = 0
    ClientNameDate = 0
    Level = 0
    
    #Class Var. - In src
    분석계정과목 = []
    검토문장 = []

    #bool
    bDask = True #Dask 사용여부

    #Time For debugging
    #timeStart:float

    #Setter
    @classmethod
    def SetGlobal(cls):
        PM = 1000000000 # 10억을 기준으로 함
        cls.PM = input("적용할 PM을 입력하세요 > ") or '6,285,000,000'
        try:
            cls.PM = cls.PM.replace(",","")
        except:
            pass
        cls.PM = int(cls.PM)
        print(f'입력하신 PM은 {cls.PM:,}입니다.')

        #De_minimis = 200000000
        cls.De_minimis = input("적용할 CTT를 입력하세요 > ") or '551,000,000'
        try:
            cls.De_minimis = cls.De_minimis.replace(",","")
        except:
            pass
        cls.De_minimis = int(cls.De_minimis)
        print(f'입력하신 CTT는 {cls.De_minimis:,}입니다.')

        cls.ClientNameDate = input("파일명에 반영할 회사명/기준월을 입력하세요. 파일명에만 영향을 줍니다. (ex. 삼성전자2309)> ") or '강원랜드2309'
        print(f'입력하신 회사명/기준월은 {cls.ClientNameDate}입니다.')

        # cls.diff_비율 = 0.2
        # print(f'기본 차이비율 Threshold는 {cls.diff_비율:.0%}')

        cls.Level = 'Detail' #기본값 : Detail

        # FOR DEBUG
        #cls.timeStart = time.time()
    #####################################################################