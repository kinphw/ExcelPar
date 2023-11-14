import time

class TimeCheck:
    __TimeTmp:float
    @classmethod
    def Set(cls):
        cls.__TimeTmp = time.time()
    @classmethod
    def Check(cls, msg:str = "Process"):
        TotalTime:float = time.time() - cls.__TimeTmp
        #print(f'{msg}: {TotalTime:.2f}초 소요') # RUN시는 삭제