class testSin:
    _instance = None
    def __new__(class_, *args, **kwargs): #생성자 오버라이딩. #class : self
        if not isinstance(class_._instance, class_): #self의 인스턴스가 나랑 같은가? 아니면, 새로 만든다.
            class_._instance = object.__new__(class_, *args, **kwargs)
        return class_._instance        
    tmp = 1
    
testSin.tmp = 10

testSin1 = testSin()
testSin2 = testSin()

testSin.tmp = 8
testSin1.tmp

testSin2.tmp

from test import testSin


def test(list:list):
    print(list[0])
    print(list[1])

    return [list[0],list[1]]

tmp = ['a','b']

tt = test(['a','b'])

type(tt)

def test1(tmp):
    print(id(tmp))
    tmp = 10
    print(id(tmp))

tmp = 1
print(id(tmp))


    
test1(tmp)



def test2():
    nonlocal a
    print(a)


import pandas as pd

val = {
    'a':[1,2,3]
}

a = pd.DataFrame(val)

a

def test(a:pd.DataFrame):
    a.loc[1] = 3

test(a)

a
