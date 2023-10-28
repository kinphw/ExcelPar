if __name__=="__main__":

    test = "a"

    while True:    
        tmp = input("추가적인 가공이 필요하면 디버깅하세요. 아니면 1 입력>>")
        print(tmp)
        if tmp =='1' :
            break

    print(test)

import pandas as pd

dic = {
    'a':[1,2]
}

tmp = pd.DataFrame(dic)

tmp['a'].isna().any()

a = '2023-01-01'

a[:4]

bin(0b1111 & 0b1000) 

bin(0b100 & 0b111)

0b00001 & 0b10000000 > 0

TMP1 = 0b1111
TMP2 = 0b1000

bin(TMP1 & TMP2)



0b1111

switch:

tmp = 5

if tmp > 1:
    print("1")
if tmp > 2:
    print("2")

if tmp > 1: print("1")

t1 = 0b00
t2 = 0b10

TEST = t1 & t2

TEST = 0b10

def test(flag:bin):
    if 0b01 & flag:
        print("hello")
    else:
        print("adsf")
test(TEST)

from hehe import CONSTA

import pandas as pd

import ExcelPar.mylib.myFileDialog as myfd
path = myfd.askopenfilename()

tmp = pd.read_excel(path)

tt = {
    'a' : ['a', 'aa', 'abc']
}

tmp = pd.DataFrame(tt)

tmp['a'].apply(lambda x:x.replace('a','d'))
