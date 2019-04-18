import re
a='1 - 2 * ( (60-30 +(-40.0/5) * (9-2*5/3 + 7 /3*99/4*2998 +10 * 568/14 )) - (-4*3)/ (16-3*2) )'
#循环去除括号
while True:
    ret = re.split('\(([^()]+)\)', a, 1)
    if len(ret) == 3:
        a,b,c = re.split('\(([^()]+)\)', a, 1)
        rec = fun(b)
        a = a + str(rec) + c

    else:
        red = fun(a)
        print(red)
        break