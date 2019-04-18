# 输出指定范围内素数
import math
a = int(input('素数下限：'))
b = int(input('素数上限：'))
print('%d-%d以内素数：' % (a, b))
for i in range(a, b):
    for j in range(2, round(math.sqrt(i)) + 1):
        if i % j == 0:
            break
    else:
        print(i, end=',')
