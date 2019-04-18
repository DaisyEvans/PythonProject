# 斐波那契数列

# nterms = int(input('你需要几项？'))
#
# n1 = 0
# n2 = 1
# count = 2
#
# if nterms <= 0:
#     print('请输入正整数！')
# elif nterms == 1:
#     print('斐波那契数列：%d' % n1)
# else:
#     print('斐波那契数列：')
#     print(n1, ',', n2, end=' , ')
#     while nterms > count:
#         nth = n1 + n2
#         print(nth, end=' , ')
#         n1 = n2
#         n2 = nth
#         count += 1

nterms = int(input('需要查看几个月的兔子数(对)：'))
a = 0
b = 1
i = 0
for i in range(nterms):
    print('第%d个月的兔子数(对)：%d' % (i + 1, b))
    a, b = b, a + b