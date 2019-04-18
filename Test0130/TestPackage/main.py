# -*- coding:utf-8 -*-
from TestPackage.test0130 import Student
import time

if __name__ == '__main__':
    lisa = Student('Lisa', 'female')
    lisa.get_gender()
    lisa.set_gender('male')
    lisa.get_gender()
    # str = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
    # print(str)
