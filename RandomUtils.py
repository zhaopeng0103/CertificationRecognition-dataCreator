#!/usr/bin/env python3
# encoding=UTF-8
import random
import string
import time


# 生成随机整数
def produceRandomInt(min, max):
    return random.randint(min, max)


# 生成随机偶数
def produceRandomEven(min, max):
    return random.randrange(min, max, 2)


# 生成随机浮点数
def produceRandomFloat(min, max):
    return random.uniform(min, max)


# 选取字符串中的随机字符
def produceRandomChar(char):
    return random.choice(char)


# 多个字符中选取特定数量的字符
def produceRandomChars(char, num):
    return random.sample(char, num)


# 多个字符中选取特定数量的字符组成新字符串
def joinRandomString(list, num):
    return string.join(random.sample(list, num)).replace(" ", "")


# 随机选取字符串
# list:['apple', 'pear', 'peach', 'orange', 'lemon']
def produceRandomString(list):
    return random.choice(list)


# 生成随机日期字符串
def produceRandomTime(startYear, endYear):
    a1 = (startYear, 1, 1, 0, 0, 0, 0, 0, 0)  # 设置开始日期时间元组（2020-01-01 00：00：00）
    a2 = (endYear, 12, 31, 23, 59, 59, 0, 0, 0)  # 设置结束日期时间元组（2050-12-31 23：59：59）

    start = time.mktime(a1)  # 生成开始时间戳
    end = time.mktime(a2)  # 生成结束时间戳

    t = produceRandomInt(start, end)  # 在开始和结束时间戳中随机取出一个
    date_touple = time.localtime(t)  # 将时间戳生成时间元组
    return time.strftime("%Y-%m-%d", date_touple)
