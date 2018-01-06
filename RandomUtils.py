import random
import string

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
    return  random.choice(list)