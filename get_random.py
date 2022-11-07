import random
import time
from decimal import *

test = ['a', 'b', 'c', 'd', 'e', 'f', '若云', 'h', 'i', 'j', 'k', 'l', 'm', 'n']

# 概率获取
# 主名单表 main_list
# 正在抽取礼物的人 + 已获得礼物的人名单表 get_gift_list
# 逻辑上从 main_list - get_gift_list 来平分概率
def get_pr(main_list: list, get_gift_list, num_of_gift_people):
    pr_result = {}
    main_length = len(main_list)
    n = main_length - num_of_gift_people
    pr = (Decimal('1.0000') / Decimal(n)).quantize(Decimal('.0001'))
    remain_pr = Decimal('1.0000')           # 概率修正
    for i in range(main_length):
        if main_list[i] in get_gift_list:
            pr_result[main_list[i]] = Decimal('0.0000')
        else:
            pr_result[main_list[i]] = pr
            remain_pr -= pr
    pr_result['remain'] = remain_pr
    return pr_result

def get_pr_1(main_list: list, get_gift_list, num_of_gift_people):
    pass


# 第二种：随机抽取
def function_random(people):
    random.seed(time.time())
    return random.choice(people)

# 第三种：调整概率
def function_adjustable(listA: list, name: str):
    random.seed(time.time())
    temp = random.random()

    # step1. 构造概率映射
    n = len(listA)
    safufu_index = listA.index(name)
    pr = (Decimal('1.0000') / Decimal(n)).quantize(Decimal('.0001'))
    safufu_pr = pr - Decimal('0.01')
    other_pr = ((Decimal('1.0000') - safufu_pr) / Decimal(n - 1)).quantize(Decimal('.0001'))
    pr_dict = {}
    remain_pr = Decimal('1.0000')
    for i in range(n):
        if i == safufu_index:
            pr_dict[listA[i]] = safufu_pr
            remain_pr -= safufu_pr
        else:
            pr_dict[listA[i]] = other_pr
            remain_pr -= other_pr
    pr_dict[listA[safufu_index]] += remain_pr      # 概率补全


    # 先构造概率数组
    # safufu_index = listA.index(name)
    # safufu_probability = "{:.4f}".format(probability - 0.01)
    # print(safufu_probability)
    # probability_list = [0.00] * (n + 1)
    # other_probability = "{:.4f}".format((1 - decimal(safufu_probability)) / (n - 1))
    # print(other_probability)
    # remain_probability = 1.0000
    # print(probability_list)
    # for i in range(n):
    #     if i == safufu_index:
    #         probability_list[i] = float(safufu_probability)
    #         remain_probability -= float(safufu_probability)
    #     else:
    #         probability_list[i] = float(other_probability)
    #         remain_probability -= float(other_probability)

function_adjustable(test, "若云")
