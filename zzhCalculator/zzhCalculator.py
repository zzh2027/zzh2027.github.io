#!/usr/bin/env python
# coding: utf-8

import time
import math

class Calculator(object):
    def __init__(self):
        print(">> 函数1:self.getSqrt(num, threshold)\n\t@num-->目标数字\n\t@threshold-->精度\nreturn:返回目标数字的平方根\n")
        print(">> 函数2:self.getLvl(num)\n\t@num-->目标数字\nreturn:返回目标数字的级数，如1000的级数为3\n")
        pass
    
    def getSqrt(self, num, threshold = 5):
        start = time.perf_counter()
        print(f'开始计算{num}的平方根,目标精度为{10**(-threshold)}')
        ans = self.__shrink(num, precision = threshold)
        end = time.perf_counter()
        my_cost = end-start
        print(f'计算耗时：{my_cost}秒')
        def check(target, my_ans):
            start = time.perf_counter()
            real_ans = math.sqrt(num)
            end = time.perf_counter()
            cost = end-start
            print(f'math包耗时{cost}秒', end = '')
            diff = abs(real_ans - my_ans)
            if not diff<10**(-threshold):
                print('函数精度不够！')
            return diff, cost
        diff, cost = check(num, ans)
        print(f"math包的速度是我的{my_cost/cost}倍")
        return ans
        
    def getLvl(self,num):
        ## 计算数字的级数
        lvl = 0
        while 10**lvl < num:
            lvl += 1
        return lvl-1
    
    def __shrink(self, target, rep = None, gradient = None, precision = None):
        """
        我们取平方根时会考虑到先取一半，但是如果数字本身很大，那么误差也会很大。
        此函数意在从一半往下走，直到误差可接受为止
        默认精度为0.00001
        """
        if gradient and gradient < 10**(-precision):
            return rep
        if not rep:
            rep = target*0.5
#         print('开始缩减')
        diff = rep**2 - target
#         print(f'>>>>初始误差:{diff}\t目标值:{target}')
        # 取级数
        lvl = self.getLvl(diff)
        if not gradient:
            diy = False
        else:
            diy = True
        while diff > 0 and rep > 0:
            if not diy:
                gradient = 10**int(0.5*(lvl-1))
#             print(f'当前值:{rep}\t梯度值:{gradient}')
#             print(f"梯度下降过程：")
            while diff > 10**(lvl-1):
                rep -= gradient
                diff = rep**2 - target
#                 print(f"\t值:{rep}\t|\tdiff:{diff}")
            lvl -= 1

#         if diff < 10**(-precision) and diff > -10**(-precision):
#             print(f'{target}的平方根约为{rep}')
#             return rep
#         else:
        new_gradient = gradient*0.1
#             print(f"新梯度：{new_gradient}")
        return self.__expand(target, rep, new_gradient, precision)

    def __expand(self, target, rep, gradient = None, precision = None):
        """
        经过shrink后，我们的数字从target的一半降到了很低，甚至导致rep**2 - target < 0。
        此函数需要根据这个num的数量级做一次新的递增
        """
        if gradient and gradient < 10**(-precision):
            return rep
#         print('\n\n开始增长')
        diff = rep**2 - target
#         print(f'>>>>初始误差:{diff}\t目标值:{target}') 
#         print(f'初始值:{rep} \t初始误差:{diff}')
#         print('当前值递增过程:')
        while diff < 0 and rep > 0:
            rep += gradient
            diff = rep**2 - target
#             print(f'\t当前值:{rep} \t当前误差:{diff}')
        new_gradient = gradient*0.1
#         print(f"新梯度：{new_gradient}\n\n")
        return self.__shrink(target, rep, new_gradient, precision)

if __name__ == '__main__':
    helper = Calculator()
    num = int(input('你想求哪个数字的平方根?'))
    ans = helper.getSqrt(num, 15)
    print(f"{num}的平方根约为{ans}")

