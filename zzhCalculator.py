#!/usr/bin/env python
# coding: utf-8

# In[50]:


def getSquaredRoot(num, threshold, gradient, direction = 'down'):
    
    ## 1. recursively
#     def smallHelper(res, num, accuracy, gradient):
#         diff = res**2 - num
#         while abs(diff) > 10**(-accuracy):
#             res += 10**(-gradient)
#             diff = res**2 - num
#             print(f'res:{res}; diff:{diff}')
#             if diff > 0:
#                 return bigHelper(res, num, accuracy, gradient+1)
#         return res
    
#     def bigHelper(res, num, accuracy, gradient):
#         diff = res**2 - num
#         while abs(diff) > 10**(-accuracy):
#             res -= 10**(-gradient)
#             diff = res**2 - num
#             print(f'res:{res}; diff:{diff}')
#             if diff < 0:
#                 return smallHelper(res, num, accuracy, gradient + 1)
#         return res
    ## 2. iteratively
    if threshold > 2:
        threshold = 10**(-threshold)
    if gradient > 2:
        gradient  = 10**(-gradient)
    def helper(res, num, threshold, gradient, direction = 'down'):
        diff = res**2 - num
        sign = diff > 0
        print(f'res:{res}; diff:{diff}')
        while abs(diff) > threshold:
            if direction == 'up':
                res += gradient
            else:
                res -= gradient
            diff = res**2 - num
            print(f'res:{res}; diff:{diff}')
            if sign != diff > 0:
                accuracy = 0
                while diff * 10**accuracy < 1:
                    accuracy += 1
                if accuracy == 0:
                    gradient *= 0.1
                gradient = 10**(-accuracy-1)
                sign = diff > 0
            if diff < 0:
                direction = 'up'
            else:
                direction = 'down'
        print(f'res:{res}; diff:{diff}')
        return res
    
#     if num == 1:
#         return 1
#     elif num < 4:
#         return smallHelper(num/2, num, 10, 6)
#     elif num == 4:
#         return 2
#     else:
#         return bigHelper(num/2, num, 10, 6)
    return helper(num/2, num, threshold, gradient, direction)



if __name__ == '__main__':
    getSquaredRoot(121, 30, 1, 'down')

