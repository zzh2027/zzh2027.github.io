#!/usr/bin/env python
# coding: utf-8

# In[11]:


class Painting(object):
    def __init__(self):
        print('函数1：self.red_heart(num)\n\t@num-->红心维度')
        pass
    
    def red_heart(self, num):
        #num = int(input("Enter your number here~"))
        for i in range(num):
            for j in range(2*(num - i -1)):
                print("\033[1;31;46m ",end = "")
            for j in range(i + 1):
                print("\033[1;31;41m    ",end = "")
            for j in range(4*(num-i-1)):#0,1,2,3  --> 6,4,2,0
                print("\033[1;31;46m ",end= "")
            for j in range(i+1):
                print("\033[1;31;41m    ",end = "")
            for j in range(2*(num - i -1)):
                print("\033[1;31;46m ",end = "")
            print()
        for i in range(num//2):
            for j in range(2*num):
                print("\033[1;31;41m    ",end = "")
            print()
        for i in range(2*num):
            for j in range(i):
                print("\033[1;31;46m  ",end = "")
            for j in range(2*num-i):
                print("\033[1;31;41m    ",end = "")
            for j in range(i):
                print("\033[1;31;46m  ",end = "")
            print()
            
if __name__ == '__main__':
    a = Painting()
    a.red_heart(9)


# In[ ]:




