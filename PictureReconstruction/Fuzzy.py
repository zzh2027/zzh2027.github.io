#!/usr/bin/env python
# coding: utf-8


### read your picture


class SVDImg(object):
    def __init__(self, img):
        

        """
        原矩阵的转置与原矩阵的矩阵乘法，得到的矩阵，求其特征向量，其所有的特征向量便是右奇异矩阵【Right Singular Matrix】
        同理， 原矩阵与原矩阵的矩阵乘法，求得的矩阵的特征向量便是左奇异矩阵【Left Singular Matrix】
        """
        get_ipython().run_line_magic('matplotlib', 'inline')
        try:
            from loguru import logger
        except Exception as ex:
            logger.info(ex)
            get_ipython().system('pip install loguru')
        logger.add("Fuzz your pictures through SVD.log")    
        try:
            import PIL
        except Exception as ex:
            logger.info(f'{ex} do not have library "Image" here, installing....')
            get_ipython().system('pip install Image')

        import matplotlib.pyplot as plt
        import matplotlib.image as mpimg
        import numpy as np
        self.img = img
        if len(self.img.shape) == 3:
            y = self.img.shape[1]*self.img.shape[2]
            self.img_svd = self.img.reshape(-1, y)
        else:
            self.img_svd = self.img
        
    def _svd(self, ratio = 0.1):
        """
        n --> select part of the sigular values to reconstruct pictures
        """
        
        U, Sigma, VT = np.linalg.svd(self.img_svd)
        n = int(ratio*self.img.shape[0])
        try:
            return self.__reconstructImg(U, Sigma, VT, n)
        except:
            print("Try some smaller values less than {}!".format(self.img.shape[0]))
        
    def __reconstructImg(self, U, Sigma, VT, n):
        """
        the helper function to reconstruct the images but it takes a long time
        !!!DO NOT TRY TOO BIG PICTURES!!!
        Once tried a picture whose shape was (4320, 7680, 3) and my kernel went dead
        """
        res = (U[:, 0:n]).dot(np.diag(Sigma[0:n])).dot(VT[0:n, :])
        return res.reshape(len(U), int(len(VT)/3), 3)
    
    def _svd_2(self, ratio = 0.1):
        """
        self written functions based on the study, but it is never right, why??
        """
        n = int(ratio*self.img.shape[0])
        right = self.img_svd.T.dot(self.img_svd)
        left  = self.img_svd.dot(self.img_svd.T)
        rv, rsm = np.linalg.eig(right)
        lv, lsm = np.linalg.eig(left)
        sigma   = np.array(list(map(np.sqrt, rv)))
        return self.__reconstructImg(lsm, sigma, rsm, n)
    
    def show_img(self, ratio = 0.1):
        """
        use this function to fuzz your pictures
        """
        good_new = self._svd(ratio)
        new = self._svd_2(ratio)
        logger.info('Drawing pictures...')
        fig, ax = plt.subplots(1,3, figsize = [15,5])
        ax[0].imshow(self.img)
        ax[0].set_title('original picture')
        
        ax[1].imshow(good_new.astype(np.uint8))
        ax[1].set_title('fuzzed picture')
        
        ax[2].imshow(new.astype(np.uint8))
        ax[2].set_title('fucked picture')

if __name__ == 'main':
    test = mpimg.imread("test.jpg")
    s = SVDImg(test)    
    s.show_img(0.01)

