import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # 空间三维画图
import random


def K_means(data, K):
    """
    程序说明：
    本函数实现二维和三维数据的K_means聚类算法
    data:输入的数据，维度(m, 2)或者(m, 3)
    K:表示希望分出来的类数
    """
    num = np.shape(data)[0]

    cls = np.zeros([num], np.int)

    random_array = np.random.random(size=k)
    random_array = np.floor(random_array * num)
    rarray = random_array.astype(int)
    print('数据集中随机索引', rarray)

    center_point = data[rarray]
    print('初始化随机中心点', center_point)

    change = True  # change表示簇中心是否有过改变，又改变了就需要继续循环程序，没改变则终止程序
    while change:
        for i in range(num):
            temp = data[i] - center_point  # 此句执行之后得到的是两个数或三个数：x-x_0,y-y_0或x-x_0, y-y_0, z-z_0
            temp = np.square(temp)  # 得到(x-x_0)^2等
            distance = np.sum(temp, axis=1)  # 按行相加，得到第i个样本与所有center point的距离
            cls[i] = np.argmin(distance)  # 取得与该样本距离最近的center point的下标

        change = False
        for i in range(k):
            # 找到属于该类的所有样本
            club = data[cls == i]
            newcenter = np.mean(club, axis=0)  # 按列求和，计算出新的中心点
            ss = np.abs(center_point[i] - newcenter)  # 如果新旧center的差距很小，看做他们相等，否则更新之。run置true，再来一次循环
            if np.sum(ss, axis=0) > 1e-4:
                center_point[i] = newcenter
                change = True

    print('K-means done!')
    return center_point, cls




def show_picture(data, center_point, cls, k):
    num, dim = data.shape
    color = ['r', 'g', 'b', 'c', 'y', 'm', 'k']
    if dim == 2:
        for i in range(num):
            mark = int(cls[i])
            plt.plot(data[i, 0], data[i, 1], color[mark] + 'o')

        # 下面把中心点单独标记出来：
        for i in range(k):
            plt.plot(center_point[i, 0], center_point[i, 1], color[i] + 'x')

    elif dim == 3:
        ax = plt.subplot(111, projection='3d')
        for i in range(num):
            mark = int(cls[i])
            ax.scatter(data[i, 0], data[i, 1], data[i, 2], c=color[mark])

        for i in range(k):
            ax.scatter(center_point[i, 0], center_point[i, 1], center_point[i, 2], c=color[i], marker='x')
    plt.show()


k = 24  ##分类个数
z_MF = []
yl_OSNR = []
pn = np.random.normal(0, 10, 6400)

for i in range(6400):
    index_y = random.uniform(15, 30)
    index_z = random.randint(0, 3)
    z_MF.append(index_z)
    yl_OSNR.append(index_y)

pn = pn[:, np.newaxis]
# print(pn)

y = np.array(yl_OSNR)
y = y[:, np.newaxis]
# print(y)

z_MF = np.array(z_MF)
z = z_MF[:, np.newaxis]

temp = np.hstack((pn, y))
# data = np.hstack((temp, z))
# print(data.shape)
center_point, cls = K_means(temp, k)

show_picture(temp, center_point, cls, k)
