import matplotlib.pyplot as plt
import numpy as np
import math

t = np.linspace(0, math.pi, 1000)
x = np.sin(t)
y = np.cos(t) + np.power(x, 2.0 / 3)  # 心型曲线的参数方程

plt.scatter(x, y, c=y, cmap=plt.cm.Reds, edgecolor='none', s=40)
plt.scatter(-x, y, c=y, cmap=plt.cm.Reds, edgecolor='none', s=40)  # 渐变颜色曲线
# 填充曲线
plt.fill(x, y, 'r', alpha=0.6)
plt.fill(-x, y, 'r', alpha=0.6)

plt.axis([-2, 2, -2, 2])  # 坐标轴范围
plt.title("WDS ♥ YYD", fontsize=30)
# 取消坐标轴显示
plt.axis('off')
# 保存文件
plt.savefig("love_Adi.png")  # 在 plt.show() 之前调用 plt.savefig()
plt.show()
