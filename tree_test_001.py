#

from itertools import product
import numpy as np
import matplotlib.pyplot as plt

from sklearn import datasets
from sklearn.tree import DecisionTreeClassifier

# 使用iris数据
iris = datasets.load_iris()
X = iris.data[:, [0, 2]]
y = iris.target

# 训练模型, 限制树的最大深度为4
clf = DecisionTreeClassifier(max_depth=4)
clf.fit(X,y)

# Plot
x_min, x_max = X[:, 0].min() - 1, X[:, 0].max() + 1
y_min, y_max = X[:, 1].min() - 1, X[:, 1].max() + 1
xx, yy = np.meshgrid(np.arange(x_min, x_max, .1),
                    np.arange(y_min, y_max, .1))

Z = clf.predict(np.c_[xx.ravel(), yy.ravel()])
Z = Z.reshape(xx.shape)

plt.contourf(xx, yy, Z, alpha=.4)
plt.scatter(X[:, 0], X[:, 1], c=y, alpha=.8)
plt.show()
