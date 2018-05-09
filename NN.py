# -*- coding:utf-8 -*-

import numpy as np
from sympy import *

init_printing(use_unicode=True)

x = Symbol('x')

X = np.array([[cos(x),-sin(x)],[sin(x),cos(x)]])
M = Matrix(X)

print('-'*32)
pprint(X)
print('-'*32)
pprint(M.T)
print('-'*32)
pprint(M.inv())
print('-'*32)
pprint(M.inv('LU'))
