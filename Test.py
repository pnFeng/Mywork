from Mylib.bench_resource import brsc
from ICs import PineHurst
from bit_manipulate import bm
import time
import re
import torch
import torch.nn as nn

a = '23238522'
patten1 = '2'
patten = '^((\d)3)\1[0-9](\d)\2{2}$'
match = re.findall(patten, a)
match1 = re.findall(patten1, a)

print(match)
print(match1)

data = [('the dog ate apples'.split(), ['DET', 'N', 'VV', 'DET', 'NN'])]
print(data)

rnn = nn.LSTM(10, 20, 2)
input = torch.randn(5, 3, 10)
h0 = torch.randn(2, 3, 20)
c0 = torch.randn(2, 3, 20)
output, (hn, cn) = rnn(input, (h0, c0))

list1 = [x + 100 for x in range(10)]
alpha = list(enumerate(list1))
for i, data in enumerate(list1):
    beta, gamma = data
    print(beta, gamma)
    pass
