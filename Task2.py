import copy
import pdb

value = [5461.7563, 5787.7928, 4791.393, 5942.5664, 53625.7064, 27051.516]
current_val = copy.copy(value)
arr_empty = []

for a in range(0,len(value)):
    for b in range(0,len(value)):
        for c in range(0,len(value)):
            for d in range(0,len(value)):
                for e in range(0,len(value)):
                    for f in range(0,len(value)):
                        if len(arr_empty) == 1000:
                            break
                        arr_empty.append([value[a], value[b], value[c], value[d], value[e], value[f]])

print(arr_empty)


import itertools
from itertools import product
from pdb import set_trace

value = [5461.7563, 5787.7928, 4791.393, 5942.5664, 53625.7064, 27051.516]
current_val = list(product(value, repeat=6))
final_val = []

# set_trace()
for a in current_val:
    if len(final_val) == 1000:
        break
    final_val.append(a)

print(final_val)

from openpyxl import Workbook

wb = Workbook()
ws = wb.active
for row in final_val:
    ws.append(row)

wb.save("Task_2.xlsx")