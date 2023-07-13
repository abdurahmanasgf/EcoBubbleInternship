import itertools
from pdb import set_trace as byebug

def amans_logic(*args, repeat=1):
    pools = [tuple(pool) for pool in args] * repeat

    result = [[]]
    for pool in pools:
        result = [x+[y] for x in result for y in pool]
    
    end_result = []
    final_result = []
    checks = ['RED', 'GREEN', 'BLUE', 'NIR735', 'NIR850', 'NIR880']

    for idx,prod in enumerate(result):
        current_data = prod[0:(len(prod)-1)]
        found=False
        for check in checks:
            if current_data.count(check) > 1:
                found=True
                break
        if found:
            continue
        print(f'No {idx}: {current_data}')
        end_result.append(current_data)

    for value in end_result:
        final_result.append("".join(value))

    return list(set(map(str, final_result)))

dadu = list(amans_logic(['RED', 'GREEN', 'BLUE', 'NIR735', 'NIR850', 'NIR880'], ['*', '/', '+'] ,repeat=6))


from openpyxl import Workbook

wb = Workbook()
ws = wb.active
for row in dadu[:1000]:
    ws.append(list(tuple(row.split(" "))))

byebug()
 
wb.save("Task_3.xlsx")



 