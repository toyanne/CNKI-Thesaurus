# 任务：把VRD的txt格式提取成有效的excel格式
# 键：序号 DE UF MT BT RT USE NT SN

import xlwt

with open('Virtual Reference Desk Thesaurus.txt', 'r', encoding="UTF-8") as f:
    data = f.readlines()

    len_data = len(data)

    num = []
    for j in range(1, 500):
        num.append(str(j))
    #print(num)

    index = []
    for i in range(0, len_data - 1):
        data[i] = data[i][1:-1]  # 删除换行符'\n'
        print(data[i])
        if data[i] in num:
            index.append(i)
    #print(index)#获得数字的索引

len_index = len(index)

# 类型 0 empty, 1 string, 2 number, 3 date, 4 boolean, 5 error
heading = ['No.', 'Word', 'Use', 'UF', 'BT', 'NT', 'RT']
file2 = xlwt.Workbook(encoding='utf-8', style_compression=0)
table = file2.add_sheet('VRD', cell_overwrite_ok=True)

# 写入数据table.write(行,列,value)
for i in range(0, 7):
    table.write(0, i, heading[i])  # 写入标题

for j in range(0, len_index):
    table.write(j + 1, 0, j + 1)  # 写入序号No.
    table.write(j + 1, 1, data[index[j] + 1])  # 写入词DE

for k in range(0, len_index - 1):  # 最后一个index没有k+1,所以这里减1
    uf = []
    bt = []
    nt = []
    rt = []
    for m in range(index[k], index[k + 1]):
        if 'Use' in data[m]:
            table.write(k + 1, 2, data[m][3:])
        elif 'UF' in data[m]:
            uf.append(data[m][3:] + '/')
            # table.write(k+1,3,uf)
        elif 'BT' in data[m]:
            bt.append(data[m][3:] + '/')
            # table.write(k+1,4,data[m][3:])
        elif 'NT' in data[m]:
            nt.append(data[m][3:] + '/')
        elif 'RT' in data[m]:
            rt.append(data[m][3:] + '/')
            # table.write(k+1,5,rt)

    table.write(k + 1, 3, uf)
    table.write(k + 1, 4, bt)
    table.write(k + 1, 5, nt)
    table.write(k + 1, 6, rt)

uf1 = []
bt1 = []
nt1 = []
rt1 = []
for n in range(index[len_index - 1], len_data):
    if 'Use' in data[n]:
        table.write(len_index, 2, data[n][3:])
    elif 'UF' in data[n]:
        uf1.append(data[n][3:] + '/')
        # table.write(k+1,3,uf)
    elif 'BT' in data[n]:
        bt1.append(data[n][3:] + '/')
        # table.write(k+1,4,data[m][3:])
    elif 'NT' in data[n]:
        nt1.append(data[n][3:] + '/')
    elif 'RT' in data[n]:
        rt1.append(data[n][3:] + '/')
        # table.write(k+1,5,rt)

table.write(len_index, 3, uf1)
table.write(len_index, 4, bt1)
table.write(len_index, 5, nt1)
table.write(len_index, 6, rt1)

# 保存文件
file2.save('file2.xls')
