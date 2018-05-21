#任务：把UKAT的txt格式提取成有效的excel格式
#键：序号 DE UF MT BT RT USE NT SN

import xlwt
with open('UKATall.txt', 'r',encoding="UTF-8") as f:
    data=f.readlines()

    len_data=len(data)

    num = []
    for j in range(1, 23000):
        num.append(str(j))

    index=[]
    for i in range(0,len_data-1):
        data[i]=data[i][0:-1]#删除换行符'\n'
        #print(data[i])
        if data[i] in num:
            index.append(i)
            #print(data[i])
    #print(index)#获得数字的索引

len_index=len(index)

#d = {}
#类型 0 empty, 1 string, 2 number, 3 date, 4 boolean, 5 error
heading=['No.','DE','USE','UF','MT','BT','RT','NT','SN']
file2 = xlwt.Workbook(encoding='utf-8', style_compression=0)
table =file2.add_sheet('ukat_all',cell_overwrite_ok=True)

#写入数据table.write(行,列,value)
for i in range(0,9):
    table.write(0,i,heading[i])#写入标题

for j in range(0,len_index):
    table.write(j + 1, 0, j + 1)#写入序号No.
    table.write(j + 1, 1, data[index[j] + 1][3:])  # 写入词DE

for k in range(0,len_index-1):#最后一个index没有k+1,所以这里减1
    uf = []
    rt = []
    nt = []
    for m in range(index[k],index[k+1]):
        if 'USE' in data[m]:
            table.write(k+1,2,data[m][3:])
        elif 'UF' in data[m]:
            uf.append(data[m][3:]+'/')
            #table.write(k+1,3,uf)
        elif 'MT' in data[m]:
            table.write(k+1,4,data[m][3:])
        elif 'BT' in data[m]:
            table.write(k+1,5,data[m][3:])
        elif 'RT' in data[m]:
            rt.append(data[m][3:]+'/')
            #table.write(k+1,6,rt)
        elif 'NT' in data[m]:
            nt.append(data[m][3:] + '/')
        elif 'SN' in data[m]:
            table.write(k+1,8,data[m][3:])

    table.write(k + 1, 3, uf)
    table.write(k + 1, 6, rt)
    table.write(k + 1, 7, nt)

uf1 = []
rt1 = []
nt1 = []
for n in range(index[len_index-1],len_data):
    if 'USE' in data[n]:
        table.write(len_index, 2, data[n][3:])
    elif 'UF' in data[n]:
        uf1.append(data[n][3:]+'/')
    elif 'MT' in data[n]:
        table.write(len_index, 4, data[n][3:])
    elif 'BT' in data[n]:
        table.write(len_index, 5, data[n][3:])
    elif 'RT' in data[n]:
        rt1.append(data[n][3:]+'/')
    elif 'NT' in data[n]:
        nt1.append(data[m][3:] + '/')
    elif 'SN' in data[n]:
        table.write(len_index, 8, data[n][3:])

table.write(len_index, 3, uf1)
table.write(len_index, 6, rt1)
table.write(len_index, 7, nt1)

#保存文件
file2.save('file2.xls')