
import xlrd
import xlwt

heading = ['No.', 'Word', 'Use', 'UF', 'BT', 'NT', 'RT', 'French', 'SC', 'SN']
file5 = xlwt.Workbook(encoding='utf-8', style_compression=0)
table = file5.add_sheet('CA', cell_overwrite_ok=True)

workbook = xlrd.open_workbook('C:\\Users\\KATHY\\Desktop\\分类汇总\\政府&组织Government&Organization\\CA.xlsx')
sheet = workbook.sheet_by_index(0)
cols = sheet.col_values(0)
words = []

str = []
for i in cols:
    s = i.split(',')  # 将大字符串切分为词
    str.append(s)
    word = s[0]  # 获得词

    words.append(word)
    # print(word)
w = []
for j in words:
    if j not in w:
        w.append(j)  # 去除重复元素
# print(w[29])
le = len(w)

for k in range(0, 10):
    table.write(0, k, heading[k])  # 写入标题

for m in range(0, le):
    table.write(m + 1, 1, w[m])  # 写入词
    uf = []
    bt = []
    nt = []
    rt = []
    sc = []
    for n in str:
        if w[m] in n:
            if '"Use"' in n:
                table.write(m + 1, 2, n[2])
            elif '"Used For"' in n:
                uf.append(n[2]+ '/')
            elif '"Broader Term"' in n:
                bt.append(n[2]+ '/')
            elif '"Narrower Term"' in n:
                nt.append(n[2]+ '/')
            elif '"Related Term"' in n:
                rt.append(n[2]+ '/')
            elif '"French"' in n:
                table.write(m + 1, 7, n[2])
            elif '"Subject Category"' in n:
                sc.append(n[2]+ '/')
            elif '"Scope Note"' in n:
                table.write(m + 1, 9, n[2])
    table.write(m + 1, 3, uf)
    table.write(m + 1, 4, bt)
    table.write(m + 1, 5, nt)
    table.write(m + 1, 6, rt)
    table.write(m + 1, 8, sc)

        #else:
            #print('F')
        #if w[m] in n and n[1] is '"Use"':
            #table.write(m + 1, 2, n[2])

file5.save('file5.xls')
