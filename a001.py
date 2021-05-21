import pandas as pd
import re         #引用re正则模块
panda_name = input("请输入表格名称（须和代码文件在同一目录下，带上后缀名）：") # 获取表格名称
def y_1(x):
    try:
        y = pd.read_excel(x)  # 读取数据来检测
        return True
    except FileNotFoundError:
        return False
if y_1(panda_name):
    nnn = 10
else:
    print("文件名错误，请重试！")
    while True:
        panda_name = input("请输入表格名称（须和代码文件在同一目录下，带上后缀名）：") # 获取表格名称
        if y_1(panda_name):
            break
        else:
            print("文件名错误，请重试！")        
data = pd.read_excel(panda_name)
data1 = data.head(2)   #　起始行（取一行作为value，单学科仅需一个成绩）
data2 = data1.drop(data1.index[[0]],axis=0)   # 删除不必要的行数
Maths = data2.to_dict('list')   # 将DataFrame对象转换成value为列表的字典
Maths.pop('成绩\\姓名')  # 删除最前面不需要的标签
for delta in Maths:       # 得到每一个key
    for ppt in Maths[delta]:  # 从key得到value，再转换为数字
        ppt1 = float(ppt)  # 将其转换为浮点类型方便计算
        Maths[delta] = ppt1   # 赋值给原来的对象
data1 = data.head(1)   # 下同
Chinese = data1.to_dict("list")   #  因为只提取了一行，所以不用删除
Chinese.pop("成绩\\姓名") 
for test1 in Chinese:
    for ppt in Chinese[test1]:
        ppt1 = float(ppt)  # 将其转换为浮点类型方便计算
        Chinese[test1] = ppt1
data1 = data.head(3)   # 英语在第3行，因此选3
data2 = data1.drop(data1.index[[0,1]],axis=0) # 前面两行需要删除，因此传入两个参数
English = data2.to_dict("list")
English.pop("成绩\\姓名")
for delta in English:
    for ppt in English[delta]:
        ppt1 = float(ppt)  # 将其转换为浮点类型方便计算
        English[delta] = ppt1
# 以上都是将Excel表格变成字典的过程

print("基本操作指南：a表示增，d表示删，c表示改，f表示查，break表示退出，Chinese表示语文，Maths表示数学，English表示英语，Export表示导出Excel表格") # 基本操作指南
while True:    
    act = input("请输入操作：")
    if act != "a":
        if act != "d":
            if act != "f":
                if act != "c":
                    if act != "break":
                        if act != "Export":
                            print("请输入正确的操作！")   # 检查输入是否正确
    if act == "a":
        name = input("请输入人名：")
        subject = input("请输入学科：")
        mark1  = input("请输入分数：")  # 输入信息
        if mark1.isnumeric():
            mark = float(mark1)
        else:
            value = re.compile(r'^[-+]?[0-9]+\.[0-9]+$')       # 定义正则表达式
            result = value.match(mark1)
            if result:
                mark = float(mark1)
            else:
                print("请输入正确的数字！")
                k_a = 3   #  为了防止错误播报成功消息
                subject = 564984   # 检查是否为数字或浮点数
        for a_55 in Chinese:
            if a_55 == name:                
                if subject == "Chinese":
                    Chinese[name] = mark
                    k_a = 5
                elif subject == "Maths":
                    Maths[name] = mark
                    k_a = 5
                elif subject == "English":
                    English[name] = mark
                    k_a = 5
                else:
                    if subject != 564984:      # 检查是否为数字或浮点数
                        print("请输入正确的科目！")       
                    k_a = 3
                if k_a == 5:
                    print("成功！")
                

    if act == "d":
        name = input("请输入人名：")
        subject = input("请输入学科：")
        for i in Chinese:
            if i == name:
                k = 5
                break
        else:
            k = 3
        if k !=5:
            print("人名不存在，请重试")           # 检查是否存在此人（下同）
            subject = 0
        if subject == "Chinese":
            d_test=input(F"TA的分数是{Chinese[name]},您确定要继续吗？（Y/N）")
            if d_test == "Y":
                Chinese.pop(name)
                print("成功！")
            else:
                print("操作已取消")
        elif subject == "Maths":
            d_test=input(F"TA的分数是{Maths[name]},您确定要继续吗？（Y/N）")
            if d_test == "Y":
                Maths.pop(name)
                print("成功！")
            else:
                print("操作已取消")
        elif subject == "English":
            d_test=input(F"TA的分数是{English[name]}，您确定要继续吗？（Y/N）")
            if d_test == "Y":
                English.pop(name)
                print("成功！")
            else:
                print("操作已取消")
        else:
            print("请输入正确的科目！")

    if act == "c":
        name_c = input("请输入人名：")
        subject_1 = input("请输入学科：")
        mark_h = input("请输入分数：")
        for i in Chinese:
            if i == name_c:
                k = 5
                break
        else:
            k = 3
        if k != 5:
            print("人名不存在，请重试!")
        if mark_h.isnumeric():
            mark_c = float(mark_h)
        else:
            value = re.compile(r'^[-+]?[0-9]+\.[0-9]+$')       # 定义正则表达式
            result = value.match(mark_h)
            if result:
                mark_c = float(mark_h)
            else:
                print("请输入正确的数字！")
                subject_1 = 46546846
        if subject_1 == "Chinese":
            Chinese[name_c] = mark_c
            print("成功修改！")
        if subject_1 == "Maths":
            Maths[name_c] = mark_c
            print("成功修改！")
        if subject_1 == "English":
            English[name_c] = mark_c
            print("成功修改！")
        if subject_1 != "Chinese":
            if subject_1 != "Maths":
                if subject_1 != "English":
                    if subject_1 != 46546846:
                        print("请输入正确的科目！")

    if act == "f":
        test_f = input("您想要以分数还是人名或科目查询？(M/N/S):")
        if test_f == "M":
            mark_cg = input("请输入分数:")
            if mark_cg.isnumeric():
                        mark_f = float(mark_cg)
                        test_cg = 5
            else:
                value = re.compile(r'^[-+]?[0-9]+\.[0-9]+$')       # 定义正则表达式
                result = value.match(mark_h)
                if result:
                    mark_f = float(mark_cg)
                    test_cg = 5
                else:
                    print("请输入正确的数字！")
                    test_cg = 3
                if test_cg == 5:
                    mark_f = float(mark_cg)
                    list_f = ["语文"]
                    for i in Chinese:
                        if mark_f == Chinese[i]:
                            list_f.append(i)
                    list_f.append("数学")
                    for i in Maths:
                        if mark_f == Maths[i]:
                            list_f.append(i)
                    list_f.append("英语")
                    for i in English:
                        if mark_f == English[i]:
                            list_f.append(i)
                    print("查询到的人：",list_f)
        if test_f == "N":
            name_f = input("请输入姓名：")
            for i in Chinese:
                if i == name_f:
                    k = 5
                    break
            else:
                k = 3
            if k != 5:
                print("人名不存在，请重试!")                
            if k == 5:
                print(F"{name_f}的语文成绩是{Chinese[name_f]}，数学成绩是{Maths[name_f]}，英语成绩是{English[name_f]}")
        if test_f == "S":
            subject_f = input("请输入学科名称：")
            if subject_f == "Chinese":
                for f in Chinese:
                        print(F"{f}，{Chinese[f]}分")
            if subject_f == "Maths":
                for f in Maths:
                        print(F"{f}，{Maths[f]}分")
            if subject_f == "English":
                for f in English:
                        print(F"{f}，{English[f]}分")
            if subject_f != "Chinese":
                if subject_f != "Maths":
                    if subject_f != "English":
                        print("请输入正确的科目！")
    if act == "break":
        break

    if act == "Export":
        ex_ch = input("请输入Excel表格名称（带上后缀名，允许不填）：")
        if ex_ch == "":
            bbb = 5
        
        Chinese.update({'成绩\姓名':'语文'})
        Maths.update({'成绩\姓名':'数学'})
        English.update({'成绩\姓名':'英语'})
        
        if len(Chinese) > len(Maths):
            if len(English) > len(Chinese):
                title = list(English)
            if len(English) <  len(Chinese):
                title = list(Chinese)
        if len(Chinese) < len(Maths):
            if len(Maths) > len(English):
                title = list(Maths)
            if len(Maths) < len(English):
                title = list(English)

        title.insert(0,title.pop())
        table = pd.DataFrame()
        table = table.append(Chinese,ignore_index=True)
        table = table.append(Maths,ignore_index=True)
        table = table.append(English,ignore_index=True)
        
        table = table[title]
        table.to_excel(ex_ch,index=False,header=True)
        print("成功！")
