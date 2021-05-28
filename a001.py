import pandas as pd
import re         #引用re正则模块
import time
panda_name = input("请输入表格名称（须和代码文件在同一目录下，带上后缀名）：") # 获取表格名称
def y_1(x):
    try:
        y = pd.read_excel(x)  # 读取数据来检测
        return True   # 正确则返回True
    except FileNotFoundError:
        return False   #  查询不到否则返回False
if y_1(panda_name):
    nnn = 10   # no ture thing
else:
    print("文件名错误，请重试！")
    while True:
        panda_name = input("请输入表格名称（须和代码文件在同一目录下，带上后缀名）：") # 获取表格名称
        if y_1(panda_name):
            break   # 如果成功就跳出循环
        else:
            print("文件名错误，请重试！")        #  否则继续循环
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
                k_a = 3   #为了防止错误播报成功消息
                subject = 564984   # 防止语句错误执行
                # 以上代码是为了检测输入是否为整数或浮点数，下同

        if subject != "Chinese":
            if subject != "Maths":
                if subject != "English":
                    print("请输入正确的科目！")
                    k_a = 3     # 判断是否输入正确的科目
        if subject == "Chinese":
            Chinese[name] = mark
            Maths[name] = "nan"     # 覆盖性赋值
            English[name] = "nan"  # 防止出现访问错误
            k_a = 5     # 中心语句（此行是为了防止错误播报成功消息，下同）
        elif subject == "Maths":
            Maths[name] = mark
            Chinese[name] = "nan"
            English[name] = "nan"
            k_a = 5
        elif subject == "English":
            English[name] = mark
            Chinese[name] = "nan"
            Maths[name] = "nan"
            k_a = 5
        if k_a == 5:
            print("成功！")    # 播报成功消息


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
            subject = 0    # 防止执行后面的语句
        if subject == "Chinese":
            d_test=input(F"TA的分数是{Chinese[name]},您确定要继续吗？（Y/N）")   # 确认信息
            if d_test == "Y":
                Chinese.pop(name)   # 删除成绩
                print("成功！")
            else:
                print("操作已取消")    #  否则取消操作
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
            print("请输入正确的科目！")  # 检测是否是正确的科目

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
            print("人名不存在，请重试!")   # 检测人名是否存在
        if mark_h.isnumeric():
            mark_c = float(mark_h)
        else:
            value = re.compile(r'^[-+]?[0-9]+\.[0-9]+$')       # 定义正则表达式
            result = value.match(mark_h)
            if result:
                mark_c = float(mark_h)
            else:
                print("请输入正确的数字！")
                subject_1 = 46546846  # 确保是整数或浮点数（此行是为了防止执行后面的语句）
        if subject_1 == "Chinese":   # 判断科目
            Chinese[name_c] = mark_c   #  核心语句（赋值）
            print("成功修改！")   # 返回结果（下同）
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
                        print("请输入正确的科目！")   # 检测是否输入正确的科目

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
                    test_cg = 3     # 确保是整数或浮点数（此行是为了防止执行后面的语句）
            if test_cg == 5:
                mark_f = float(mark_cg)
                list_f = ["语文："]   # 首先添加科目名称便于标识，下同
                for i in Chinese:
                    if mark_f == Chinese[i]:
                        list_f.append(i)       #  如果符合，则添加相应元素
                if list_f[len(list_f)-1] == "语文：":
                    list_f.pop()
                list_f.append("数学：")
                for i in Maths:
                    if mark_f == Maths[i]:
                        list_f.append(i)
                if list_f[len(list_f)-1] == "数学：":
                    list_f.pop()
                list_f.append("英语：")
                for i in English:
                    if mark_f == English[i]:
                        list_f.append(i)
                if list_f[len(list_f)-1] == "英语：":
                    list_f.pop()
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
                print("人名不存在，请重试!")             #  检测此人是否存在
            if k == 5:
                print(F"{name_f}的语文成绩是{Chinese[name_f]}，数学成绩是{Maths[name_f]}，英语成绩是{English[name_f]}")
        if test_f == "S":
            subject_f = input("请输入学科名称：")
            if subject_f == "Chinese":
                for f in Chinese:
                        print(F"{f}，{Chinese[f]}分")  #先遍历字典得到value，再用value得到key
            if subject_f == "Maths":
                for f in Maths:
                        print(F"{f}，{Maths[f]}分")
            if subject_f == "English":
                for f in English:
                        print(F"{f}，{English[f]}分")
            if subject_f != "Chinese":
                if subject_f != "Maths":
                    if subject_f != "English":
                        print("请输入正确的科目！")   # 判断是否输入了正确的科目
    if act == "break":
        break      #   检测使用

    if act == "Export":
        ex_ch = input("请输入Excel表格名称：")
        k = 3
        test_last = ex_ch.split(".")   #  用.分割文件名（没有.则直接输出只有1个元素的列表）
        if ex_ch == "":
            ex_ch = time.strftime('%Y-%m-%d')   #  没有输入则用时间输出
            test_last_1 = ".xlsx"
            ex_ch = F"{ex_ch}{test_last_1}"    # 接上一个后缀名
        else:
            if test_last[len(test_last)-1] != "xlsx":
                print("文件名错误，请重试！")
                k = 5   #  如果既不为空也不以.xlsx结尾，则重新输入
                
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
        title = list(Chinese)     #  以上是为了判断用哪个字典获取人名

        title.insert(0,title.pop())  
        table = pd.DataFrame()
        table = table.append(Chinese,ignore_index=True)
        table = table.append(Maths,ignore_index=True)
        table = table.append(English,ignore_index=True)
        
        table = table[title]
        if k != 5:    # 执行正确的语句
            table.to_excel(ex_ch,index=False,header=True)
            print("成功！")
