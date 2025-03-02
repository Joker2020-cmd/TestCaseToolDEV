import xmind
from itertools import zip_longest
import chXl

# 加载现有的 XMind 文件
workbook = xmind.load("大模型应用_姓名_手机号.xmind")
sheet = workbook.getData()
#标题
title = sheet[0].get('topic').get('title')
title0 = title.split('_')
title1 = sheet[0].get('topic').get('topics')
#数据
data0 = []
data1 = []
# data2 = []
# data3 = []
#递归结束标记
T = 0
#递归函数
def rand1(title_1, T, empty_list):
        title_1_name = dict(title_1[0]).get('title')
        print(title_1_name)
        empty_list.append(title_1_name)
        x1 = dict(title_1[0]).get('topics')

        #print(x1)
        T=T+1
        if T != 5:
            return rand1(x1,T, empty_list)
        else:
           return empty_list
#读取Xmind函数
def rand(title, Funtion_module):
    for k in title:
        title_name =dict(k).get('title')
        print(title_name)
        #data1.append(title_name)
        empty_list = []
        empty_list.append(Funtion_module)
        empty_list.append(title_name)
        x = dict(k).get('topics')
        #print(dict(x[0]))
        print(empty_list)
        g = rand1(x, T, empty_list)
        print(g)
        data1.append(g)
        print('*************')
        #return rand(x)
#
def split_list(data, num_parts):
    it = iter(data)  # 创建迭代器
    return [list(filter(None, group)) for group in zip_longest(*[it] * (len(data) // num_parts))]

#
# #功能模块
for h in title1:
    Funtion_module = dict(h).get('title')
    print(Funtion_module)
    data0.append(Funtion_module)
    #print(h)
    rand(dict(h).get('topics'), Funtion_module)
#    # print(dict(h).get('topics'))


#需求名称
Demand_name = title0[0]
#测试人
Tester = title0[1]
#手机号
phone_number = title0[2]

print(Demand_name+'---'+Tester+'---'+phone_number)

# for x in data0:
#     for t in data1:
#         empty_list = []
#         empty_list.append(x)
#         empty_list.append(t)
#         print(empty_list)
#         for f in split_list(data2, 4):
#             empty_list.extend(f)
#         data3.append(empty_list)


#split_arrays = split_list(data2, 4)
chXl.create_xl(data1, Demand_name, Tester, phone_number, len(data1))
#构造数据
print(data0)
print(len(data1))
#print(data2)
#print(data3)
#print(split_arrays)