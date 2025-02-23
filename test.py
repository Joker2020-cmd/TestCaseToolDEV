import xmind

# 加载现有的 XMind 文件
workbook = xmind.load("大模型应用_姓名_手机号.xmind")
sheet = workbook.getData()
#标题
title = sheet[0].get('topic').get('title')
title0 = title.split('_')
title1 = sheet[0].get('topic').get('topics')
#递归结束标记
T = 0
#递归函数
def rand1(title_1,T):
        title_1_name = dict(title_1[0]).get('title')
        print(title_1_name)
        x1 = dict(title_1[0]).get('topics')
        #print(x1)
        T=T+1
        if T != 5:
            return rand1(x1,T)
        else:
           return
#读取Xmind函数
def rand(title):
    for k in title:
        title_name =dict(k).get('title')
        print(title_name)
        x = dict(k).get('topics')
        #print(dict(x[0]))
        rand1(x,T)
        #return rand(x)
#
#
# #功能模块
for h in title1:
    Funtion_module = dict(h).get('title')
    print(Funtion_module)
    #print(h)
    rand(dict(h).get('topics'))
#    # print(dict(h).get('topics'))


#需求名称
Demand_name = title0[0]
#测试人
Tester = title0[1]
#手机号
phone_number = title0[2]

print(Demand_name+'---'+Tester+'---'+phone_number)
