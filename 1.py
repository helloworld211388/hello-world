import numpy as np
import openpyxl
import pandas as pd
workbook = openpyxl.load_workbook('1647848272130494 .xlsx')
worksheet = workbook['cph']



#用户信息列表，每个元素为一个列表[[usernum,'date',money]]
data = []
#用户编号列表
usernum = []
#缴费日期列表
date = []
#缴费金额列表
money = []

row_range = worksheet[2:665]
for row in row_range:
    usernum.append(row[0].value)
    date.append(str(row[1].value))
    money.append(row[2].value)


#将三个列表整合为data，列表中的每个元素为个用户的编码、缴费日期、缴费金额
for i in range(len(usernum)):
    data.append([])
    data[i].append(usernum[i])
    data[i].append(date[i])
    data[i].append(money[i])



#--------------------------任务1--------------------
#变量count储存excel表中不同元素个数，即用户个人
count = 1
for j in range(1,len(usernum)):
    if usernum[j] != usernum[j-1]:
        count += 1


#excel表中每一行数据代表一次缴费，故usernum的长度为总缴费次数，usernum中不同元素的个数（count）为用户人数
#我修改了平均缴费金额的计算方法
avenumpay = len(usernum)/count
aveamopay = sum(money)/len(usernum)
#创建一个dataframe数据
df = pd.DataFrame([[avenumpay,aveamopay]])
#写出csv文件
df.to_csv('居民客户的用电缴费习惯分析1.csv',header=['平均缴费次数','平均缴费金额'],index=False)


#-----------------------任务2-----------------
#高价值型客户
usertype11 = []
#大众型客户
usertype12 = []
#潜力型客户
usertype21 = []
#低价值型客户
usertype22 = []



#countlist用于存放各用户缴费次数
countlist = []
#moneylist用于存放各用户缴费金额
moneylist = []



#创建mylist列表用来存放usernum中的不同元素
mylist = []
for num in usernum:
    if num not in mylist:
        mylist.append(num)
        countlist.append(usernum.count(num))


#创建一个空列表new_mylist，其元素为count个空列表，每个列表用来存放各用户的各次缴费金额
new_mylist = []
for k in range(count):
    new_mylist.append([])


for m in range(len(data)):
    for n in range(count):
        if data[m][0] == mylist[n]:
            new_mylist[n].append(data[m][2])

#计算各个用户的平均缴费金额,sum(new_mylist[l])是单个用户总缴费金额，len(new_mylist[l])是单个用户总缴费次数
for l in range(count):
    moneylist.append(sum(new_mylist[l])/len(new_mylist[l]))



for user in range(count):
    if countlist[user] >= avenumpay and moneylist[user] >= aveamopay:
        usertype11.append(mylist[user])
    elif countlist[user] >= avenumpay and moneylist[user] < aveamopay:
        usertype12.append(mylist[user])
    elif countlist[user] < avenumpay and moneylist[user] >= aveamopay:
        usertype21.append(mylist[user])
    elif countlist[user] < avenumpay and moneylist[user] < aveamopay:
        usertype22.append(mylist[user])


print("高价值型客户有：",usertype11)
print("大众型客户有：",usertype12)
print("潜力型客户有：",usertype21)
print("低价值型客户有：",usertype22)
#将列表转化成数组，因为dataframe数据类型只能由数组类型合并而成
u11=pd.Series(usertype11)
u12=pd.Series(usertype12)
u21=pd.Series(usertype21)
u22=pd.Series(usertype22)
#将数组合并成dataframe类型
ys=pd.concat([u11, u12,u21,u22], axis=1)
#将dataframe类型数据写出成csv文件
ys.to_csv('居民客户的用电缴费习惯分析 2.csv',header=['高价值型客户','大众型客户','潜力型客户','低价值型客户'],index=False)