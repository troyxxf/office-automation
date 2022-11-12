# encoding=gbk

import xlwt
import pandas as pd

#登记姓名信息
def get_name():
    name=[]
    with open("name1.txt","r",encoding='utf-8') as f:
        line=f.readline()
        while(line):
            name.append(line)
            line=f.readline()
    for i in range(len(name)):
        if "\n" in name[i]:
            name[i]=name[i][:len(name[i])-1]
    return name
#登记聊天记录信息
def get_info():
    info=[]
    with open("text1.txt","r",encoding='utf-8') as f:
        line=f.readline()
        while(line):
            info.append(line)
            line=f.readline()
    for i in range(len(info)):
        if "\n" in info[i]:
            info[i]=info[i][:len(info[i])-1]
    return info
#判断是否有数字，用于判断聊天消息是否为所需要的登记信息
def isDigitIn(strs):
    for s in strs:
        if s.isdigit():
            return True
    else:
        return False
#主函数
if __name__ == '__main__':
#字典存信息
    dic_list=[]
    keys=("name","loc","time")
    name=get_name()
    # print(name)
#先把名字放进去
    for name_tmp in name:
        dic = dict.fromkeys(keys)
        dic["name"]=name_tmp
        dic_list.append(dic)
    info=get_info()
    # print(info)
#用于存还没交信息的人名
    Warning_name=[]
#匹配人名与聊天记录
    for name_index in range(len(name)):
        flag=0
        for info_index in range(len(info)):
            if name[name_index] in info[info_index]:
                info_tmp=info[info_index+1]
                #判断是否为有效记录
                if isDigitIn(info_tmp):
                    info_tmp=info_tmp.split(" ")
                    if "." in info_tmp[0] or ":" in info_tmp[0]:
                        time_tmp=info_tmp[0]
                        loc_tmp=info_tmp[1]
                    else:
                        time_tmp=info_tmp[1]
                        loc_tmp=info_tmp[0]
                    #修改.为:
                    if "." in time_tmp:
                        time_tmp=time_tmp.replace('.',':')
                    #将所需信息存入字典
                    dic_list[name_index]["time"]=time_tmp
                    dic_list[name_index]["loc"] = loc_tmp
                    flag=1
                    break
                else:
                    continue
        if flag==0:
            Warning_name.append(name[name_index])

    print("还未核酸名单",Warning_name)
    # print(dic_list)
#导出为excel
    pf=pd.DataFrame(list(dic_list))
    #修改表格头
    columns_map = {
        'name':'姓名',
        'loc':'地点',
        'time':'时间'
     }
    pf.rename(columns=columns_map, inplace=True)
    #指定生成的Excel表格名称
    file_path = pd.ExcelWriter('1.xlsx')
    #替换空单元格
    pf.fillna(' ',inplace = True)
    #输出
    pf.to_excel(file_path,encoding = 'utf-8',index = False)
    #保存表格
    file_path.save()