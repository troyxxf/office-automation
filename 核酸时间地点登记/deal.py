# encoding=gbk

import xlwt
import pandas as pd

#�Ǽ�������Ϣ
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
#�Ǽ������¼��Ϣ
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
#�ж��Ƿ������֣������ж�������Ϣ�Ƿ�Ϊ����Ҫ�ĵǼ���Ϣ
def isDigitIn(strs):
    for s in strs:
        if s.isdigit():
            return True
    else:
        return False
#������
if __name__ == '__main__':
#�ֵ����Ϣ
    dic_list=[]
    keys=("name","loc","time")
    name=get_name()
    # print(name)
#�Ȱ����ַŽ�ȥ
    for name_tmp in name:
        dic = dict.fromkeys(keys)
        dic["name"]=name_tmp
        dic_list.append(dic)
    info=get_info()
    # print(info)
#���ڴ滹û����Ϣ������
    Warning_name=[]
#ƥ�������������¼
    for name_index in range(len(name)):
        flag=0
        for info_index in range(len(info)):
            if name[name_index] in info[info_index]:
                info_tmp=info[info_index+1]
                #�ж��Ƿ�Ϊ��Ч��¼
                if isDigitIn(info_tmp):
                    info_tmp=info_tmp.split(" ")
                    if "." in info_tmp[0] or ":" in info_tmp[0]:
                        time_tmp=info_tmp[0]
                        loc_tmp=info_tmp[1]
                    else:
                        time_tmp=info_tmp[1]
                        loc_tmp=info_tmp[0]
                    #�޸�.Ϊ:
                    if "." in time_tmp:
                        time_tmp=time_tmp.replace('.',':')
                    #��������Ϣ�����ֵ�
                    dic_list[name_index]["time"]=time_tmp
                    dic_list[name_index]["loc"] = loc_tmp
                    flag=1
                    break
                else:
                    continue
        if flag==0:
            Warning_name.append(name[name_index])

    print("��δ��������",Warning_name)
    # print(dic_list)
#����Ϊexcel
    pf=pd.DataFrame(list(dic_list))
    #�޸ı��ͷ
    columns_map = {
        'name':'����',
        'loc':'�ص�',
        'time':'ʱ��'
     }
    pf.rename(columns=columns_map, inplace=True)
    #ָ�����ɵ�Excel�������
    file_path = pd.ExcelWriter('1.xlsx')
    #�滻�յ�Ԫ��
    pf.fillna(' ',inplace = True)
    #���
    pf.to_excel(file_path,encoding = 'utf-8',index = False)
    #������
    file_path.save()