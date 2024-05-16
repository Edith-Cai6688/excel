import openpyxl
import pandas as pd


file1='D:/BaiduSyncdisk/Corita/乱七八糟/python/new/综测汇总.xlsx/'
file2='D:/BaiduSyncdisk/Corita/乱七八糟/python/new/2.xlsx'
l1='劳动值'
l2='劳动值-日常劳动分'
l3='寝室得分'
l4='劳动值-志愿服务分'
l5='志愿服务得分'
l6='劳动值-实习实训分'


def getdata1( file, row):
    data = pd.read_excel(file, skiprows=range(1,row), nrows=1, usecols=['学号', '姓名', '日常劳动得分L1'])
    data = data.fillna(0)
    data = data[['学号', '姓名', '日常劳动得分L1']]
    # column_data=pd.Series(l1,index=data.index)
    data.insert(2, '一级指标名称', l1)
    data.insert(3, '二级指标名称', l2)
    data.insert(4, '事项名称', l3)
    data_list1 = data.iloc[0].values.tolist()
    print(data_list1)
    return data_list1


def getdata2( file, row):
    data = pd.read_excel(file, skiprows=range(1, row), nrows=1, usecols=['学号', '姓名', '志愿服务得分L2'])
    data = data.fillna(0)
    data = data[['学号', '姓名', '志愿服务得分L2']]
    # column_data=pd.Series(l1,index=data.index)
    data.insert(2, '一级指标名称', l1)
    data.insert(3, '二级指标名称', l4)
    data.insert(4, '事项名称', l5)
    data_list2 = data.iloc[0].values.tolist()
    # print(data_list2)

    return data_list2


def getdata3( file, row):
    data = pd.read_excel(file, skiprows=range(1, row), nrows=1, usecols=['学号', '姓名', '技能与专业证书', '技能与专业证书得分L3'])
    data = data.fillna(0)
    data = data[['学号', '姓名', '技能与专业证书', '技能与专业证书得分L3']]
    # column_data=pd.Series(l1,index=data.index)
    data.insert(2, '一级指标名称', l1)
    data.insert(3, '二级指标名称', l6)
    data_list3 = data.iloc[0].values.tolist()
    # print(data_list2)
    return data_list3


def addlist(row):
    list = [getdata1(file1, row), getdata2(file1, row), getdata3(file1, row)]
    print(list)
    return list


def writeonedata( file2, row):
    new_data = pd.DataFrame(addlist(row))
    existing_data = pd.read_excel(file2, sheet_name='Sheet1')
    updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    print(updated_data)
    updated_data.to_excel(file2, index=False, header=True)


def allwrite():
    for x in range(1, 406):
        writeonedata(file2, x)

if __name__=='__main__':
    # pd.set_option('display.max_columns', None)
    # getdata1(file1,5)
    # addlist(6)
    # openexcel().writeonedata(file2,1)
    allwrite()


