"""
2020.01.30
Created by Yuance Xu
"""

import xlrd
import xlwt


## 初始化数据

### 定义老老鼠的类型、老鼠的编号、老鼠组织的编号

#### 老鼠的编号
mouse_typeOne = "A"
mouse_indexOne = [1,3,4,5]

mouse_typeTwo = "C" 
mouse_indexTwo = [1,2,3,4]

mouse_set_All = []
mouse_set_One = []
mouse_set_Two = []

for i in mouse_indexOne:
    mouse_index = mouse_typeOne + str(i)
    mouse_set_One.append(mouse_index)

for i in mouse_indexTwo:
    mouse_index = mouse_typeTwo + str(i)
    mouse_set_Two.append(mouse_index)

mouse_set_All = mouse_set_One + mouse_set_Two

#### 组织的编号
organ_index = ["p","i","cp","ci","T","S"] #组织的类型

mouse_organ_indexAll = []
mouse_organ_indexOne = []
mouse_organ_indexTwo = []

for m_index in mouse_indexOne:
    for o_index in organ_index:
        mouse_organ = mouse_typeOne + str(m_index) + "-" + o_index
        mouse_organ_indexOne.append(mouse_organ)

for m_index in mouse_indexTwo:
    for o_index in organ_index:
        mouse_organ = mouse_typeTwo + str(m_index) + "-" + o_index
        mouse_organ_indexTwo.append(mouse_organ)

mouse_organ_indexAll = mouse_organ_indexOne + mouse_organ_indexTwo


#### 定义老鼠的细胞数量

cell_length = 21
cell_index = []

for i in range(0,cell_length):
    index = str(i) + ":"
    cell_index.append(i)


## 获取输入数据
list_InputData = []
def InputData():
    wb = xlrd.open_workbook("data.xls") #打开文件
    data_sheet = wb.sheet_by_index(0)
    nrows = data_sheet.nrows #列数
    ncols = data_sheet.ncols #行数

    cell_value = data_sheet.col_values(1, start_rowx=1, end_rowx=None) #获取细胞数据
    data_value = data_sheet.col_values(3, start_rowx=1, end_rowx=None) #获取细胞数据对应的数值
    
    ###确保数据是正确的顺序
    if (len(cell_value) != len(data_value)) or (len(data_value) % cell_length != 0):
        print("data error")
        return
    else:
        print("start")

    for i in range(len(cell_value)):
        for mouse_Organ in mouse_organ_indexAll:
            if mouse_Organ in cell_value[i]:
                for j in range(0,cell_length):
                    cell_data = str(j) + ":" + mouse_Organ +  ":" + str(data_value[i+j])
                    list_InputData.append(cell_data)

    #print(list_InputData)            


## 重构数据
dict_DataMap = {}
def DataMap():
    for i in range(0,cell_length):
        index = str(i) + ":"
        index_data_list = []
        for data in list_InputData:
            if (index in data) and (data.find(index) == 0):
                t_index = data.find(":") + 1
                index_data = data[t_index:len(data)]
                index_data_list.append(index_data)

        dict_DataMap[i] = index_data_list

        if len(index_data_list) > 48:
            print(i)


## 输出数据
def OutputData():
    wb = xlwt.Workbook()

    for i in range(0,cell_length):
        sheet_name = "sheet" + str(i)
        f = wb.add_sheet(sheet_name)
        f.write(0,0,i)
        for j in range(len(mouse_set_All)):
            f.write(0,j+1,mouse_set_All[j])

        for j in range(len(organ_index)):
            f.write(j+1,0,organ_index[j])
                
        data_list = dict_DataMap[i]

        for n in range(len(organ_index)):
            for m in range(len(mouse_set_All)): 
                index = mouse_set_All[m] + "-" + organ_index[n]
                outputdata = ""
                for data in data_list:
                    if index in data:
                        t_index = data.find(":") + 1
                        outputdata = data[t_index:len(data)]

                f.write(n + 1,m + 1,outputdata)

    wb.save("output.xls")
    


if __name__ == '__main__':
    InputData()
    DataMap()
    OutputData()
