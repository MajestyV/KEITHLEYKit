import os
from KEITHLEYKit import MakeVoltageList

# 这个函数可以用于创建文件夹
def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
        print("---  New folder...  ---")
        print("---  Created  ---")

    else:
        print("---  Folder already existed!  ---")
    return

MVL = MakeVoltageList.voltage_list()  # 调用函数包

directory = 'D:/Projects/PhaseTransistor/Data/I-V sweeping/VoltageList/Set1'  # 测试文件存放的主目录

Range = [1,2,3,4,5,6,7,8,9,10,15,20,25,30]        # 电压范围
Step = [0.01,0.025,0.05,0.1,0.15,0.2,0.25,0.5,1.0]  # 步长

for i in Range:
    sub_directory = directory + '/' + str(i)
    mkdir(sub_directory)

    for j in Step:
        voltage_list = MVL.MakeVoltageList([0,i,-i,0],j)[0]
        MVL.WriteVoltageList(str(i)+'V_step_'+str(j),sub_directory,voltage_list)

