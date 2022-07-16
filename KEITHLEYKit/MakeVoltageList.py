import xlwt
import numpy as np

class voltage_list:
    """ This class of function is designed to create voltage list for voltage list sweeping on KEITHLEY. """
    def __init__(self):
        self.name = voltage_list

    # 这个函数可以生成一系列的电压点，以用于旧版KEITHLEY的电压列表扫描测试
    # voltage_node应为一个列表，其内容是扫描路径的端点，如要扫描0V-5V-0V-(-5V)-0V这样的路径，则应输入
    # voltage_node = [0,5,0,-5,0]或者[0,5,-5,0]
    def MakeVoltageList(self,voltage_node,step=0.25):
        num_node = len(voltage_node)                      # 电压端点个数
        num_seg = num_node-1                              # 路径段的个数等于电压端点的个数减一

        path_expanded = []                                # 将端点路径扩展成散点路径并保存在path_expanded中
        for i in range(num_seg):
            path = [voltage_node[i],voltage_node[i+1]]    # 将完整路径拆分各个电压小段, path[0]和path[1]分别为每段路径的起点与终点
            num_point = int(float(path[1]-path[0])/step)  # 每段路径的插点个数，先把路径长度转换为浮点数以免进行除法出错，最后再取整

            # 将每段路径扩展成散点路径
            point_list = []
            if num_point >= 0:                             # 如果点数是正的，意味着电压逐渐增加，则不需要额外操作
                for j in range(num_point):
                    point_list.append(path[0]+j*step)      # path[0]为路径起点，每个点相当于起点加一定倍数的步长
            else:
                num_point = abs(num_point)                 # 点数为负，意味着电压减少，为了迭代，我们先对点数取绝对值
                for j in range(num_point):
                    point_list.append(path[0]-j*step)      # 电压减小，所以是减
            path_expanded = path_expanded+point_list       # 进行列表的并接

        path_expanded.append(voltage_node[num_node-1])     # 把整个路径的终点加入path_expanded
        path_expanded = np.array(path_expanded)            # 将散点路径转换为数组
        num_point_total = len(path_expanded)               # 电压路径的总电压点数

        return path_expanded, num_point_total

    # 这个函数可以根据voltage_list数据生成excel文件以方便保存到旧KEITHLEY-4200中进行测试
    def WriteVoltageList(self,filename,saving_directory,voltage_list):
        file_address = saving_directory+'/'+filename+'.xls'  # 电压列表文件的绝对地址，xlwt包只能处理xls文件，故在此给出扩展名

        excel = xlwt.Workbook()                 # 创建excel workbook
        sheet = excel.add_sheet('Sheet1')       # 在workbook中创建sheet表格，命名为'Sheet1'

        sheet.write(0,0,'Voltage points')       # 写入表头，在第一行第一列写入'Voltage points
        sheet.write(0,1,'Number of points')     # 在第一行第二列写入'Number of points'，这一个单元格的下面就是总电压点数
        num_point = len(voltage_list)           # 电压点的个数
        sheet.write(1,1,num_point)              # 在第二行第二列写入总点数

        for i in range(num_point):
            sheet.write(i+1,0,voltage_list[i])  # 将各个电压点的值写入第一列

        excel.save(file_address)                # 将excel文件保存到特定地址

        return

if __name__ == '__main__':
    vl = voltage_list()

    directory = 'D:/Projects/PhaseTransistor/Data/I-V sweeping/VoltageList/test'

    a = vl.MakeVoltageList([0,10,-10,0],0.05)

    #print(a[1])
    #print(a[0])

    vl.WriteVoltageList('test',directory,a[0])