import xlrd
import re
import matplotlib.pyplot as plt
import numpy as np

class geda:
    """This module is designed to extract testing data generated by KEITHLEY 4200."""

    def __init__(self):
        self.name = geda

    # This function is written to extract all measurement data from a data file.
    def Extract(self,Datafile):
        data = xlrd.open_workbook(Datafile)

        sheet_list = data.sheet_names()
        data_sheets = []
        for n in sheet_list:
            if n == 'Data' or ('Append' in n):
                data_sheets.append(n)

        total_data_dict = dict.fromkeys(tuple(data_sheets))  # The input of dict.fromkeys is tuple.
        # The following loop is written to extract all the testing results from ever sheet.
        # And then, distributing them to their corresponding key in the total_data_dict.
        for n in data_sheets:
            testing_result = data.sheet_by_name(n)
            parameters = testing_result.row_values(0)  # Extracting the parameter names from the first row of the sheet.
            data_dict = dict.fromkeys(parameters)
            for m in range(testing_result.ncols):
                result = testing_result.col_values(m)
                del result[0]  # Deleting the name of parameters from the test data.
                data_dict[parameters[m]] = result

            total_data_dict[n] = data_dict

        return total_data_dict,data_sheets

    # This function uses Extract() to extract data from different datafiles and merge them into a dict.
    # Format of Plotting_list: [[Datafile 1, [specific curve list]], [Datafile 2, [specific curve list]], ......]
    def Merge(self,Plotting_list):
        plotting_dict = {}
        for n in range(len(Plotting_list)):
            datafile, curve_list = Plotting_list[n]
            data_dict = self.Extract(datafile)[0]
            plotting_dict[datafile] = {}
            for m in curve_list:
                plotting_dict[datafile][m] = data_dict[m]
        return plotting_dict

    # This function is written to visualize data received from KEITHLEY 4200.
    # The settings of the plot, packed inside dictionary kwargs, are the same as matplotlib.
    # The curve_type variable indicates the characteristic of the curve.
    def Visualize(self,Plotting_list,curve_type,**kwargs):
        if isinstance(Plotting_list,str):  # This is the case when there is only one datafile to plot.
            data_sheets = self.Extract(Plotting_list)[1]  # Extracting all the plottable data in this file.
            Plotting_list = [[Plotting_list, data_sheets]]  # Adjusting the format of this file to the input format of function Merge().

        data_dict = self.Merge(Plotting_list)

        # Visualization module
        curve_dict = {'Vgs-Id': ['GateV', 'DrainI'],
                      'Transfer Characteristics': ['GateV', 'DrainI'],
                      'Vds-Id': ['DrainV', 'DrainI'],
                      'Output Characteristics': ['DrainV', 'DrainI'],
                      'Resistive Characteristics': ['AV', 'AI'],
                      'R': ['AV', 'AI']}
        title_dict = {'Vgs-Id': [r'$\mathit{V}_gs$', r'$\mathit{I}_d'],
                      'Transfer Characteristics': ['Vgs', 'Id'],
                      'Vds-Id': [r'$\mathit{V}_{ds}$', r'$\mathit{I}_d$'],
                      'Output Characteristics': [r'$\mathit{V}_{ds}$', r'$\mathit{I}_d$'],
                      'Resistive Characteristics': ['V', 'I'],
                      'R': ['V', 'I']}

        unit_dict = {'A': 1, 'mA': 1e3, 'uA': 1e6, 'nA': 1e9,
                     'V': 1, 'mV': 1e3, 'uV': 1e6, 'nV': 1e9}
        xunit = kwargs['xunit'] if 'xunit' in kwargs else 'V'
        yunit = kwargs['yunit'] if 'yunit' in kwargs else 'A'

        num_curves = 0  # number of curves to plot
        for i in range(len(Plotting_list)):
            curve_list = Plotting_list[i][1]
            for j in curve_list:
                if 'label' in kwargs:
                    label = kwargs['label'][num_curves]
                else:
                    if j == 'Data':
                        label = 'loop 1'
                    else:
                        # index_set = set(n)-set('Append')  # 利用集合的不重复特性来提取测试序号
                        # index_list = list(index_set)
                        # index = int(''.join(index_list))
                        index = re.sub(r'\D', "", j)
                        label = 'loop ' + str(int(index) + 1)

                data = data_dict[Plotting_list[i][0]][j]  # Selecting the data of specific testing round from the specific file
                data_type = curve_dict[curve_type]  # Selecting the input data for the specific curve type

                # Adjusting the scale of the axes and the plotting data.
                xscale = kwargs['xscale'] if 'xscale' in kwargs else 'linear'
                yscale = kwargs['yscale'] if 'yscale' in kwargs else 'linear'
                if xscale != 'linear':
                    x = np.abs(data[data_type[0]])  # For log plotting, we need to set all data to be positive.
                    xunit = 'V'                      # Because the definition of loga(x) is that x belongs to (0,positive infinity)
                else:
                    x = [data[data_type[0]][n]*unit_dict[xunit] for n in range(len(data[data_type[0]]))]  # Adjusting the scale of the data.
                if yscale != 'linear':
                    y = np.abs(data[data_type[1]])
                    yunit = 'A'                                          # Setting the unit to SI units to avoid the influence of scales.
                else:
                    y = [data[data_type[1]][n]*unit_dict[yunit] for n in range(len(data[data_type[1]]))]

                # Setting some parameters of the plotting curves.
                linewidth = kwargs['linewidth'] if 'linewidth' in kwargs else 1.0  # Deciding the linewidth of the curve
                linestyle = kwargs['linestyle'] if 'linestyle' in kwargs else '-'  # The default value of linestyle is solid line
                pointsize = kwargs['pointsize'] if 'pointsize' in kwargs else 5.0  # Setting the size of the markers
                if 'color' in kwargs:
                    if isinstance(kwargs['color'],list):
                        color = kwargs['color'][num_curves]
                    else:
                        color = kwargs['color']
                else:
                    color = None  # Setting color value to None, and matplotlib will choose color on its own.
                if 'marker' in kwargs:
                    if isinstance(kwargs['marker'],list):
                        marker = kwargs['marker'][num_curves]
                    else:
                        marker = kwargs['marker']
                else:
                    marker = None

                # Setting ticks to point inward.
                plt.rcParams['xtick.direction'] = 'in'
                plt.rcParams['ytick.direction'] = 'in'

                if 'curve_mode' in kwargs:
                    if kwargs['curve_mode'] == 'scatter':
                        plt.scatter(x,y,s=pointsize,label=label,color=color,marker=marker)
                else:
                    plt.plot(x, y, label=label, linewidth=linewidth, linestyle=linestyle, color=color, marker=marker,markersize=pointsize)
                plt.xscale(xscale)
                plt.yscale(yscale)

                num_curves = num_curves+1

        plt.xlabel(title_dict[curve_type][0]+' ('+xunit+')',fontsize=16)
        yunit = r'$\mu$A' if yunit == 'uA' else yunit  # 利用matplotlib生成希腊字母mu
        plt.ylabel(title_dict[curve_type][1]+' ('+yunit+')',fontsize=16)
        if 'legend' in kwargs and kwargs['legend'] == None:
            pass  # 不做任何事情，只起到占位作用
        else:
            plt.legend(loc='best')
        if 'xlim' in kwargs:
            plt.xlim(kwargs['xlim'][0],kwargs['xlim'][1])
        if 'ylim' in kwargs:
            plt.ylim(kwargs['ylim'][0],kwargs['ylim'][1])
        if 'title' in kwargs:
            plt.title(kwargs['title'],fontsize=18)

        return

if __name__ == '__main__':
    file_name = 'D:/Projects/PhaseTransistor/DataGallery/2020-12-11/G3C90-Gminus200V-50V-Sweep-20.xls'
    geda = geda()
    # geda.Visualize(file_name,'Output Characteristics',['Append10','Append11','Append12'],['loop 10', 'loop 11', 'loop 12'],linestyle='-.',marker='^',yunit='uA',curve_mode='scatter',yscale='log',ylim=(1.e-12,1.e-6))
    # geda.Visualize(file_name,'Output Characteristics',['Append3'], ['Vg = -3 V'])

    directory = 'D:/Projects/PhaseTransistor/DataGallery/2020-12-11/'
    plotting_list = [[directory + 'G4C80-Gminus3V-7.5V-Sweep-04.xls', ['Append3']],
                     [directory + 'G4C80-Gminus4V-7.5V-Sweep-04.xls', ['Append3']],
                     [directory + 'G4C80-Gminus5V-7.5V-Sweep-04.xls', ['Append3']],
                     [directory + 'G4C80-Gminus6V-7.5V-Sweep-04.xls', ['Append3']],
                     [directory + 'G4C80-Gminus7V-7.5V-Sweep-04.xls', ['Append3']]]
    # geda.MakeFigure(plotting_list,yunit='uA')
    # geda.Visualize(directory + 'G4C80-Gminus3V-7.5V-Sweep-04.xls', 'Output Characteristics', ['Append3'], ['Vg = -3 V'])
    # a = geda.Merge(plotting_list)
    # print(a[directory + 'G4C80-Gminus3V-7.5V-Sweep-04.xls'])

    geda.Visualize(directory + 'G4C80-Gminus3V-7.5V-Sweep-04.xls','Output Characteristics',yunit='uA')




