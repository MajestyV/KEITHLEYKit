from KEITHLEY_wheels import Get_KEITHLEY_Data

gd = Get_KEITHLEY_Data.geda()

file = 'D:/Projects/PhaseTransistor/DataGallery/KEITHLEY-cmd/Phasetransistor/2021-03-22/Resistor/2mg-mL-C-ink-C100S30-8.xls'

directory = 'D:/Projects/PhaseTransistor/DataGallery/2021-01-31/'
plotting_list = [[file,['Data','Append1','Append2','Append3','Append4']]]
plotting_list_origin = [[file,['Append4']]]

gd.Visualize(plotting_list,'R',xlim=(-50,50),yunit='uA',yscale='linear',title='I-V characteristics of MoS2-P(VDF-TrFE) resistor')