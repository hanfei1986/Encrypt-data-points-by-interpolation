import glob
import csv
import numpy as np
from matplotlib import pyplot as plt
from scipy.interpolate import interp1d
import xlwt

path = 'C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Raw data/Production wafer measurement\*.csv'
for filename in glob.glob(path):
    wafer_ID = str.upper(filename.split('\\')[1][0:6])+'-'+filename.split('\\')[1][7:9]
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        reader_rows = [row for row in reader]
    i = 0
    while True:
        if reader_rows[i][0] == '700':
            row_700 = i
        elif reader_rows[i][0] == '350':
            row_350 = i
            break
        i +=1        
    wavelength = [float(reader_rows[i][0]) for i in range(row_700,row_350+1)]
    baseline = [float(reader_rows[i][1]) for i in range(row_700,row_350+1)]
    reflectance = [float(reader_rows[i][3]) for i in range(row_700,row_350+1)]
    transmittance = [float(reader_rows[i][5]) for i in range(row_700,row_350+1)]
    
    plt.plot(wavelength,baseline,'o')
    plt.show()
    plt.plot(wavelength,reflectance,'o')
    plt.show()
    plt.plot(wavelength,transmittance,'o')
    plt.show()
    
    new_wavelength = list(np.arange(700, 349, -1))
    
    f_baseline = interp1d(wavelength, baseline, kind='linear')  
    f_reflectance = interp1d(wavelength, reflectance, kind='linear')
    f_transmittance = interp1d(wavelength, transmittance, kind='linear')

    new_baseline = f_baseline(new_wavelength)    
    new_reflectance = f_reflectance(new_wavelength)
    new_transmittance = f_transmittance(new_wavelength)
    
    plt.plot(new_wavelength,new_baseline,'o')
    plt.show()
    plt.plot(new_wavelength,new_reflectance,'o')
    plt.show()
    plt.plot(new_wavelength,new_transmittance,'o')
    plt.show()
    
    workbook_newdata = xlwt.Workbook(encoding='utf-8')
    sheet1 = workbook_newdata.add_sheet(wafer_ID)
    for sheet1_row in range(0,len(new_wavelength)):
        sheet1.write(sheet1_row,0,int(new_wavelength[sheet1_row]))
        sheet1.write(sheet1_row,1,new_baseline[sheet1_row])
        sheet1.write(sheet1_row,2,int(new_wavelength[sheet1_row]))
        sheet1.write(sheet1_row,3,new_reflectance[sheet1_row])
        sheet1.write(sheet1_row,4,int(new_wavelength[sheet1_row]))
        sheet1.write(sheet1_row,5,new_transmittance[sheet1_row])
    workbook_newdata.save(r'C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Summary/'+wafer_ID+'.xls')
    