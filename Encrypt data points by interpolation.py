import glob
import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from scipy.interpolate import interp1d

path = 'C:/Users/fhan/Desktop/Encrypt data points by interpolation\*.xlsx'
for file in glob.glob(path):
    filename = file.split('\\')[1].split('.')[0]
    df = pd.read_excel(file)
    new_df = pd.DataFrame()

    if len(df.columns)>=1:
        x = df[df.columns[0]]
        new_x = np.arange(x.min(), x.max()+1, 1)
        new_df[df.columns[0]] = new_x
        for i in range(1,len(df.columns)):
            y = df[df.columns[i]]
            f_xy = interp1d(x, y, kind='cubic')
            new_y = f_xy(new_x)
            new_df[df.columns[i]] = new_y

            plt.plot(x, y, 'o-', c='blue', markersize=10)
            plt.plot(new_x, new_y, marker='.', c='red', markersize=5)
            plt.xlabel(df.columns[0])
            plt.ylabel(df.columns[i])
            plt.show()

        new_df.to_excel('C:/Users/fhan/Desktop/Encrypt data points by interpolation/{} interpolated.xlsx'.format(filename), index=False)