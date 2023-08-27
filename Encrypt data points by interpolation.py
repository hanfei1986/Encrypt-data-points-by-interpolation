import glob
import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from scipy.interpolate import interp1d

path = 'C:/Users/fhan/Desktop/Encrypt data points by interpolation\*.xlsx'
for file in glob.glob(path):
    filename = file.split('\\')[1].split('.')[0]
    df = pd.read_excel(file)
    x = df['x']
    y = df['y']
    
    new_x = np.arange(x.min(), x.max()+1, 1)

    print(new_x)
    
    f_xy = interp1d(x, y, kind='cubic')

    new_y = f_xy(new_x)

    plt.scatter(x, y)
    plt.plot(new_x, new_y)
    plt.xlabel('x')
    plt.ylabel('y')
    plt.show()

    new_df = pd.DataFrame({'x': new_x, 'y': new_y})
    new_df.to_excel('C:/Users/fhan/Desktop/Encrypt data points by interpolation/{} interpolated.xlsx'.format(filename), index=False)