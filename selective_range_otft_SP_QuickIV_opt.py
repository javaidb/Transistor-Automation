#from multiprocessing import Pool
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
from os import listdir
import os
import PIL
from PIL import Image
import csv
import matplotlib.ticker as mtick
import sys
import time
from functools import partial
from math import log10, floor
import warnings

def round_to_1(x):
    return (10**(round(log10(abs(x)))))

def do_plot(epsilon, Ci, color_code, polarity, name):
    warnings.filterwarnings('ignore', category=UserWarning, append=True)
    if len(sys.argv)>1:
       WL = float(sys.argv[1])
    else:
       WL = 33

    if len(sys.argv)>1:
        transfer= abs(float(sys.argv[2]))
    else:
        transfer=100
    
    if 'O' in name:
        if 'PO' in name:
            df = pd.read_csv(name, encoding='utf-8', header=0, usecols=[3,4,9,10], names = ['V_Drain', 'I_Drain', 'V_Gate', 'I_Gate'])
        if 'NO' in name:
            df = pd.read_csv(name, encoding='utf-8', header=0, usecols=[3,4,10,11], names = ['V_Drain', 'I_Drain', 'V_Gate', 'I_Gate'])
        df['V_Gate'].iloc[0]=0
        plt.rc('xtick', labelsize=20) 
        plt.rc('ytick', labelsize=20)
        plt.rc('axes', linewidth=1.5)
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        ax1 = plt.gca()
        ax1.tick_params(width=1.5)
        ax1.yaxis.set_tick_params(which='minor', width=1.5)
        ax1.set_xlabel('Drain Voltage (V)', fontsize=26)
        ax1.set_ylabel('Drain Current ($\mu$A)', fontsize=26)
        plt.xlim([ 0, df['V_Drain'].iloc[-1]])
        if 'PO' in name:
            plt.gca().invert_yaxis()
        for i in range (int(len(df.index)/51)):
            x = i * 51
            y = 51 * (i + 1)
            ax1.plot(df['V_Drain'].iloc[x:y],df['I_Drain'].iloc[x:y]/epsilon, color=color_code[i], linewidth=1.5, label=str(int(df['V_Gate'].iloc[x])) +" V")
            #print(x)
            #print(y)

        legend = ax1.legend(loc="best", fontsize=11, title="V$_G$$_S$")
        legend.get_title().set_fontsize('11')
        plt.tight_layout()
        figname = '%s.png' % name
        plt.savefig('processed_data/' + figname, format='png', dpi=600)
        basewidth = 1080
        img = Image.open('processed_data/' + figname)
        wpercent = (basewidth / float(img.size[0]))
        hsize = int((float(img.size[1]) * float(wpercent)))
        img = img.resize((basewidth , hsize), PIL.Image.ANTIALIAS)
        img.save('processed_data/' + figname, dpi=(600.0, 600.0))
        df.to_csv('processed_data/' + name + '_processed' + '.csv')
    else:
        df = pd.read_csv(name, encoding='utf-8', header=0, usecols=[3,4,10,11], names = ['V_Drain', 'I_Drain', 'V_Gate', 'I_Gate'])
        df['I_Drain_abs'] = df['I_Drain'].abs()
        df['I_Drain_sqrt'] = df['I_Drain_abs'].apply(np.sqrt)
        if 'PT' in name:
            transfer=-transfer
        df_specified = df.loc[df['V_Drain'] == transfer]
        plt.plot(df_specified['V_Gate'],df_specified['I_Drain_sqrt'], 'k-', linewidth=1.5, label=str(transfer)+" V")
        pts = plt.ginput(2) # it will wait for two clicks
        x_val = [int(x[0]) for x in pts]
        plt.close("all")
        if 'NT' in name:
            first=(df_specified[df_specified['V_Gate'] >= min(x_val)].index.tolist()[0])
            last=(df_specified[df_specified['V_Gate'] <= max(x_val)].index.tolist()[-1])
        else:
            last=(df_specified[df_specified['V_Gate'] >= min(x_val)].index.tolist()[-1])
            first=(df_specified[df_specified['V_Gate'] <= max(x_val)].index.tolist()[0])
        start=first-(floor(first/61))*61
        finish=last-(floor(last/61))*61
        for i in np.arange(start,finish):
            x = df_specified['V_Gate'].iloc[i:finish] 
            y = df_specified['I_Drain_sqrt'].iloc[i:finish]
            slope, intercept, r_value, p_value, std_err = stats.linregress(x,y)
            r_sq = 0.99
            if r_value**2 >= r_sq: 
                break
        predict_y = intercept + slope * x
        pred_error = y - predict_y
        degrees_of_freedom = len(x) - 2
        try:
            residual_std_error = np.sqrt(np.sum(pred_error**2) / degrees_of_freedom)#HUGE ISSUE, deg of error is 0 if only 2 points and no need of this line
        except ZeroDivisionError as err:
            print("There are only two points for linear reg. file: " + name)

        mob = 2 * slope ** 2 / (WL * Ci)
        Vth = -(intercept / slope)
        plt.rc('xtick', labelsize=20) 
        plt.rc('ytick', labelsize=20)
        plt.rc('axes', linewidth=1.5)
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        ax1 = plt.gca()
        ax1.tick_params(width=1.5)
        ax1.yaxis.set_tick_params(which='minor', width=1.5)
        ax1.set_yscale('log')
        ax1.plot(df_specified['V_Gate'],df_specified['I_Drain_abs'], 'k-', linewidth=1.5, label=str(transfer)+" V")
        ax1.set_xlabel('Gate Voltage (V)', fontsize=26)
        ax1.set_ylabel('Drain Current (A)', fontsize=26)
        ax2 = ax1.twinx()
        ax2 = plt.gca()
        ax2.tick_params(width=1.5)
        ax2.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.1e'))
        ax2.plot(df_specified['V_Gate'],df_specified['I_Drain_sqrt'],'r-', linewidth=1.5)
        ax2.set_ylabel('Drain Current$^{1/2}$ (A$^{1/2}$)', color='r', fontsize=26)
        for tl in ax2.get_yticklabels():
            tl.set_color('r')
        ax2.plot(x,predict_y,'b-', linewidth=1.5)
        ax1.set_title('Mobility: %.5f Vth: %.2f R$^2$: %.2f' % (mob,Vth,r_value**2), fontsize=18)
        legend = ax1.legend(loc="best", fontsize=11, title="V$_D$$_S$")
        legend.get_title().set_fontsize('11')
        plt.tight_layout()
        figname = '%s.png' % name
        plt.savefig('processed_data/' + figname, format='png', dpi=600)
        basewidth = 1080
        img = Image.open('processed_data/' + figname)
        wpercent = (basewidth / float(img.size[0]))
        hsize = int((float(img.size[1]) * float(wpercent)))
        img = img.resize((basewidth , hsize), PIL.Image.ANTIALIAS)
        img.save('processed_data/' + figname, dpi=(600.0, 600.0))
        df['x_lin'] = x
        df['y_lin'] = predict_y
        if 'PT' in name:
            polarity="P-type"
            I_on_off=round_to_1((df_specified['I_Drain_abs'].iloc[-1])/df_specified['I_Drain_abs'].min())
        else:
            I_on_off=round_to_1((df_specified['I_Drain_abs'].iloc[-1])/df_specified['I_Drain_abs'].min())

        parameters = [[polarity, mob, Vth, r_value**2, '%E' %I_on_off]]
        new_df = pd.DataFrame(parameters)
        new_df.to_csv('processed_data/parameters.csv', mode = 'a', index=False, header=False)
        df.to_csv('processed_data/' + name + '_processed' + '.csv')
    plt.close()


def find_csv_filenames( path_to_dir, suffix=".csv" ):
   filenames = listdir(path_to_dir)
   return [ filename for filename in filenames if filename.endswith( suffix ) or filename.endswith (suffix.upper()) ]

if __name__ == '__main__':    
    #Constants
    epsilon = 10**(-6.0)
    Ci = 1.16E-08
    color_code=['k','b','r','g','c','m', 'y', 'lime', 'crimson', 'teal', 'aqua']
    polarity="N-type"

    #pool = Pool(processes=1)
    try: 
        os.makedirs('processed_data')
    except OSError:
        if not os.path.isdir('processed_data'):
            raise
    parameters = [[" ", "Mobility", "Vth", "Coeff. Det.", "I on/off"]]
    new_df = pd.DataFrame(parameters)
    new_df.to_csv('processed_data/parameters.csv', mode = 'w', index=False, header=False)
    filenames = find_csv_filenames(".")
    #func=partial(do_plot, epsilon, Ci, color_code, polarity)
    for x in range(0,len(filenames)):
        do_plot(epsilon, Ci, color_code, polarity, filenames[x])
    #pool.map(func, filenames)
    #pool.close()
    #pool.join()
