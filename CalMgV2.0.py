'''
    Author: Zijie Gao
    Date: 11/29/2020
    Function: Calculating the Molar content of MgCO3 from XRD data
    Version : 2.0
    New function:
    1) write results into a txt document
    2) get the filename of the reading data, and output it together with the corresponded results
'''

import pandas as pd
import xlrd
import math
from math import pi
import os

def get_filename(path,filetype):
    '''
        Read file name
    '''
    name = []
    for root,dirs,files in os.walk(path):
        for i in files:
            if filetype in i:
                name.append(i.replace(filetype,''))
    return name

def write_txt(list):
    '''
        write results into .txt document
    '''
    fileObject = open('MgCO3.txt', 'a')

    for ip in list:
        fileObject.writelines(str(ip))
        fileObject.write(' ')
    fileObject.write('\n')

    fileObject.close()

def main():
    path = 'C:\\Users\CUPGZ\PycharmProjects\CalMg_new'
    filetype = '.xlsx'
    name = get_filename(path, filetype)
    for word in name:
        data = xlrd.open_workbook(word + '.xlsx')
        table = data.sheets()[0]
        twotheta = table.col_values(0)
        intensity = table.col_values(1)

        # Combine two lists into one tuple
        data_tuples = list(zip(twotheta, intensity))

        # Converting list of tuples to pandas dataframe
        data_pd = pd.DataFrame(data_tuples, columns=['twotheta', 'intensity'])



        # Use two conditions to filter data and get the index
        # print(data_pd[(data_pd['twotheta']>=29.7) & (data_pd['twotheta']<=29.8)])
        index_list = data_pd[(data_pd['twotheta']>=29.7) & (data_pd['twotheta']<=29.8)].index.tolist()


        # Sum of intensity
        sum_intensity = data_pd[index_list[0]:max(index_list)+1]['intensity'].sum()
        # print(sum_intensity)
        # Sum of products
        sum_product = 0

        for i in index_list:
            product = data_pd.at[i, 'twotheta'] * data_pd.at[i, 'intensity']
            sum_product = sum_product + product

        # print(sum_product)
        division = sum_product/sum_intensity
        # print(division)

        # Calculate x
        x = 1.54178/(2*math.sin((division/2)/180*pi))
        # print(x)
        list_result = []
        list_result.append(word)
        # Arvidson & Mackenzie 1999
        mgco3_am = -363.96*x + 1104.5
        # print(mgco3_am)
        list_result.append(mgco3_am)

        # Lumsden and Chimahusky 1980
        mgco3_lc = 100-(333.33*x-911.99)
        list_result.append(mgco3_lc)
        write_txt(list_result)
        print(list_result)






if __name__ == '__main__':
        main()