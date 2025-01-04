import pandas as pd
import os
from openpyxl import load_workbook
from typing import List

# global paramaters
FILE_PATH = os.path.join('.', 'data', 'Others-HF - 1.xlsx')
SHEET_NAME = 'Raw data'
METHOD = ['FA', 'BA']
WEIGHT = [[0.257, 0.2559, 0.2593, 0.2504, 0.2532], 
          [0.2536, 0.2548, 0.2541, 0.2522, 0.2574]]
BATCH = 5
DF = 50
ELEMENT = ['Al', 'As', 'Ca', 'Cd', 'Cr', 'Cu', 'Fe', 'K', 'Mg', 'Mn', 'Na', 'Ni', 'Pb', 'Si', 'Ti', 'Zn']


def load_excel_sheet(name:str) -> List[List[pd.DataFrame]]:
    # local parameters
    columns = ['Mean']
    n = len(ELEMENT)
    data = [] # [[DataFrames for FA], [DataFrames for BA]]
    # load sheet data into DataFrame
    sheet = pd.read_excel(io=FILE_PATH, sheet_name=name, header=0, usecols=columns)
    print(f'load sheet "{name}" from {FILE_PATH}')
    # delete empty rows and reset index
    sheet = sheet.dropna(axis=0, how='all')
    sheet['Mean'] = pd.to_numeric(sheet['Mean'], 'coerce')
    sheet = sheet.dropna(axis=0, how='any')
    sheet = sheet.reset_index(drop=True)
    # add each DataFrame to a list
    for i in range(len(METHOD)):
        dfs = []
        for j in range(BATCH):
            df = sheet.loc[(i*BATCH*n + j*n):(i*BATCH*n + (j+1)*n - 1), :]
            df.index = ELEMENT
            dfs.append(df)
        data.append(dfs)
    return data


def export_excel_sheet(df:pd.DataFrame, name:str) -> None:
    book = load_workbook(FILE_PATH)
    with pd.ExcelWriter(FILE_PATH, engine='openpyxl') as writer:
        writer.book = book
        df.to_excel(writer, sheet_name=name)
    print(f'add sheet "{name}" to {FILE_PATH}')


def process(sheet_name:str) -> None:
    # get raw data
    raw_data = load_excel_sheet(sheet_name)
    # calculate for each method
    for i in range(len(METHOD)):
        method = METHOD[i]
        data = raw_data[i]
        weight = WEIGHT[i]
        result = []
        # calculate for each batch
        for j in range(BATCH):
            columns = [f'{method}{j+1} Conc.(mg/L)', f'{method}{j+1} Conc before diliton (mg/L)', 
                        f'{method}{j+1} Conc.(mg/g)']
            df = data[j]
            # formulas
            df['original'] = df['Mean'] * DF
            df['res'] = df['original']/1000/weight[j]
            # rename columns
            df.columns = columns
            result.append(df)
        # merge all data for one method
        output = pd.concat(result, axis=1)
        export_excel_sheet(output, method+'_verify')
    print('process finished')

def hello(name:str):
    print(f'hello {name}')


if __name__ == "__main__":
    # process(SHEET_NAME)
    hello('123')