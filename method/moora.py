# Dibuat oleh Dewa Jayon

import pandas as pd
from tabulate import tabulate
import numpy as np

class MOORA:
    def __init__(self, datasheet, sheet_name, weight):
        self.datasheet = datasheet
        self.sheet_name = sheet_name
        self.weight = weight

    def moora_method(datasheet, sheet_name, weight):

        df_moora = pd.read_excel(datasheet, sheet_name, header=0, index_col=0)
        # print("Dataset:")
        # print(tabulate(df_moora, headers=df_moora.columns, tablefmt='grid'))
        # print("\n")

        # Mengambil baris yang mengandung data kategori
        # Diasumsikan kategori berada di baris terakhir (indeks -1)
        categories = df_moora.iloc[-1].values
        
        # print("Kategori untuk setiap kriteria:")
        # for i, category in enumerate(categories, start=1):
        #     print(f"C{i}: {category}")

        # Menghapus baris kategori sebelum dirubah menjadi float
        df_moora = df_moora.iloc[:-1]
        
        # Mengubah baris kategori menjadi float
        df_moora = df_moora.astype(float)

        # print(tabulate(df_moora, headers=df_moora.columns, tablefmt='grid'))
        # print("\n") 

        # Mencari pembagi dengan rumus sqrt
        pembagi = np.sqrt((df_moora ** 2).sum(axis=0))
        
        # print("Pembagi per kriteria:")
        # for i, val in enumerate(pembagi, start=1):
        #     pembagi_df = pd.DataFrame({'Pembagi': pembagi}, index=df_moora.columns)

        # print(tabulate(pembagi_df.round(), headers='keys', tablefmt='pretty'))
        # print("\n")

        # Normalisasi
        df_moora_norm = df_moora / pembagi
        
        # print("Normalisasi:")
        # print(tabulate(df_moora_norm.round(4), headers=df_moora.columns, tablefmt='grid'))
        # print("\n")

        # Optimasi nilai atribut
        df_moora_opt = df_moora_norm * weight
       
        # print("Optimasi Nilai Atribut:")
        # print(tabulate(df_moora_opt.round(4), headers=df_moora.columns, tablefmt='grid'))
        # print("\n")

        # Hitung Nilai Maksimal
        max_list = []
        for row_idx in range(len(df_moora_opt)):
            benefit = 0
            cost = 0
            for i, kat in enumerate(categories):
                val = df_moora_opt.iloc[row_idx, i]
                if kat.upper() == 'BENEFIT':
                    benefit += val
                else:
                    cost
            skor = benefit - cost
            max_list.append(round(skor, 4))
        
        # df_maximun = pd.DataFrame({'Maximun': max_list}, index=df_moora.index)
        # print(tabulate(df_maximun, headers='keys', tablefmt='pretty'))

        # Hitung Nilai minimum
        min_list = []
        for row_idx in range(len(df_moora_opt)):
            benefit = 0
            cost = 0
            for i, kat in enumerate(categories):
                val = df_moora_opt.iloc[row_idx, i]
                if kat.upper() == 'BENEFIT':
                    benefit
                else:
                    cost += val
            skor = cost - benefit
            min_list.append(round(skor, 4))
        
        # df_minimun = pd.DataFrame({'Minimun': min_list}, index=df_moora.index)
        # print(tabulate(df_minimun, headers='keys', tablefmt='pretty'))

        # Menghitung nilai YI
        yi = np.array(max_list) - np.array(min_list)
        
        # df_yi = pd.DataFrame({'Yi': yi.round(4)}, index=df_moora.index)
        # print(tabulate(df_yi, headers='keys', tablefmt='pretty'))

        # Ranking
        rank = yi.argsort()[::-1].argsort() + 1
        
        # df_rank = pd.DataFrame({
        #     'Maximum': max_list,
        #     'Minimum': min_list,
        #     'Yi': yi.round(4), 
        #     'Rank': rank,
        #     }, index=df_moora.index)
        # print(tabulate(df_rank, headers='keys', tablefmt='pretty'))

        return yi, rank