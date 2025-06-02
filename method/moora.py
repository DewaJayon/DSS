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
        print("Dataset:")
        print(tabulate(df_moora, headers=df_moora.columns, tablefmt='grid'))
        print("\n")

        # Mengambil baris yang mengandung data kategori
        # Diasumsikan kategori berada di baris terakhir (indeks -1)
        categories = df_moora.iloc[-1].values
        
        print("Kategori untuk setiap kriteria:")
        for i, category in enumerate(categories, start=1):
            print(f"C{i}: {category}")

        # Menghapus baris kategori sebelum dirubah menjadi float
        df_moora = df_moora.iloc[:-1]
        
        # Mengubah baris kategori menjadi float
        df_numeric = df_moora.iloc[:-1].astype(float)

        print(tabulate(df_moora, headers=df_moora.columns, tablefmt='grid'))
        print("\n")


        # Menghitung pembagi untuk normalisasi
        squared = df_numeric.pow(2)
        sum_squared = squared.sum(axis=0)
        denominator = np.sqrt(sum_squared.values)  # Gunakan .values untuk array numpy
        
        denominator_df = pd.DataFrame([denominator], columns=df_numeric.columns, index=['Pembagi']).round(4)
        print("Pembagi untuk normalisasi:")
        print(tabulate(denominator_df, headers='keys', tablefmt='grid'))
        print("\n")

        # Normalisasi
        norm = df_numeric / denominator

        norm_df = pd.DataFrame(norm, columns=df_numeric.columns, index=df_numeric.index).round(4)
        print("Normalisasi:")
        print(tabulate(norm_df, headers='keys', tablefmt='grid'))
        print("\n")

        # Optimasi Nilai Atribut
        optimized = norm * weight

        optimized_df = pd.DataFrame(optimized, columns=df_numeric.columns, index=df_numeric.index).round(4)
        print("Optimasi Nilai Atribut:")
        print(tabulate(optimized_df, headers='keys', tablefmt='grid'))
        print("\n")


        # Menentukan maxsimum
        for i, kat in enumerate(categories):
            col = optimized.columns[i]
            if kat == "BENEFIT":
                optimized[i] = optimized[col].max()
            elif kat == "COST":
                optimized[i] = optimized[col].min()
            else:
                optimized[i] = 0

        result_df = pd.DataFrame({
            'Max (Benefit)': optimized,

        }).round(4)

        print("Hasil:")
        print(tabulate(result_df, headers='keys', tablefmt='grid', showindex=False))
        print("\n")

        return result_df
