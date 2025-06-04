import pandas as pd
from tabulate import tabulate
import numpy as np

class TOPSIS:
    def __init__(self, datasheet, sheet_name, weight):
        self.datasheet = datasheet
        self.sheet_name = sheet_name
        self.weight = weight

    def topsis_method(datasheet, sheet_name, weight):
        
        # Membaca dataset dari excel
        df_topsis = pd.read_excel(datasheet, sheet_name, header=0, index_col=0)
        
        # print("Dataset TOPSIS:")
        # print(tabulate(df_topsis, headers=df_topsis.columns, tablefmt='grid'))
        # print("\n")

        # Mengambil baris yang mengandung data kategori
        categories = df_topsis.iloc[-1].values
        
        # print("Kategori untuk setiap kriteria:")
        # for i, category in enumerate(categories, start=1):
        #     print(f"C{i}: {category}")

        # Menghapus baris kategori sebelum dirubah menjadi float
        df_topsis = df_topsis.iloc[:-1]

        # Mengubah baris kategori menjadi float
        df_topsis = df_topsis.astype(float)

        # print("\nDataset TOPSIS:")
        # print(tabulate(df_topsis, headers=df_topsis.columns, tablefmt='grid'))
        # print("\n")

        # Mencari pembagi untuk normalisasi matriks
        pembagi = np.sqrt((df_topsis ** 2).sum(axis=0))

        # print("Pembagi per kriteria:")
        # for i, val in enumerate(pembagi, start=1):
        #     pembagi_df = pd.DataFrame({'Pembagi': pembagi}, index=df_topsis.columns)

        # print(tabulate(pembagi_df.round(4), headers='keys', tablefmt='pretty'))
        # print("\n")

        # Matriks ternormalisasi R
        normalisasi = df_topsis / pembagi

        # print("Matriks ternormalisasi R:")
        # print(tabulate(normalisasi.round(4), headers=normalisasi.columns, tablefmt='grid'))
        # print("\n")

        # Matriks ternormalisasi terbobot
        normalisasi_terbobot = normalisasi * weight

        # print("Matriks ternormalisasi terbobot:")
        # print(tabulate(normalisasi_terbobot.round(4), headers=normalisasi_terbobot.columns, tablefmt='grid'))
        # print("\n")

        # Solusi ideal positif A+
        solusi_ideal_positif = []
        for i, kat in enumerate(categories):
            if kat == 'BENEFIT':
                solusi_positif = normalisasi_terbobot.max(axis=0).iloc[i]
            else:
                solusi_positif = normalisasi_terbobot.min(axis=0).iloc[i]
            solusi_ideal_positif.append(solusi_positif)

        df_solusi_ideal_positif = pd.DataFrame({'Solusi Ideal Positif': solusi_ideal_positif}, index=df_topsis.columns)

        # print("Solusi Ideal Positif:")
        # print(tabulate(df_solusi_ideal_positif.round(4), headers='keys', tablefmt='pretty'))
        # print("\n")

        # Solusi ideal negatif A-
        solusi_ideal_negatif = []
        for i, kat in enumerate(categories):
            if kat == 'BENEFIT':
                solusi_negatif = normalisasi_terbobot.min(axis=0).iloc[i]
            else:
                solusi_negatif = normalisasi_terbobot.max(axis=0).iloc[i]
            solusi_ideal_negatif.append(solusi_negatif)

        # print("Solusi Ideal Negatif:")
        # df_solusi_ideal_negatif = pd.DataFrame({'Solusi Ideal Negatif': solusi_ideal_negatif}, index=df_topsis.columns)
        # print(tabulate(df_solusi_ideal_negatif.round(4), headers='keys', tablefmt='pretty'))
        # print("\n")

        # Konversi solusi ideal positif menjadi array
        solusi_ideal_positif_array = np.array(solusi_ideal_positif)

        # Menghitung jarak solusi ideal positif
        jarak_positif = np.sqrt(((normalisasi_terbobot - solusi_ideal_positif_array) ** 2).sum(axis=1))

        # print("\nJarak Solusi Ideal Positif:")
        # jarak_positif_series = pd.Series(jarak_positif, index=normalisasi_terbobot.index, name='Jarak Positif')
        # print(tabulate(jarak_positif_series.round(4).to_frame(), headers='keys', tablefmt='grid'))
        # print("\n")

        # Konversi solusi ideal negatif menjadi array
        solusi_ideal_negatif_array = np.array(solusi_ideal_negatif)

        # Menghitung jarak solusi ideal negatif
        jarak_negatif = np.sqrt(((normalisasi_terbobot - solusi_ideal_negatif_array) ** 2).sum(axis=1))

        # print("Jarak Solusi Ideal Negatif:")
        # jarak_negatif_series = pd.Series(jarak_negatif, index=normalisasi_terbobot.index, name='Jarak Negatif')
        # print(tabulate(jarak_negatif_series.round(4).to_frame(), headers='keys', tablefmt='grid'))
        # print("\n")

        # Nilai Refrensi
        nilai_refrensi = jarak_negatif / (jarak_negatif + jarak_positif)

        # print("Nilai Refrensi:")
        # print(tabulate(nilai_refrensi.round(4).to_frame(), headers='keys', tablefmt='grid'))
        # print("\n")

        # Ranking
        rank = nilai_refrensi.rank(ascending=False)

        # print("Ranking:")
        # print(tabulate(rank.round(4).to_frame(), headers='keys', tablefmt='grid'))
        # print("\n")

        return nilai_refrensi, rank, df_topsis
