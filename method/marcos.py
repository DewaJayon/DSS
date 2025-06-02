# Dibuat oleh Dewa Jayon

import pandas as pd
from tabulate import tabulate

class MARCOS:
    def __init__(self, file_path, weight, sheet_name ):
        self.file_path = file_path
        self.weight = weight
        self.sheet_name = sheet_name

    def marcos_method(file_path, weight, sheet_name):
        
        file_path = file_path

        # Membaca datasheet
        marcos_data = pd.read_excel(file_path, sheet_name, header=0, index_col=0)

        # Menghapus kolom dan baris yang tidak perlu
        marcos_data = marcos_data.dropna(how='all').dropna(axis=1, how='all')

        # print("Dataset:")
        # print(tabulate(marcos_data, headers=marcos_data.columns, tablefmt='grid'))
        # print("\n")

        # Mengambil bobot kriteria dari AHP
        weights = weight

        # print("Bobot kriteria:")
        # weight_df = pd.DataFrame({'Bobot': weights.round(4)})
        # print(tabulate(weight_df, headers='keys', tablefmt='pretty'))

        # Mengambil baris yang mengandung data kategori
        # Diasumsikan kategori berada di baris terakhir (indeks -1)
        categories = marcos_data.iloc[-1, 1:].values

        # print("Kategori untuk setiap kriteria:")
        # for i, kat in enumerate(categories, start=1):
        #     print(f"C{i}: {kat}")

        # Menghapus baris kategori sebelum dirubah menjadi float
        marcos_data = marcos_data.iloc[:-1]

        # Mengubah baris kategori menjadi float
        marcos_data.to_numpy(dtype=float)

        # print(tabulate(marcos_data, headers=marcos_data.columns, tablefmt='grid'))
        # print("\n")

        # Normalisasi Matriks
        for i, kat in enumerate(categories):
            if kat == "BENEFIT":
                marcos_data = marcos_data / marcos_data.max()
            else:
                marcos_data = marcos_data.min() / marcos_data

        # print("Normalisasi Matriks:")
        # print(tabulate(marcos_data.round(4), headers=marcos_data.columns, tablefmt='grid'))
        # print("\n")

        # Matriks ternormalisasi terbobot
        weight_norm = marcos_data * weights

        # print("Matriks ternormalisasi terbobot:")
        # print(tabulate(marcos_data.round(4), headers=marcos_data.columns, tablefmt='grid'))
        # print("\n")

        # Menghitung Tingkat Utilitas Alternatif
        # Sai
        sai = weight_norm.max(axis=0)
        sai_sum = sai.sum()
        
        # sai_df = pd.DataFrame({'SAI': sai.round(4)}, index=marcos_data.columns)
        # print("Tingkat Utilitas Alternatif SAI:")
        # print(tabulate(sai_df.round(4), headers='keys', tablefmt='grid'))
        # print(f"Total SAI: {sai_sum:.4f}")
        # print("\n")

        # Saai
        saai = weight_norm.min(axis=0)
        saai_sum = saai.sum()
        
        # saai_df = pd.DataFrame({'SAAI': saai.round(4)}, index=marcos_data.columns)
        # print("Tingkat Utilitas Alternatif SAAI:")
        # print(tabulate(saai_df.round(4), headers='keys', tablefmt='grid'))
        # print(f"Total SAAI: {saai_sum:.4f}")
        # print("\n")

        # Si
        si = weight_norm.sum(axis=1)
        
        # si_df = pd.DataFrame({'SI': si.round(4)}, index=marcos_data.index)
        # print("Tingkat Utilitas Kriteria SI:")
        # print(tabulate(si_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")


        # Menghitung Fungsi Utilitas Substitusi
        # KI-
        ki_minus = si / saai_sum
        
        # ki_df = pd.DataFrame({'KI-': ki_minus.round(4)}, index=marcos_data.index)
        # print("Fungsi Utilitas Substitusi KI-:")
        # print(tabulate(ki_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        # KI+
        ki_plus = si / sai_sum
        
        # ki_df = pd.DataFrame({'KI+': ki_plus.round(4)}, index=marcos_data.index)
        # print("Fungsi Utilitas Substitusi KI+:")
        # print(tabulate(ki_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        # Fki-
        fki_minus = ki_plus / (ki_plus + ki_minus)
        
        # fki_minus_df = pd.DataFrame({'FKI-': fki_minus.round(4)}, index=marcos_data.index)
        # print("Fungsi Utilitas Substitusi FKI-:")
        # print(tabulate(fki_minus_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")
        
        # Fki+
        fki_plus = ki_minus / (ki_plus + ki_minus)
        
        # fki_plus_df = pd.DataFrame({'FKI+': fki_plus.round(4)}, index=marcos_data.index)
        # print("Fungsi Utilitas Substitusi FKI+:")
        # print(tabulate(fki_plus_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        # Perankingan
        # pembilang
        numerator = ki_plus + ki_minus

        # print("Pembilang:")
        # numerator_df = pd.DataFrame({'Numerator': numerator.round(4)}, index=marcos_data.index)
        # print(tabulate(numerator_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        # penyebut
        # =1 + ((1-F83)/F83)+((1-E83)/E83)
        denominator = 1 + ((1 - fki_minus)/ fki_minus) + ((1 - fki_plus) / fki_plus)

        # denominator_df = pd.DataFrame({'Denominator': denominator.round(4)}, index=marcos_data.index)
        # print("Penyebut:")
        # print(tabulate(denominator_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        # Fki
        fki = numerator / denominator

        # fki_df = pd.DataFrame({'FKI': fki.round(4)}, index=marcos_data.index)
        # print("Fungsi Utilitas Substitusi FKI:")
        # print(tabulate(fki_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        # Ranking
        rank = fki.rank(ascending=False)
        
        # rank_df = pd.DataFrame({'Rank': rank.round(4)}, index=marcos_data.index)
        # print("Ranking:")
        # print(tabulate(rank_df.round(4), headers='keys', tablefmt='grid'))
        # print("\n")

        return fki, rank, marcos_data