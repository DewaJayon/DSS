# Dibuat oleh Dewa Jayon

import pandas as pd
from tabulate import tabulate

class AHP:
    def __init__(self, file_path):
        self.file_path = file_path

    def ahp_method(file_path, sheet_name):

        f_path = file_path

        # Membaca dataset dari sheet AHP
        df_ahp = pd.read_excel(f_path, sheet_name, header=0, index_col=0)

        # Menghapus kolom dan baris yang tidak perlu (jika ada)
        df_ahp = df_ahp.dropna(how='all').dropna(axis=1, how='all')

        # Mengkonversi ke numpy array
        dataset = df_ahp.to_numpy(dtype=float)

        # print("Dataset:")
        # print(tabulate(dataset, headers=df_ahp.columns, tablefmt='grid'))
        # print("\n")

        # Normalisasi matriks
        column_sums = dataset.sum(axis=0)
        normalized_matrix = dataset / column_sums
        
        # print("Normalisasi Matriks:")
        # print(tabulate(normalized_matrix.round(4), headers=df_ahp.columns, tablefmt='grid'))
        # print("\n")

        # Perhitungan P.Vector
        p_vector = normalized_matrix.sum(axis=1)

        # p_vector_df = pd.DataFrame({'P.Vector': p_vector.round(4)}, index=df_ahp.index)
        # print("\nP.Vector:")
        # print(tabulate(p_vector_df, headers='keys', tablefmt='pretty'))

        # perhitungan bobot
        criteria_count = df_ahp.shape[0]
        weight = p_vector / criteria_count

        # weight_df = pd.DataFrame({'Bobot': weight.round(4)}, index=df_ahp.index)
        # print("\nBobot:")
        # print(tabulate(weight_df, headers='keys', tablefmt='pretty'))
        # print(weight.sum())

        # menghitung Eigen Value
        dataset_sum = dataset.sum(axis=0)
        eigen_value = dataset_sum * weight
        
        # eigen_value_df = pd.DataFrame({'Eigen Value': eigen_value.round(4)}, index=df_ahp.index)
        # print("\nEigen Value:")
        # print(tabulate(eigen_value_df, headers='keys', tablefmt='pretty'))

        # menghitung konsistensi
        # Menghitung CI
        sum_eigen_value = eigen_value.sum()
        CI = (sum_eigen_value - criteria_count) / (criteria_count - 1)
        
        # Indeks Random (RI)
        RI_dict = {1: 0.0, 2: 0.0, 3: 0.58, 4: 0.9, 5: 1.12,
                   6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}
        RI = RI_dict[criteria_count]

        # print(RI)

        # Menghitung CR
        CR = CI / RI

        # print("\nConsistency Ratio (CR): " + str(CR.round(4)))
        # if CR < 0.1:
        #     print("Konsisten")
        # else:
        #     print("Tidak konsisten")

        return weight, CR, df_ahp