from method.ahp import AHP as ahp
from method.marcos import MARCOS as marcos
from method.moora import MOORA as moora
from method.topsis import TOPSIS as topsis
import pandas as pd
from tabulate import tabulate
from utils.copyright import copyright

copyright()

weight, cr, df_ahp = ahp.ahp_method(file_path="datasheet.xlsx", sheet_name="AHP")

weight_df = pd.DataFrame({'Bobot': weight.round(4)}, index=df_ahp.index)

print("\n========Metode AHP========")
print("\nBobot:")
print(tabulate(weight_df, headers='keys', tablefmt='pretty'))
print(weight.sum())

print(f"\nConsistency Ratio (CR): {cr:.4f}")
if cr < 0.1:

    print("\nKonsisten, melanjutkan ke metode MARCOS...")

    fki, rank, marcos_data  = marcos.marcos_method(file_path="datasheet.xlsx", sheet_name="MARCOS", weight=weight)

    print("\n========Metode MARCOS========")
    print("Ranking:")
    rank_df = pd.DataFrame({'Fki': fki.round(4), 'Rank': rank.round(4)}, index=marcos_data.index)
    print(tabulate(rank_df.round(4), headers='keys', tablefmt='grid'))
    print("\n")

    print("\n========Metode MOORA========")
    yi, rank = moora.moora_method(datasheet="datasheet.xlsx", sheet_name="MOORA", weight=weight)

    print("Ranking:")
    rank_df = pd.DataFrame({'Yi': yi.round(4), 'Rank': rank.round(4)}, index=marcos_data.index)
    print(tabulate(rank_df.round(4), headers='keys', tablefmt='grid'))
    print("\n")

    print("\n========Metode TOPSIS========")
    nilai_refrensi, rank, df_topsis = topsis.topsis_method(datasheet="datasheet.xlsx", sheet_name="TOPSIS", weight=weight)

    print("Ranking:")
    rank_df = pd.DataFrame({'Nilai Refrensi': nilai_refrensi.round(4), 'Rank': rank.round(4)}, index=df_topsis.index)
    print(tabulate(rank_df.round(4), headers='keys', tablefmt='grid'))
    print("\n")

else:
    print("Tidak konsisten")