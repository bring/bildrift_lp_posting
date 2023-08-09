#%%
import pandas as pd
from pathlib import Path
df = pd.read_excel(Path(__file__).parents[0]/"Bilfåte - Rapport.xlsx")
df = df[["Kostsenternummer", "Kostsenter, beskrivelse"]]
df = df.drop_duplicates(subset = "Kostsenternummer")
df = df.reset_index(drop = True)
df.to_excel(Path(__file__).parents[1] / "mapping/mapping_lp.xlsx", index = False)

# %%
