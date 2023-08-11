#%%
import pandas as pd
from pathlib import Path
from kontering_NN import unique_listdir

def create_mapping():
    # df = pd.read_excel(Path(__file__).parents[0]/"Bilfåte - Rapport.xlsx")

    path = Path(__file__).parents[1] / "data"
    direc = unique_listdir(path)
    direc.sort()
    df = pd.read_excel(path / direc[-1], sheet_name = "Konteringsbilag")
    df = df.rename(columns = {
        "Kostsenter beskrivelse 3":"Kostsenter, beskrivelse",
        "kopcno":"Kostsenternummer"
    })

    df = df[["Kostsenternummer", "Kostsenter, beskrivelse"]]
    df = df.drop_duplicates(subset = "Kostsenternummer")
    df = df.reset_index(drop = True)
    df.to_excel(Path(__file__).parents[1] / "mapping/mapping_lp.xlsx", index = False)

# %%
