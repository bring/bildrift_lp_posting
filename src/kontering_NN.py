# %%
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 14 09:50:40 2023

@author: granenga
"""

import pandas as pd
import numpy as np
import datetime as dt
import os
from pathlib import Path
import create_mapping

cuno_pbb = 99389  # Customer nr of PBB in Leaseplan


def unique_listdir(path):
    """
    returns list of directories in path without temporary files
    """

    dirs = os.listdir(path)
    count = 0
    while count < len(dirs):
        if dirs[count][:2] == "~$":
            del dirs[count]
        else:
            count += 1

    return dirs


def test_mapping_nols(infile):
    """
    Ensures all pass_on codes has assigned account code and process
    """
    path_mapp = Path(__file__).parents[1] / "mapping/mapping_nols.xlsx"
    df = pd.read_excel(path_mapp)
    df["test"] = 1
    df = df[["Kodeforklaring", "test"]]
    infile["Kodeforklaring"] = infile.Kodeforklaring.str.strip()
    df["Kodeforklaring"] = df.Kodeforklaring.str.strip()
    infile = infile.merge(df, on="Kodeforklaring", how="left")
    infile = infile.loc[infile.test != 1, :]
    rest = infile.Kodeforklaring.unique()
    str_ = ""
    if len(rest) > 0:
        str_ = str_ + "Kodeforklaring missing in mapping_nols:\n"
        for pic in rest:
            str_ = str_ + str(pic)
            str_ = str_ + "\n"

        raise LookupError(str_)


def kontering_NN(infile):
    """
    Konteringsgenerator Nettverk Norden
    """

    # clean konteringsark folder
    for direc in unique_listdir(Path(__file__).parents[1] / "konteringsark"):
        try:
            os.remove(Path(__file__).parents[1] / "konteringsark" / direc)
        except:
            continue

    today = dt.date.today()
    yr = str(today.year)
    if today.month < 10:
        mnth = "0%s" % str(today.month)
    else:
        mnth = str(today.month)
    YrMnth = "%s - %s" % (yr, mnth)

    curr_period = "%s.%s.%s" % (mnth, mnth, yr)

    yr = int(YrMnth[:4])
    mnth = int(YrMnth[7:])

    if mnth == 12:
        next_yr = str(yr + 1)
        next_mnth = "01"
    elif mnth < 9:
        next_mnth = "0" + str(mnth + 1)
        next_yr = str(yr)
    else:
        next_mnth = str(mnth + 1)
        next_yr = str(yr)

    next_YrMnth = "%s - %s" % (next_yr, next_mnth)
    next_period = "%s.%s.%s" % (next_mnth, next_mnth, next_yr)

    path_konteringskoder = Path(__file__).parents[1] / "mapping/mapping_nols.xlsx"

    infile = infile[
        [
            "Enhetsnummer",
            "RGNO",
            "IVNO",
            "Kodeforklaring",
            "PERIODE",
            "IVAM",
            "MOMS",
            "IVAM_INK_MOMS",
            "VTCD",
            "CUNO",
        ]
    ]
    infile = infile.reset_index(drop=True)
    infile.Kodeforklaring = infile.loc[:, "Kodeforklaring"].str.strip()

    infile_konteringskoder = pd.read_excel(path_konteringskoder)
    infile_konteringskoder.Kodeforklaring = (
        infile_konteringskoder.Kodeforklaring.str.strip()
    )
    infile_konteringskoder.drop(columns=["Periodisering"], inplace=True)

    # infile.IVAM = infile.IVAM.str.replace(",", ".")
    # infile.MOMS = infile.MOMS.str.replace(",", ".")
    # infile.IVAM_INK_MOMS = infile.IVAM_INK_MOMS.str.replace(",", ".")

    infile = infile.astype({"IVAM": float, "MOMS": float, "IVAM_INK_MOMS": float})

    infile["fakt_periode"] = infile.PERIODE.str[:4] + " - " + infile.PERIODE.str[4:6]
    infile["Antall perioder [Num1]"] = ["" for i in range(len(infile))]
    infile.loc[infile.fakt_periode == next_YrMnth, "Antall perioder [Num1]"] = 1

    # infile["moms"] = infile.MOMS.abs() > 0
    infile = infile.groupby(
        [
            "Enhetsnummer",
            "IVNO",
            "Kodeforklaring",
            "Antall perioder [Num1]",
            "VTCD",
            "CUNO",
        ],
        as_index=False,
    ).agg("sum", numeric_only=True)
    infile = infile.merge(infile_konteringskoder, how="left", on="Kodeforklaring")
    infile = infile.groupby(
        [
            "Enhetsnummer",
            "IVNO",
            "Antall perioder [Num1]",
            "Konto",
            "Kontonavn",
            "Prosess",
            "Prosessnavn Posten",
            "Konto tekst",
            "VTCD",
            "CUNO",
        ],
        as_index=False,
    ).agg("sum", numeric_only=True)

    for ivno in infile.IVNO.unique():
        curr = infile.loc[infile.IVNO == ivno, :]
        curr.reset_index(inplace=True)

        empty = ["" for i in range(len(curr))]
        if curr.CUNO[0] == cuno_pbb:
            balanserende_segment = ["000726" for i in range(len(curr))]
        else:
            balanserende_segment = ["000020" for i in range(len(curr))]
        kontokode = curr.Konto.copy()
        enhet = curr.Enhetsnummer.copy()
        prosess = curr.Prosess.copy()
        perioder = curr["Antall perioder [Num1]"].copy()
        kommentar = []
        for i in range(len(curr)):
            if perioder[i] == 1:
                kommentar.append(
                    next_YrMnth[:4]
                    + "_"
                    + next_YrMnth[-2:]
                    + "_"
                    + curr["Konto tekst"][i]
                )
            else:
                kommentar.append(
                    YrMnth[:4] + "_" + YrMnth[-2:] + "_" + curr["Konto tekst"][i]
                )

        # lastbærer = curr.RGNO.copy()

        VTCD = curr.VTCD
        net = curr.IVAM.copy()
        moms = curr.MOMS.copy()
        # net.loc[moms.abs() > 0] = moms.loc[moms.abs() > 0]*4
        net.loc[VTCD == 2] = moms.loc[VTCD == 2] * 4
        net.loc[VTCD == 4] = moms.loc[VTCD == 4] / (0.12)
        brutto = net + moms

        avgiftskode = pd.Series(np.zeros(len(curr)))
        # avgiftskode.loc[moms.abs() > 0] = 1
        avgiftskode.loc[VTCD == 2] = 1
        avgiftskode.loc[VTCD == 4] = 13

        perioder = curr["Antall perioder [Num1]"].copy()

        periodisering = pd.Series(["" for i in range(len(curr))])
        periodisering.loc[perioder == 1] = next_period

        df = pd.DataFrame()
        df["Neste godkjenner [NextApproverName]"] = empty
        df["Balanserende segment [Text61]"] = balanserende_segment
        df["Kontokode [AccountCode]"] = kontokode
        df["Enhetsnummer [CostCenterCode]"] = enhet
        df["Motpart [Text69]"] = empty
        df["Prosess [Text25]"] = prosess
        df["Prosjekt [ProjectCode]"] = empty
        df["Objekt [Text21]"] = empty
        df["Kommentar [LastComment]"] = kommentar
        # df["Lastbærer [Text29]"] = lastbærer
        df["Lastbærer [Text29]"] = empty
        df["Avgiftsprosent [TaxPercent]"] = empty
        df["Nettobeløp [NetSum]"] = net
        df["Avgiftsbeløp [TaxSum]"] = moms
        df["Bruttobeløp [GrossSum]"] = brutto
        df["Avgiftskode [TaxCode]"] = avgiftskode
        df["Service [Text65]"] = empty
        df["Periodisering start [Date1]"] = periodisering
        df["Antall perioder [Num1]"] = perioder
        df["Siste godkjenner [LatestApproverName]"] = empty
        df["Siste kontrollør [LatestReviewerName]"] = empty
        df["Artikkel [Text41]"] = empty
        df["Hovedkategori [Text37]"] = empty
        df["Underkategori [Text39]"] = empty
        df["Brukstid År [Text1]"] = empty
        df["Brukstid Mnd [Text2]"] = empty
        df["Brukstid kommentar [Text3]"] = empty
        df["Tilknyttet Aktiva [Text17]"] = empty
        df["Nettobeløp (selskap) [NetSumComp]"] = net
        df["Bruttobeløp (selskap) [GrossSumComp]"] = brutto
        df["Artikkel Beskrivelse [Text42]"] = empty
        df["Matchet nettobeløp [MatchedNetSum]"] = empty
        df["Matchet antall [MatchedQuantity]"] = empty
        df["Innkjøpsordrenummer [OrderNumber]"] = empty
        df["Ordrelinjernummer [OrderRowNumber]"] = empty
        df["Bestilt mengde [Num10]"] = empty
        df["Regnskapsår [Text24]"] = empty
        df["IC Partner [Text49]"] = empty

        # ørediff
        if np.abs((curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum())) > 100:
            print("Oh no! Ørediff for invoice %s is quite large!" % str(ivno))
        diff = pd.DataFrame()
        diff["Prosess [Text25]"] = ["0000"]
        diff["Kontokode [AccountCode]"] = ["779000"]
        diff["Enhetsnummer [CostCenterCode]"] = enhet[0]

        if curr.CUNO[0] == 99389:
            diff["Balanserende segment [Text61]"] = ["000726"]
        else:
            diff["Balanserende segment [Text61]"] = ["000020"]

        diff["Kommentar [LastComment]"] = ["%s_Ørediff" % YrMnth]
        diff["Avgiftskode [TaxCode]"] = [0]
        diff["Nettobeløp [NetSum]"] = curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        diff["Avgiftsbeløp [TaxSum]"] = [0]
        diff["Bruttobeløp [GrossSum]"] = (
            curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        )
        diff["Nettobeløp (selskap) [NetSumComp]"] = (
            curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        )
        diff["Bruttobeløp (selskap) [GrossSumComp]"] = (
            curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        )

        df = pd.concat([df, diff])

        try:
            df.to_excel(
                Path(__file__).parents[1] / Path("konteringsark/%s.xlsx" % str(ivno)),
                index=False,
            )
        except:
            pass


def get_grunnlag():
    """
    Retrives newest invoice data from folder
    """
    path = Path(__file__).parents[1] / "data"
    direc = unique_listdir(path)
    direc.sort()
    df = pd.read_excel(path / direc[-1], sheet_name="Grunnlag_nols")

    df_mapping = pd.read_excel(Path(__file__).parents[1] / "mapping/mapping_lp.xlsx")
    df_mapping = df_mapping.rename(
        columns={
            "Kostsenternummer": "kostsenter",
            "Kostsenter, beskrivelse": "Enhetsnummer",
        }
    )

    df = df.merge(df_mapping, on="kostsenter", how="left")
    df = df.reset_index(drop=True)
    return df


def get_grunnlag_passon():
    """
    Retrieves newest pass-on invoice data from folder
    """
    path = Path(__file__).parents[1] / "data"
    direc = unique_listdir(path)
    direc.sort()
    df = pd.read_excel(path / direc[-1], sheet_name="Grunnlag_pass_on")

    df_mapping = pd.read_excel(Path(__file__).parents[1] / "mapping/mapping_lp.xlsx")
    df_mapping = df_mapping.rename(
        columns={
            "Kostsenternummer": "kostsenter",
            "Kostsenter, beskrivelse": "Enhetsnummer",
        }
    )

    df = df.merge(df_mapping, on="kostsenter", how="left")
    df = df.reset_index(drop=True)
    return df


def test_mapping_pass_on(infile):
    """
    Ensures all pass_on codes has assigned account code and process
    """
    path_mapp = Path(__file__).parents[1] / "mapping/mapping_pass_on.xlsx"
    df = pd.read_excel(path_mapp)
    df["test"] = 1
    df = df[["picd", "test"]]
    infile = infile.merge(df, on="picd", how="left")
    infile = infile.loc[infile.test != 1, :]
    rest = infile.picd.unique()
    str_ = ""
    if len(rest) > 0:
        str_ = str_ + "picd missing in mapping_pass_on:\n"
        for pic in rest:
            str_ = str_ + str(pic)
            str_ = str_ + "\n"

        raise LookupError(str_)


def kontering_pass_on(infile, rgno=False):

    today = dt.date.today()
    yr = str(today.year)
    if today.month < 10:
        mnth = "0%s" % str(today.month)
    else:
        mnth = str(today.month)
    YrMnth = "%s - %s" % (yr, mnth)

    infile = infile.copy()

    infile = infile[
        ["kundnr", "picd", "rgno", "piam", "pivt", "piiv", "mvA", "Enhetsnummer"]
    ]

    df = infile.groupby(
        ["Enhetsnummer", "kundnr", "picd", "rgno", "pivt", "piiv"], as_index=False
    ).agg("sum", numeric_only=True)

    if not rgno:
        df = infile.groupby(
            ["Enhetsnummer", "kundnr", "picd", "pivt", "piiv"], as_index=False
        ).agg("sum", numeric_only=True)

    path_konteringskoder = Path(__file__).parents[1] / "mapping/mapping_pass_on.xlsx"
    infile_konteringskoder = pd.read_excel(path_konteringskoder)

    df = df.merge(infile_konteringskoder, on="picd", how="left")

    df = df.groupby(
        [
            "Enhetsnummer",
            "kundnr",
            "pivt",
            "piiv",
            "pids",
            "Konto",
            "Kontonavn",
            "Prosess",
            "Prosessnavn Posten",
            "Konto tekst",
        ],
        as_index=False,
    ).agg("sum", numeric_only=True)
    df = df.sort_values(by="Konto tekst")

    for piiv in df.piiv.unique():
        curr = df.loc[df.piiv == piiv, :]
        curr = curr.reset_index(drop=True)

        empty = ["" for i in range(len(curr))]
        if curr.kundnr[0] == cuno_pbb:
            balanserende_segment = ["000726" for i in range(len(curr))]
        else:
            balanserende_segment = ["000020" for i in range(len(curr))]
        kontokode = curr.Konto.copy()
        enhet = curr.Enhetsnummer.copy()
        prosess = curr.Prosess.copy()
        perioder = empty
        kommentar = []
        for i in range(len(curr)):
            kommentar.append(
                YrMnth[:4] + "_" + YrMnth[-2:] + "_" + curr["Konto tekst"][i]
            )

        if rgno:
            lastbærer = curr.RGNO.copy()
        else:
            lastbærer = empty

        VTCD = curr.pivt
        net = curr.piam.copy()
        moms = curr.mvA.copy()

        net.loc[VTCD == 2] = moms.loc[VTCD == 2] * 4
        net.loc[VTCD == 4] = moms.loc[VTCD == 4] / (0.12)
        brutto = net + moms

        avgiftskode = pd.Series(np.zeros(len(curr)))
        # # avgiftskode.loc[moms.abs() > 0] = 1
        avgiftskode.loc[VTCD == 2] = 1
        avgiftskode.loc[VTCD == 4] = 13

        periodisering = empty

        df2 = pd.DataFrame()
        df2["Neste godkjenner [NextApproverName]"] = empty
        df2["Balanserende segment [Text61]"] = balanserende_segment
        df2["Kontokode [AccountCode]"] = kontokode
        df2["Enhetsnummer [CostCenterCode]"] = enhet
        df2["Motpart [Text69]"] = empty
        df2["Prosess [Text25]"] = prosess
        df2["Prosjekt [ProjectCode]"] = empty
        df2["Objekt [Text21]"] = empty
        df2["Kommentar [LastComment]"] = kommentar
        # df2["Lastbærer [Text29]"] = lastbærer
        df2["Lastbærer [Text29]"] = empty
        df2["Avgiftsprosent [TaxPercent]"] = empty
        df2["Nettobeløp [NetSum]"] = net
        df2["Avgiftsbeløp [TaxSum]"] = moms
        df2["Bruttobeløp [GrossSum]"] = brutto
        df2["Avgiftskode [TaxCode]"] = avgiftskode
        df2["Service [Text65]"] = empty
        df2["Periodisering start [Date1]"] = periodisering
        df2["Antall perioder [Num1]"] = perioder
        df2["Siste godkjenner [LatestApproverName]"] = empty
        df2["Siste kontrollør [LatestReviewerName]"] = empty
        df2["Artikkel [Text41]"] = empty
        df2["Hovedkategori [Text37]"] = empty
        df2["Underkategori [Text39]"] = empty
        df2["Brukstid År [Text1]"] = empty
        df2["Brukstid Mnd [Text2]"] = empty
        df2["Brukstid kommentar [Text3]"] = empty
        df2["Tilknyttet Aktiva [Text17]"] = empty
        df2["Nettobeløp (selskap) [NetSumComp]"] = net
        df2["Bruttobeløp (selskap) [GrossSumComp]"] = brutto
        df2["Artikkel Beskrivelse [Text42]"] = empty
        df2["Matchet nettobeløp [MatchedNetSum]"] = empty
        df2["Matchet antall [MatchedQuantity]"] = empty
        df2["Innkjøpsordrenummer [OrderNumber]"] = empty
        df2["Ordrelinjernummer [OrderRowNumber]"] = empty
        df2["Bestilt mengde [Num10]"] = empty
        df2["Regnskapsår [Text24]"] = empty
        df2["IC Partner [Text49]"] = empty

        # ørediff
        if np.abs((curr.piam.sum() - df2["Nettobeløp [NetSum]"].sum())) > 100:
            print("Oh no! Ørediff for invoice %s is quite large!" % str(piiv))
        diff = pd.DataFrame()
        diff["Prosess [Text25]"] = ["0000"]
        diff["Kontokode [AccountCode]"] = ["779000"]
        diff["Enhetsnummer [CostCenterCode]"] = enhet[0]

        if curr.kundnr[0] == 99389:
            diff["Balanserende segment [Text61]"] = ["000726"]
        else:
            diff["Balanserende segment [Text61]"] = ["000020"]

        diff["Kommentar [LastComment]"] = ["%s_Ørediff" % YrMnth.replace(" - ", "_")]
        diff["Avgiftskode [TaxCode]"] = [0]
        diff["Nettobeløp [NetSum]"] = curr.piam.sum() - df2["Nettobeløp [NetSum]"].sum()
        diff["Avgiftsbeløp [TaxSum]"] = [0]
        diff["Bruttobeløp [GrossSum]"] = (
            curr.piam.sum() - df2["Nettobeløp [NetSum]"].sum()
        )
        diff["Nettobeløp (selskap) [NetSumComp]"] = (
            curr.piam.sum() - df2["Nettobeløp [NetSum]"].sum()
        )
        diff["Bruttobeløp (selskap) [GrossSumComp]"] = (
            curr.piam.sum() - df2["Nettobeløp [NetSum]"].sum()
        )

        df2 = pd.concat([df2, diff])

        outpath = Path(__file__).parents[1] / ("konteringsark/%s.xlsx" % (str(piiv)))
        try:
            df2.to_excel(outpath, index=False)
        except:
            pass


if __name__ == "__main__":
    if not os.path.exists(Path(__file__).parents[1] / "konteringsark"):
        os.mkdir(Path(__file__).parents[1] / "konteringsark")
    if not os.path.exists(Path(__file__).parents[1] / "data"):
        os.mkdir(Path(__file__).parents[1] / "data")
    create_mapping.create_mapping()
    infile = get_grunnlag()

    test_mapping_nols(infile)
    kontering_NN(infile)
    infile_passon = get_grunnlag_passon()
    test_mapping_pass_on(infile_passon)
    kontering_pass_on(infile_passon)
