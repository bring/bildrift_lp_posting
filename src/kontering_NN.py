# -*- coding: utf-8 -*-
"""
Created on Wed Jun 14 09:50:40 2023

@author: granenga
"""
#%%
import pandas as pd
import numpy as np
import datetime as dt
import os
from pathlib import Path

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

    return(dirs)

def kontering_NN(infile):
    """
    Konteringsgenerator Nettverk Norden
    """

    #clean konteringsark folder
    for direc in unique_listdir(Path(__file__).parents[1]/"konteringsark"):
        try:
            os.remove(Path(__file__).parents[1]/"konteringsark"/direc)
        except:
            continue
    
    today = dt.date.today()
    yr = str(today.year)
    if today.month < 10:
        mnth = "0%s" % str(today.month)
    else:
        mnth = str(today.month)
    YrMnth = "%s - %s" % (yr, mnth)
    
    curr_period = "10.%s.%s" % (mnth, yr)
    
    
    yr = int(YrMnth[:4])
    mnth = int(YrMnth[7:])
    
    if mnth == 12:
        next_yr = str(yr+1)
        next_mnth = "01"
    elif mnth < 9:
        next_mnth = "0" + str(mnth+1)
        next_yr = str(yr)
    else:
        next_mnth = str(mnth+1)
        next_yr = str(yr)
    
    next_YrMnth = "%s - %s" % (next_yr, next_mnth)
    next_period = "10.%s.%s" % (next_mnth, next_yr)
    
    path_konteringskoder = Path(__file__).parents[1] / "mapping/mapping_nols.xlsx"
    
    
    infile = infile[["Enhetsnummer", "RGNO", "IVNO","Kodeforklaring","PERIODE","IVAM", "MOMS", "IVAM_INK_MOMS", "VTCD"]]
    infile = infile.reset_index(drop = True)
    infile.Kodeforklaring = infile.loc[:, "Kodeforklaring"].str.strip()
    
    infile_konteringskoder = pd.read_excel(path_konteringskoder)
    infile_konteringskoder.Kodeforklaring = infile_konteringskoder.Kodeforklaring.str.strip()
    infile_konteringskoder.drop(columns = ["Periodisering"], inplace = True)
    
    
    
    # infile.IVAM = infile.IVAM.str.replace(",", ".")
    # infile.MOMS = infile.MOMS.str.replace(",", ".")
    # infile.IVAM_INK_MOMS = infile.IVAM_INK_MOMS.str.replace(",", ".")
        
    infile = infile.astype({"IVAM":float,
                            "MOMS":float,
                            "IVAM_INK_MOMS":float})
    
    infile["fakt_periode"] = infile.PERIODE.str[:4] + " - " + infile.PERIODE.str[4:6]
    infile["Antall perioder [Num1]"] = ["" for i in range(len(infile))]
    infile.loc[infile.fakt_periode == next_YrMnth,"Antall perioder [Num1]"] = 1
    
    # infile["moms"] = infile.MOMS.abs() > 0
    infile = infile.groupby(["Enhetsnummer", "IVNO", "Kodeforklaring", "Antall perioder [Num1]", "VTCD"], as_index = False).agg("sum", numeric_only = True)
    infile = infile.merge(infile_konteringskoder, how = "left", on = "Kodeforklaring")
    infile = infile.groupby(["Enhetsnummer", "IVNO", "Antall perioder [Num1]", 'Konto','Kontonavn', 'Prosess', 'Prosessnavn Posten', 'Konto tekst', "VTCD"], as_index = False).agg("sum", numeric_only = True)
    
    
    for ivno in infile.IVNO.unique():
        curr = infile.loc[infile.IVNO == ivno, :]
        curr.reset_index(inplace = True)

        empty = ["" for i in range(len(curr))]
        balanserende_segment = ["000020" for i in range(len(curr))]
        kontokode = curr.Konto.copy()
        enhet = curr.Enhetsnummer.copy()
        prosess = curr.Prosess.copy()
        perioder = curr["Antall perioder [Num1]"].copy()
        kommentar = []
        for i in range(len(curr)):
            if perioder[i] == 1:
                kommentar.append(next_YrMnth[:4] + "_" + next_YrMnth[-2:] + "_" + curr["Konto tekst"][i])
            else:
                kommentar.append(YrMnth[:4] + "_" + YrMnth[-2:] + "_" + curr["Konto tekst"][i])
                
        # lastbærer = curr.RGNO.copy()
        
        VTCD = curr.VTCD
        net = curr.IVAM.copy()
        moms = curr.MOMS.copy()
        # net.loc[moms.abs() > 0] = moms.loc[moms.abs() > 0]*4
        net.loc[VTCD == 2] = moms.loc[VTCD == 2]*4
        net.loc[VTCD == 4] = moms.loc[VTCD == 4]/(0.12)
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
        
        #ørediff
        if np.abs((curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum())) > 100:
            print("Oh no! Ørediff for invoice %s is quite large!" % str(ivno))
        diff = pd.DataFrame()
        diff["Prosess [Text25]"] = ["0000"]
        diff["Kontokode [AccountCode]"] = ["779000"]
        diff["Enhetsnummer [CostCenterCode]"] = enhet[0]
        diff["Balanserende segment [Text61]"] = ["000020"]
        diff["Kommentar [LastComment]"] = ["%s_Ørediff" % YrMnth]
        diff["Avgiftskode [TaxCode]"] = [0]
        diff["Nettobeløp [NetSum]"] = curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        diff["Avgiftsbeløp [TaxSum]"] = [0]
        diff["Bruttobeløp [GrossSum]"] = curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        diff["Nettobeløp (selskap) [NetSumComp]"] = curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        diff["Bruttobeløp (selskap) [GrossSumComp]"] = curr.IVAM.sum() - df["Nettobeløp [NetSum]"].sum()
        
        df = pd.concat([df, diff])
        df.to_excel(Path(__file__).parents[1] / Path("konteringsark/%s.xlsx" % str(ivno)), index = False)


def kontering_per_process_pass_on(infile_varekost):
    """
    Reads Leaseplan Varekost rapport
    
    Creates one excel sheet for every invoice number
    to be used in Basware
    Creates one konteringsline for each process and tax percentage.
    """

    today = dt.date.today()
    yr = str(today.year)
    if today.month < 10:
        mnth = "0%s" % str(today.month)
    else:
        mnth = str(today.month)
    YrMnth = "%s - %s" % (yr, mnth)
    
    path_konteringskoder = Path(__file__).parents[1] / "mapping/mapping_pass_on.xlsx"
    infile_konteringskoder = pd.read_excel(path_konteringskoder)
    
    invoice_numbers = infile_varekost["piiv"].to_numpy(dtype=str)
    kodeforklaring_pass_on = infile_varekost["picd"].to_numpy(dtype=str)
    kodeforklaring_pass_on = np.char.strip(kodeforklaring_pass_on)
    net_nols = infile_varekost["piam"].to_numpy()
    moms_nols = infile_varekost["mvA"].to_numpy()
    brutto_nols = infile_varekost["piam_ink_mva"].to_numpy()
    enhetsnummer = infile_varekost["Enhetsnummer"].to_numpy(dtype=str)
    lastbærer_nols = infile_varekost["rgno"].to_numpy(dtype = str)
    for i in range(len(lastbærer_nols)):
        try:
            int(lastbærer_nols[i])
            lastbærer_nols[i] = ""
        except:
            pass
    
    
    tax = np.zeros_like(net_nols)
    tax[np.nonzero(net_nols)] = np.round(((np.abs(brutto_nols[np.nonzero(net_nols)])/np.abs(net_nols[np.nonzero(net_nols)]))-1)*100)
    avgiftskode = np.zeros_like(tax)
    avgiftskode[np.nonzero(tax)] == 1
    
    unique_invoice = np.unique(invoice_numbers)
    
    kodeforklaring_kode = infile_konteringskoder["picd"].to_numpy(dtype=str)
    kodeforklaring_kode = np.char.strip(kodeforklaring_kode)
    konto_kode = infile_konteringskoder["Konto"].to_numpy(dtype=str)
    kontonavn_kode = infile_konteringskoder["Kontonavn"].to_numpy(dtype=str)
    prosess_kode = infile_konteringskoder["Prosess"].to_numpy(dtype=str)
    prosessnavn_kode = infile_konteringskoder["Prosessnavn Posten"].to_numpy(dtype=str)
    kontotekst_kode = infile_konteringskoder["Konto tekst"].to_numpy(dtype=str)
    kontotekst_pass_on = []
    for i in range(len(kodeforklaring_pass_on)):
        kontotekst_pass_on.append(kontotekst_kode[np.where(kodeforklaring_kode == kodeforklaring_pass_on[i])][0])
        

    kontotekst_pass_on = np.array(kontotekst_pass_on)
    

    taxcodes = [0, 1]
    for i in range(len(unique_invoice)):
        inds = np.where(invoice_numbers == unique_invoice[i])
        
        

        tax_prosess = []
        net_prosess = []
        moms_prosess = []
        brutto_prosess = []
        avgiftskode_prosess = []
        kontokoder = []
        prosesser = []
        kontotekster = []
        balanserende_segment = []
        
        L = 0
        net_rest = 0
        for j in range(len(np.unique(kontotekst_kode))):
            for k in range(len(taxcodes)):
                req1 = kontotekst_pass_on == np.unique(kontotekst_kode)[j]
                req2 = avgiftskode == taxcodes[k]
                req3 = invoice_numbers == unique_invoice[i]
                indices = np.where(req1 & req2 & req3)
                temp_net = np.sum(net_nols[indices])
                temp_moms = np.sum(moms_nols[indices])
                
                if temp_moms != 0:
                    curr_net = 4*temp_moms
                    curr_brutto = 5*temp_moms
                    net_rest = temp_net - 4*temp_moms
                
                if temp_moms == 0:
                    curr_net = temp_net + net_rest
                    curr_brutto = curr_net
                    net_rest = 0
                
                if np.abs(curr_brutto) > 1e-3:
                    net_prosess.append(curr_net)
                    moms_prosess.append(temp_moms)
                    brutto_prosess.append(curr_brutto)
                    
                    if brutto_prosess[L] != 0:
                        tax_prosess.append(np.round((brutto_prosess[L]/net_prosess[L]-1)*100))
                    else:
                        tax_prosess.append(0)
                    if tax_prosess[L] != 0:
                        avgiftskode_prosess.append(1)
                    else:
                        avgiftskode_prosess.append(0)
                    kontokoder.append(konto_kode[np.where(kontotekst_kode == np.unique(kontotekst_kode)[j])][0])
                    prosesser.append(prosess_kode[np.where(kontotekst_kode == np.unique(kontotekst_kode)[j])][0])
                    kontotekster.append(kontotekst_kode[np.where(kontotekst_kode == np.unique(kontotekst_kode)[j])][0])
                    L += 1
                
        empty = []
        for j in range(len(prosesser)):
            empty.append("")
        empty = np.array(empty)
        
        enhet = enhetsnummer[inds][0]
        
        balanserende_segment = []
        for n in range(len(prosesser)):
            segment = "000020"
            balanserende_segment.append(segment)
        
        df = pd.DataFrame()
        df["Neste godkjenner [NextApproverName]"] = empty
        df["Balanserende segment [Text61]"] = balanserende_segment
        df["Kontokode [AccountCode]"] = kontokoder
        df["Enhetsnummer [CostCenterCode]"] = [enhetsnummer[inds][0] for prosess in prosesser]
        df["Motpart [Text69]"] = empty
        df["Prosess [Text25]"] = prosesser
        df["Prosjekt [ProjectCode]"] = empty
        df["Objekt [Text21]"] = empty
        df["Kommentar [LastComment]"] = [YrMnth[:4] + "_" + YrMnth[-2:] + "_" + kontotekst for kontotekst in kontotekster]
        df["Lastbærer [Text29]"] = empty
        df["Avgiftsprosent [TaxPercent]"] = empty
        df["Nettobeløp [NetSum]"] = net_prosess
        df["Avgiftsbeløp [TaxSum]"] = moms_prosess
        df["Bruttobeløp [GrossSum]"] = brutto_prosess
        df["Avgiftskode [TaxCode]"] = avgiftskode_prosess
        df["Service [Text65]"] = empty
        df["Periodisering start [Date1]"] = empty
        df["Antall perioder [Num1]"] = empty
        df["Siste godkjenner [LatestApproverName]"] = empty
        df["Siste kontrollør [LatestReviewerName]"] = empty
        df["Artikkel [Text41]"] = empty
        df["Hovedkategori [Text37]"] = empty
        df["Underkategori [Text39]"] = empty
        df["Brukstid År [Text1]"] = empty
        df["Brukstid Mnd [Text2]"] = empty
        df["Brukstid kommentar [Text3]"] = empty
        df["Tilknyttet Aktiva [Text17]"] = empty
        df["Nettobeløp (selskap) [NetSumComp]"] = net_prosess
        df["Bruttobeløp (selskap) [GrossSumComp]"] = brutto_prosess
        df["Artikkel Beskrivelse [Text42]"] = empty
        df["Matchet nettobeløp [MatchedNetSum]"] = empty
        df["Matchet antall [MatchedQuantity]"] = empty
        df["Innkjøpsordrenummer [OrderNumber]"] = empty
        df["Ordrelinjernummer [OrderRowNumber]"] = empty
        df["Bestilt mengde [Num10]"] = empty
        df["Regnskapsår [Text24]"] = empty
        df["IC Partner [Text49]"] = empty
        
        outpath = Path(__file__).parents[1] / ("konteringsark/%s.xlsx" % (unique_invoice[i]))
        df.to_excel(outpath, index=False)
        
def get_grunnlag():
    """
    Retrives newest invoice data from folder
    """
    path = Path(__file__).parents[1] / "data"
    direc = unique_listdir(path)
    direc.sort()
    df = pd.read_excel(path / direc[-1], sheet_name = "Grunnlag_nols")

    df_mapping = pd.read_excel(Path(__file__).parents[1] / "mapping/mapping_lp.xlsx")
    df_mapping = df_mapping.rename(columns = {
        "Kostsenternummer":"kostsenter",
        "Kostsenter, beskrivelse":"Enhetsnummer"
    })

    df = df.merge(df_mapping, on = "kostsenter", how = "left")
    df = df.reset_index(drop = True)
    return(df)

def get_grunnlag_passon():
    """
    Retrieves newest pass-on invoice data from folder
    """
    path = Path(__file__).parents[1] / "data"
    direc = unique_listdir(path)
    direc.sort()
    df = pd.read_excel(path / direc[-1], sheet_name = "Grunnlag_pass_on")

    df_mapping = pd.read_excel(Path(__file__).parents[1] / "mapping/mapping_lp.xlsx")
    df_mapping = df_mapping.rename(columns = {
        "Kostsenternummer":"kostsenter",
        "Kostsenter, beskrivelse":"Enhetsnummer"
    })

    df = df.merge(df_mapping, on = "kostsenter", how = "left")
    df = df.reset_index(drop = True)
    return(df)

if __name__ == "__main__":
    infile = get_grunnlag()
    kontering_NN(infile)
    infile_passon = get_grunnlag_passon()
    kontering_per_process_pass_on(infile_passon)

# %%
