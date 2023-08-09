# -*- coding: utf-8 -*-
"""
Created on Wed Jun 14 09:50:40 2023

@author: granenga
"""

import pandas as pd
import numpy as np
import datetime as dt
from pathlib import Path



def kontering_NN(infile, YrMnth = False):
    """
    Konteringsgenerator BLB
    """
    if YrMnth == False:
        today = dt.date.today()
        yr = str(today.year)
        if today.month < 10:
            mnth = "0%s" % str(today.month)
        else:
            mnth = str(today.month)
        YrMnth = "%s - %s" % (yr, mnth)
    else:
        yr = YrMnth[:4]
        mnth = YrMnth[7:]
    
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
    
    
    path_konteringskoder = "F:/Nettverk Norge/FO Bildrift/60 Analyse/Rapportmaler/08 Leaseplan kontering/mapping_nols.xlsx"
    
    # path_varekost = "F:/Nettverk Norge/FO Bildrift/30 Økonomi/02 Regnskapsrapporter/Leaseplan varekost/Varekost rapport.xlsx"
    # infile = pd.read_excel(path_varekost, sheet_name = "grunnlag hovedfaktura")
    
    # infile = pd.read_csv("F:/Nettverk Norge/FO Bildrift/60 Analyse/Kontering/kontering BLB/grunnlag_nols.csv", sep = ";", encoding = "latin_1", low_memory=False)
    
    infile = infile.loc[infile.ÅrMnd == YrMnth, :]
    infile = infile.loc[infile.Selskap.str.contains("Posten"), :]
    infile = infile[["Enhetsnummer", "RGNO", "IVNO","Kodeforklaring","PERIODE","IVAM", "MOMS", "IVAM_INK_MOMS", "VTCD"]]
    infile.Kodeforklaring = infile.Kodeforklaring.str.strip()
    
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
        df.to_excel(Path(__file__).parents[0] / Path("PBAS/%s.xlsx" % str(ivno)), index = False)
        
if __name__ == "__main__":
    path_varekost = "F:/Nettverk Norge/FO Bildrift/30 Økonomi/02 Regnskapsrapporter/Leaseplan varekost/Varekost rapport.xlsx"
    infile = pd.read_excel(path_varekost, sheet_name = "grunnlag hovedfaktura")
    kontering_NN(infile)