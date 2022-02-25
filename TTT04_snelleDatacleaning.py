# Script om een adlibexport.csv bestand te controleren en alle records te returnen met ontbrekende waarden
# Een snelle manier om ontbrekende data te identificeren en daarna in adlib/axiell aan te vullen


import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# NODIG: een adlib-export met de kolommen die je wil controleren
# lees de csv in: BESTANDSNAAM AANPASSEN en sla op in een dataframe
df = pd.read_csv(r"export.csv", delimiter=';')

# maak een lijst van alle kolomnamen (gekozen adlibvelden in export)
velden = list(df.columns)

# maak een nieuw workbook
wb = Workbook()

# voor elke kolom: maakt worksheet aan met "geen x" als titel + maakt dataframe met alle records met ontbrekende waarde
# + bewaar die lijst en print die als rijen in het aangemaakte worksheet + sla de creatie op in de projectmap
for i in velden:
    ws = wb.create_sheet("geen " + str(i))
    i = pd.isna(df[i])
    i = df[i]
    rows = dataframe_to_rows(i, index=False)
    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(r"datacleaning.xlsx")
