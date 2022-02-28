# trainthetrainer_04

Tool data visualisatie aanpassen op maat van DMG (europeana.datacleaning.py):

    controleren van records met webpublicatie = veldnamen: objectnummer, instelling, objectnaam, titel, beschrijving, datum.vervaardiging.begin, datum.vervaardiging.eind, vervaardiger, vervaardiger.rol, vervaardiging.plaats, rechten.type
      script maakt een excel-export met daarin de objectnummers waarin bepaalde velden niet zijn ingevuld en telt de aantallen + stelt ze grafisch voor in 3D-barchart

SCRIPT2: TTT04_snelleDatacleaning.py
  script dat zelfde doet maar dan ongeacht de gekozen kolommen: leest een csv in en maakt dataframes voor ontbrekende waarden in elke kolom en slaat ze op in individuele sheets
   zonder visualisatie of counts maar flexibeler in gebruik
