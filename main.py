#######################################################################################################################
# A. Importeer csv (uit adlib) naar dataframe (pandas)

# pip install pandas
import pandas as pd

# inlezen csv
# ingeven padnaam naar csv, en indien van toepassing aanvullen delimiter
lijstDMG = pd.read_csv(r"export (16).csv", delimiter=';')
print(lijstDMG)
