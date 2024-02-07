import numpy as np
import pandas as pd

files = {
    "soa" : "/home/tope/Desktop/SOA Laz Jabra Jan 1-11 2024.xlsx",
    "so" : "/home/tope/Desktop/625Tech SO Jan 1-11 2024.xlsx",
    "si" : "/home/tope/Desktop/625Tech SI Jan 1-11 2024.xlsx",
    "sor" : "/home/tope/Desktop/625Tech SOR Jan 1-11 2024.xlsx"
}

# get soa file
soa_df = pd.read_excel(files["soa"])

# filter soa to only rows with column `Transaction Type` to `Order Sales` or `Refund Claims`
soa_df = soa_df[(soa_df["Transaction Type"] == "Orders-Sales") | (["Transaction Type"] == "Refunds-Claims")]

print(soa_df.shape)

# open si file
si_df = pd.read_excel(files["si"])

# filter soa to entries with NO entry yet in SI file
soa_df = soa_df[~(soa_df["Order No."].isin(si_df["Reference No"]))]

# TASK: link to SO file

# get so file
so_df = pd.read_excel(files["so"])

# format types
soa_df[['Order No.']] = soa_df[['Order No.']].astype(str)
so_df[['Reference No']] = so_df[['Reference No']].astype(str)

# match soa and so
matched_soa_so = pd.merge(soa_df, so_df, how='left', left_on='Order No.', right_on='Reference No')

print(matched_soa_so.columns)

# aggregate according to order no., s. order #, and sku
so_soa_agg = {'Transaction Date' : 'first', 'Reference 1' : 'first', 'Reference 2': 'first', 'Reference 3' : 'first', 'Reference 4' : 'first', 'Reference 5' : 'first', 'Amount':'first',}
matched_soa_so_grouped = matched_soa_so.groupby(['Order No.', 'S. Order #', 'Seller SKU']).agg(so_soa_agg)

matched_soa_so_grouped.to_excel('test.xlsx')

# TASK: cross-match with soreg

# get soreg file
soreg_df = pd.read_excel(files['sor'])

# merge previously grouped soa_so with soreg
matched_soa_so_sor = pd.merge(matched_soa_so_grouped, soreg_df, how='left', left_on="S. Order #", right_on='SO #')

matched_soa_so_sor.to_excel("test_output.xlsx")