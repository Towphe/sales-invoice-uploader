import pandas as pd

def match_soa_and_sorep(soa, sorep):    
    # join two tables together
    sorep_soa_matched = pd.merge(soa, sorep, how="left", left_on="Order No.", right_on="Order No.")
    sorep_soa_matched = sorep_soa_matched[['Transaction Date', 'S. Order #', 'Order No.', 'Reference 1', 'Reference 2', 'Reference 3', 'Reference 4', 'Reference 5', 'Amount_y']]
    
    agg_funcs = {'Transaction Date' : 'first', 'S. Order #' : 'first', 'Order No.' : 'first', 'Reference 1' : 'first', 'Reference 2': 'first', 'Reference 3' : 'first', 'Reference 4' : 'first', 'Reference 5' : 'first', 'Amount_y':'sum',}
    sorep_soa_grouped = sorep_soa_matched.groupby(['Order No.']).agg(agg_funcs)

    return sorep_soa_grouped

def join_soreg_and_sorep_and_soa(groupd_soa_and_sorep, soreg):
    matched_soreg_sorep_soa = pd.merge(groupd_soa_and_sorep, soreg, left_on="S. Order #", right_on="SO #", how="left")
    return matched_soreg_sorep_soa

def create_template(soa_dir:str, sorep_dir:str, soreg_dir:str):
    # get soa data
    soa = pd.read_excel(soa_dir)
    filtered_soa = soa[(soa["Transaction Type"] == "Orders-Sales") | (soa["Transaction Type"] == "Refund-Claims")]

    # get sorep data
    sorep = pd.read_excel(sorep_dir)
    filtered_sorep = sorep[sorep["Name"] == "JABRA PH - ONLINE SALES"]

    sorep_soa_grouped = match_soa_and_sorep(filtered_soa, filtered_sorep)

    # filter to SI Report later
    # => if order already has SI, disregard

    #Batching and joining with SO Register

    soreg = pd.read_excel(soreg_dir)

    matched_soreg_sorep_soa = join_soreg_and_sorep_and_soa(sorep_soa_grouped, soreg)

    print(matched_soreg_sorep_soa)
    
    # create file
    #matched_soreg_sorep_soa.to_excel('output.xlsx', sheet_name="Sheet 1", index=False)
 
# test data   
#create_template('../../SalesInvoiceUploader-data/test-data/Dec2023 - Lazada SOA.xlsx', '../../SalesInvoiceUploader-data/test-data/Dec2023 - Sales Order Report.xlsx', '../../SalesInvoiceUploader-data/test-data/Dec2023 - Sales Order Register.xlsx')