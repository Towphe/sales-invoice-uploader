import pandas as pd
import numpy as np
import calendar

month_index = {
    "January" : 1,
    "February" : 2,
    "March" : 3,
    "April" : 4,
    "May" : 5,
    "June" : 6,
    "July" : 7,
    "August" : 8,
    "September" : 9,
    "October" : 10,
    "November" : 11,
    "December" : 12
}

def match_soa_and_sorep(soa, sorep):    
    sorep_soa_matched = pd.DataFrame()

    soa[['Order No.']] = soa[['Order No.']].astype(str)
    sorep[['Reference No']] = sorep[['Reference No']].astype(str)

    try:
        # join two tables together
        
        sorep_soa_matched = pd.merge(soa, sorep, how="left", left_on="Order No.", right_on="Reference No")
    except:
        return False
    
    sorep_soa_grouped = pd.DataFrame()
    try:
        sorep_soa_matched = sorep_soa_matched[['Transaction Date', 'S. Order #', 'Order No.', 'Reference 1', 'Reference 2', 'Reference 3', 'Reference 4', 'Reference 5', 'Amount']]
        
        agg_funcs = {'Transaction Date' : 'first', 'S. Order #' : 'first', 'Order No.' : 'first', 'Reference 1' : 'first', 'Reference 2': 'first', 'Reference 3' : 'first', 'Reference 4' : 'first', 'Reference 5' : 'first', 'Amount':'first',}
        sorep_soa_grouped = sorep_soa_matched.groupby(['Order No.']).agg(agg_funcs)
    except:
        return False
    
    return sorep_soa_grouped

def join_soreg_and_sorep_and_soa(groupd_soa_and_sorep, soreg):
    matched_soreg_sorep_soa = pd.DataFrame()
    try:
        matched_soreg_sorep_soa = pd.merge(groupd_soa_and_sorep, soreg, left_on="S. Order #", right_on="SO #", how="left")
    except:
        return False
    return matched_soreg_sorep_soa

def check_if_existing_si(matched_soreg_sorep_soa, sirep):
    #soreg_sorep_wo_si = matched_soreg_sorep_soa[matched_soreg_sorep_soa["S. Order #"].isin(sirep["Transfer From"]) == False]
    soreg_sorep_wo_si = pd.DataFrame()
    try:
        soreg_sorep_wo_si = matched_soreg_sorep_soa[~matched_soreg_sorep_soa["S. Order #"].isin(sirep["Transfer From"])]
    except:
        return False
    return soreg_sorep_wo_si    

def create_template(so_start:str, si_month:str, year:int, soa_dir:str, sorep_dir:str, sirep_dir:str, soreg_dir:str):
    # get soa data
    soa = pd.read_excel(soa_dir)
    filtered_soa = pd.DataFrame()

    # soa = soa.astype({"Order No.": "string"})
    # sorep = sorep.astype({"Reference No":"string"})
    try:
        filtered_soa = soa[(soa["Transaction Type"] == "Orders-Sales") | (soa["Transaction Type"] == "Refunds-Claims")]
    except:
        return ("Invalid Lazada SOA file", False)

    # get sorep data
    sorep = pd.read_excel(sorep_dir)
    filtered_sorep = pd.DataFrame()
    try:
        filtered_sorep = sorep[sorep["Name"] == "JABRA PH - ONLINE SALES"]
    except:
        return ("Invalid SO Report file", False)

    sorep_soa_grouped = match_soa_and_sorep(filtered_soa, filtered_sorep)
    if type(sorep_soa_grouped) == bool:
        # say invalid
        return ("Invalid SOA ", False)

    sirep = pd.read_excel(sirep_dir)

    # filter to SI Report later
    # => if order already has SI, disregard
    try:
        sorep_soa_grouped = check_if_existing_si(sorep_soa_grouped, sirep)
    except:
        return ("Error grouping Sales Orders", False)

    if type(sorep_soa_grouped) == bool:
        return ("Error grouping Sales Orders", False)
    
    if sorep_soa_grouped.shape[0] == 0:
        # empty dataframe
        return ("No match", False)

    #Batching and joining with SO Register

    soreg = pd.read_excel(soreg_dir)

    matched_soreg_sorep_soa = join_soreg_and_sorep_and_soa(sorep_soa_grouped, soreg)

    if type(matched_soreg_sorep_soa) is bool:
        return ("Invalid SO Register File", False)

    last_day = get_last_day(si_month, year)


    #matched_soreg_sorep_soa["SalesInvoiceCode"] = generate_si_batch(int(so_start), int(so_start) + len(matched_soreg_sorep_soa))
    matched_soreg_sorep_soa["SalesInvoiceCode"] = ""
    #matched_soreg_sorep_soa.insert(0, 'SalesInvoiceCode', range(123, 123 + len(matched_soreg_sorep_soa)))
    matched_soreg_sorep_soa["Order No."] = pd.to_numeric(matched_soreg_sorep_soa["Order No."])

    matched_soreg_sorep_soa["SalesInvoiceDate"] = last_day # get invoice date
    matched_soreg_sorep_soa["PostingDate"] = last_day # get posting date
    matched_soreg_sorep_soa["OurDONO"] = ""
    matched_soreg_sorep_soa["DueDate"] = last_day # set as end of month
    matched_soreg_sorep_soa["GlobalProgressInvoicingRate"] = ""
    matched_soreg_sorep_soa["IsApproved"] = True    
    matched_soreg_sorep_soa["IsDeferredVAT"] = False
    matched_soreg_sorep_soa["TaxDate"] = last_day
    matched_soreg_sorep_soa["Debtor"] = "102-J001" # allow customization
    matched_soreg_sorep_soa["CurrencyRate"] = "N8" # allow customization
    matched_soreg_sorep_soa["ReverseRate"] = "N8" # allow customization
    matched_soreg_sorep_soa["SalesPerson"] = "LAZADA-JABRA"
    matched_soreg_sorep_soa["Term"] = "C.O.D."
    matched_soreg_sorep_soa["Remark1"] = matched_soreg_sorep_soa["S. Order #"]
    matched_soreg_sorep_soa["Remark2"] = ""
    matched_soreg_sorep_soa["Remark3"] = ""
    matched_soreg_sorep_soa["Remark4"] = ""
    matched_soreg_sorep_soa["Remark5"] = ""
    matched_soreg_sorep_soa["Project"] = ""
    matched_soreg_sorep_soa["StockLocation"] = ""
    matched_soreg_sorep_soa["DORegistationNo"] = ""
    matched_soreg_sorep_soa["DOArea"] = ""
    matched_soreg_sorep_soa["CostCentre"] = ""
    matched_soreg_sorep_soa["IsCancelled"] = ""
    matched_soreg_sorep_soa["IsTaxInclusive"] = 1
    matched_soreg_sorep_soa["IsRounding"] = ""
    matched_soreg_sorep_soa["IsNonTaxInvoice"] = ""
    matched_soreg_sorep_soa["none"] = ""
    #
    matched_soreg_sorep_soa["ProgressInvoicingRate"] = ""
    matched_soreg_sorep_soa["StockType"] = ""
    matched_soreg_sorep_soa["SerialNumber"] = ""
    matched_soreg_sorep_soa["StockBatchNumber"] = ""
    matched_soreg_sorep_soa["DebtorItem"] = ""
    matched_soreg_sorep_soa["ServiceCost"] = ""
    matched_soreg_sorep_soa["PackingUOM"] = ""
    matched_soreg_sorep_soa["Packing"] = ""
    matched_soreg_sorep_soa["PackingQty"] = ""
    matched_soreg_sorep_soa["Numbering"] = ""
    matched_soreg_sorep_soa["Stock"] = matched_soreg_sorep_soa["Stock #"]
    matched_soreg_sorep_soa["StockLocation"] = ""
    # ...
    matched_soreg_sorep_soa["UnitPrice"] = pd.to_numeric(matched_soreg_sorep_soa["Amount_x"])
    matched_soreg_sorep_soa["Qty"] = pd.to_numeric(matched_soreg_sorep_soa["Qty"]) 
    matched_soreg_sorep_soa["Discount"] = ""
    matched_soreg_sorep_soa["GLAccount"] = ""
    matched_soreg_sorep_soa["CostCentre"] = ""
    matched_soreg_sorep_soa["ReferenceNo"] = ""
    #
    matched_soreg_sorep_soa["Ref"] = ""
    matched_soreg_sorep_soa["Ref2"] = ""
    matched_soreg_sorep_soa["Ref3"] = ""
    matched_soreg_sorep_soa["Ref4"] = ""
    matched_soreg_sorep_soa["Ref5"] = ""
    matched_soreg_sorep_soa["DateRef1"] = ""
    matched_soreg_sorep_soa["DateRef2"] = ""
    matched_soreg_sorep_soa["NumRef1"] = ""
    matched_soreg_sorep_soa["NumRef2"] = ""
    matched_soreg_sorep_soa["TaxCode"] = "SR-SP"
    matched_soreg_sorep_soa["TariffCode"] = ""
    matched_soreg_sorep_soa["TaxRate"] = "12.00%"
    matched_soreg_sorep_soa["WTaxCode"] = ""
    matched_soreg_sorep_soa["WTaxRate"] = "0.00"
    matched_soreg_sorep_soa["WVatCode"] = "0.00"
    matched_soreg_sorep_soa["WVatRate"] = ""

    output_tab = matched_soreg_sorep_soa[["SalesInvoiceCode", "SalesInvoiceDate", "PostingDate", "OurDONO", "DueDate", "GlobalProgressInvoicingRate", "IsApproved", "IsDeferredVAT", "TaxDate", "Debtor", "CurrencyRate", "ReverseRate", "SalesPerson", "Term", "Order No.", "Reference 1", "Reference 2", "Reference 3", "Reference 4", "Reference 5", "Remark1", "Remark2", "Remark3", "Remark4", "Remark5", "Project", "StockLocation", "DORegistationNo", "DOArea", "CostCentre", "IsCancelled", "IsTaxInclusive", "IsRounding", "IsNonTaxInvoice", "none", "ProgressInvoicingRate", "StockType", "SerialNumber", "StockBatchNumber", "DebtorItem", "ServiceCost", "PackingUOM", "Packing", "PackingQty", "Numbering", "Stock", "StockLocation", "Qty", "UOM", "UnitPrice", "Discount", "GLAccount", "CostCentre", "Description", "IsTaxInclusive", "Project", "ReferenceNo","Ref", "Ref2", "Ref3", "Ref4", "Ref5", "DateRef1", "DateRef2", "NumRef1", "NumRef2", "TaxCode", "TariffCode", "TaxRate", "WTaxCode", "WTaxRate", "WVatCode", "WVatRate"]]
    
    return output_tab
 
def generate_si(num):
    # `INV` + 20 digits (num within)
    si_id = "INV"
    
    digits = len(str(num))
    str_num = str(num)
    for i in range(0,20-digits):
        si_id += "0"
    for i in range(0,digits):
        si_id += str_num[i]
    return si_id

def generate_si_batch(start, stop):
    si_s = pd.Series(range(start, stop))
    si_list = []

    for i, v in si_s.items():
        si_list.append(generate_si(v))

    return si_list

def get_last_day(monthStr, year):
    month = month_index.get(monthStr)
        
    last_day = calendar.monthrange(year, month)[1]

    mon_str = str(month)
    if (len(mon_str) == 1):
        mon_str = "0" + mon_str

    date_str = mon_str + "/" + str(last_day) + "/" + str(year)

    return date_str

def isValidInt(text: str):
    try:
        int(text)
    except:
        return False
    return True

# temp = create_template(23062, "January", 2024, "/home/tope/Desktop/SOA Laz Jabra Jan 1-11 2024.xlsx", "/home/tope/Desktop/625Tech SO Jan 1-11 2024.xlsx", "/home/tope/Desktop/625Tech SI Jan 1-11 2024.xlsx", "/home/tope/Desktop/625Tech SOR Jan 1-11 2024.xlsx")
# print(temp.shape)
