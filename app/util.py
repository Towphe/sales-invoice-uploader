import pandas as pd
import numpy as np
import calendar

class TaxInfo():
    def __init__(self, sales_person: str, debtor: str):
        self.sales_person = sales_person
        self.debtor = debtor

tax_index = {
    "LAZADA JABRA" : TaxInfo("LAZADA JABRA", "102-J001"),
    "LAZADA LG" : TaxInfo("LAZADA LG", "102-L001")
}

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

def create_template(invoice_start:int, month:str, year:str, soa_dir:str, so_dir:str, si_dir:str, sor_dir:str, sales_person:str):
    
    # get soa file
    soa_df = pd.read_excel(soa_dir)
    soa_df[['Order No.']] = soa_df[['Order No.']].astype(str)

    # filter soa to only rows with column `Transaction Type` to `Order Sales` or `Refund Claims`
    try:
        soa_df = soa_df[(soa_df["Transaction Type"] == "Orders-Sales") | (["Transaction Type"] == "Refunds-Claims")]
    except:
        return ("Error in SOA file", False)

    # open si file
    si_df = pd.read_excel(si_dir)

    try:
        # filter soa to entries with NO entry yet in SI file
        #soa_df = soa_df[~(soa_df["Order No."].isin(si_df["Reference No"]))]
        soa_df = soa_df[~(soa_df["Order No."].isin(si_df["Reference No"]))]
    except:
        return ("Error in SI file", False)
    
    print(f"filtered soa: {soa_df.shape[0]}")
    
    # get so file
    so_df = pd.read_excel(so_dir)

    # format types
    
    so_df[['Reference No']] = so_df[['Reference No']].astype(str)

    try:
        # match soa and so
        matched_soa_so = pd.merge(soa_df, so_df, how='left', left_on='Order No.', right_on='Reference No')
    except:
        return ("Error in SO file", False)
    
    matched_soa_so_grouped = pd.DataFrame()
    so_soa_agg = {'Transaction Date' : 'first', 'Reference 1' : 'first', 'Reference 2': 'first', 'Reference 3' : 'first', 'Reference 4' : 'first', 'Reference 5' : 'first', 'Amount':'first',}

    # try:
    #     # aggregate according to order no., s. order #, and sku
    #     matched_soa_so_grouped = matched_soa_so.groupby(['Order No.', 'S. Order #', 'Seller SKU'], as_index=False).agg(so_soa_agg)
    # except:
    #     return ("Error in Aggregation", False)
    matched_soa_so_grouped = matched_soa_so.groupby(['Order No.', 'S. Order #', 'Seller SKU'], as_index=False).agg(so_soa_agg)
    
    print(f"aggregated soa & so: {matched_soa_so_grouped.shape[0]}")

    # get soreg file
    soreg_df = pd.read_excel(sor_dir)
    
    # merge previously grouped soa_so with soreg
    matched_soa_so_sor = pd.DataFrame()
    try:
        # merge previously grouped soa_so with soreg
        matched_soa_so_sor = pd.merge(matched_soa_so_grouped, soreg_df, how='left', left_on="S. Order #", right_on='SO #')
    except:
        return ("Error in SOR file", False)
    
    # fix formatting and add necessary columns
    final_agg = {
        'Amount_x' : 'first',
        'Transaction Date' : 'first',
        'Reference 1' : 'first',
        'Reference 2' : 'first',
        'Reference 3' : 'first',
        'Reference 4' : 'first',
        'Reference 5' : 'first',
        'Date' : 'first',
        'SO #' : 'first',
        'Stock #' : 'first',
        'Description' : 'first',
        'Qty' : 'first',
        'UOM' : 'first',
        'Amount_y' : 'first',
        'Tax' : 'first',
        'Net' : 'first'
    }
    
    matched_soa_so_sor = matched_soa_so_sor.groupby(['Order No.', 'S. Order #', 'Seller SKU'], as_index=False).agg(final_agg)

    print(f"matched_soa_so_sor count: {matched_soa_so_sor.shape[0]}")

    last_day = get_last_day(month, year)

    matched_soa_so_sor["SalesInvoiceCode"] = ""
    matched_soa_so_sor["SalesInvoiceDate"] = last_day # get invoice date
    matched_soa_so_sor["PostingDate"] = last_day # get posting date
    matched_soa_so_sor["OurDONO"] = ""
    matched_soa_so_sor["DueDate"] = last_day # set as end of month
    matched_soa_so_sor["GlobalProgressInvoicingRate"] = ""
    matched_soa_so_sor["IsApproved"] = True    
    matched_soa_so_sor["IsDeferredVAT"] = False
    matched_soa_so_sor["TaxDate"] = last_day
    matched_soa_so_sor["Debtor"] = tax_index[sales_person].debtor
    matched_soa_so_sor["CurrencyRate"] = "N8" # allow customization
    matched_soa_so_sor["ReverseRate"] = "N8" # allow customization
    matched_soa_so_sor["SalesPerson"] = tax_index[sales_person].sales_person
    matched_soa_so_sor["Term"] = "C.O.D."
    matched_soa_so_sor["Remark1"] = matched_soa_so_sor["SO #"]
    matched_soa_so_sor["Remark2"] = ""
    matched_soa_so_sor["Remark3"] = ""
    matched_soa_so_sor["Remark4"] = ""
    matched_soa_so_sor["Remark5"] = ""
    matched_soa_so_sor["Project"] = ""
    matched_soa_so_sor["StockLocation"] = ""
    matched_soa_so_sor["DORegistationNo"] = ""
    matched_soa_so_sor["DOArea"] = ""
    matched_soa_so_sor["CostCentre"] = ""
    matched_soa_so_sor["IsCancelled"] = ""
    matched_soa_so_sor["IsTaxInclusive"] = True
    matched_soa_so_sor["IsRounding"] = ""
    matched_soa_so_sor["IsNonTaxInvoice"] = ""
    matched_soa_so_sor["none"] = ""
    matched_soa_so_sor["ProgressInvoicingRate"] = ""
    matched_soa_so_sor["StockType"] = ""
    matched_soa_so_sor["SerialNumber"] = ""
    matched_soa_so_sor["StockBatchNumber"] = ""
    matched_soa_so_sor["DebtorItem"] = ""
    matched_soa_so_sor["ServiceCost"] = ""
    matched_soa_so_sor["PackingUOM"] = ""
    matched_soa_so_sor["Packing"] = ""
    matched_soa_so_sor["PackingQty"] = ""
    matched_soa_so_sor["Numbering"] = ""
    matched_soa_so_sor["Stock"] = matched_soa_so_sor["Stock #"]
    matched_soa_so_sor["StockLocation"] = ""
    # ...
    matched_soa_so_sor["UnitPrice"] = pd.to_numeric(matched_soa_so_sor["Amount_x"])
    matched_soa_so_sor["Qty"] = pd.to_numeric(matched_soa_so_sor["Qty"]) 
    matched_soa_so_sor["Discount"] = ""
    matched_soa_so_sor["GLAccount"] = ""
    matched_soa_so_sor["CostCentre"] = ""
    matched_soa_so_sor["ReferenceNo"] = ""
    matched_soa_so_sor["Ref"] = ""
    matched_soa_so_sor["Ref2"] = ""
    matched_soa_so_sor["Ref3"] = ""
    matched_soa_so_sor["Ref4"] = ""
    matched_soa_so_sor["Ref5"] = ""
    matched_soa_so_sor["DateRef1"] = ""
    matched_soa_so_sor["DateRef2"] = ""
    matched_soa_so_sor["NumRef1"] = ""
    matched_soa_so_sor["NumRef2"] = ""
    matched_soa_so_sor["TaxCode"] = "SR-SP"
    matched_soa_so_sor["TariffCode"] = ""
    matched_soa_so_sor["TaxRate"] = "12.00%"
    matched_soa_so_sor["WTaxCode"] = ""
    matched_soa_so_sor["WTaxRate"] = "0.00"
    matched_soa_so_sor["WVatCode"] = "0.00"
    matched_soa_so_sor["WVatRate"] = ""

    matched_soa_so_sor.to_excel("test-output.xlsx")

    output_tab = matched_soa_so_sor[["SalesInvoiceCode", "SalesInvoiceDate", "PostingDate", "OurDONO", "DueDate", "GlobalProgressInvoicingRate", "IsApproved", "IsDeferredVAT", "TaxDate", "Debtor", "CurrencyRate", "ReverseRate", "SalesPerson", "Term", "Order No.", "Reference 1", "Reference 2", "Reference 3", "Reference 4", "Reference 5", "Remark1", "Remark2", "Remark3", "Remark4", "Remark5", "Project", "StockLocation", "DORegistationNo", "DOArea", "CostCentre", "IsCancelled", "IsTaxInclusive", "IsRounding", "IsNonTaxInvoice", "none", "ProgressInvoicingRate", "StockType", "SerialNumber", "StockBatchNumber", "DebtorItem", "ServiceCost", "PackingUOM", "Packing", "PackingQty", "Numbering", "Stock", "StockLocation", "Qty", "UOM", "UnitPrice", "Discount", "GLAccount", "CostCentre", "Description", "IsTaxInclusive", "Project", "ReferenceNo","Ref", "Ref2", "Ref3", "Ref4", "Ref5", "DateRef1", "DateRef2", "NumRef1", "NumRef2", "TaxCode", "TariffCode", "TaxRate", "WTaxCode", "WTaxRate", "WVatCode", "WVatRate"]]

    return output_tab

files = {
    "soa" : "/home/tope/Desktop/SOA Laz Jabra Jan 1-11 2024.xlsx",
    "so" : "/home/tope/Desktop/625Tech SO Jan 1-11 2024.xlsx",
    "si" : "/home/tope/Desktop/625Tech SI Jan 1-11 2024.xlsx",
    "sor" : "/home/tope/Desktop/625Tech SOR Jan 1-11 2024.xlsx"
}

#output = create_template_v2(23046, "January", 2024, files['soa'], files['so'], files['si'], files['sor'])
