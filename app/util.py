import pandas as pd

def match_soa_and_sorep(soa, sorep):    
    # join two tables together
    sorep_soa_matched = pd.merge(soa, sorep, how="left", left_on="Order No.", right_on="Order No.")
    sorep_soa_matched.to_excel("sorep_soa_matched.xlsx", "Sheet 1", index=False)

    # 'Transaction Type', 'Transaction Number', 'Fee Name', 
    sorep_soa_matched = sorep_soa_matched[['Transaction Date', 'S. Order #', 'Order No.', 'Reference 1', 'Reference 2', 'Reference 3', 'Reference 4', 'Reference 5', 'Amount_y']]
    
    agg_funcs = {'Transaction Date' : 'first', 'S. Order #' : 'first', 'Order No.' : 'first', 'Reference 1' : 'first', 'Reference 2': 'first', 'Reference 3' : 'first', 'Reference 4' : 'first', 'Reference 5' : 'first', 'Amount_y':'sum',}
    sorep_soa_grouped = sorep_soa_matched.groupby(['Order No.']).agg(agg_funcs)

    return sorep_soa_grouped

def join_soreg_and_sorep_and_soa(groupd_soa_and_sorep, soreg):
    matched_soreg_sorep_soa = pd.merge(groupd_soa_and_sorep, soreg, left_on="S. Order #", right_on="SO #", how="left")
    return matched_soreg_sorep_soa

def check_if_existing_si(matched_soreg_sorep_soa, sirep):
    #soreg_sorep_wo_si = matched_soreg_sorep_soa[matched_soreg_sorep_soa["S. Order #"].isin(sirep["Transfer From"]) == False]
    soreg_sorep_wo_si = matched_soreg_sorep_soa[~matched_soreg_sorep_soa["S. Order #"].isin(sirep["Transfer From"])]
    #print(sirep["Transfer From"])
    #soreg_sorep_wo_si = matched_soreg_sorep_soa
    return soreg_sorep_wo_si    

def create_template(soa_dir:str, sorep_dir:str, sirep_dir:str, soreg_dir:str):
    # get soa data
    soa = pd.read_excel(soa_dir)
    filtered_soa = soa[(soa["Transaction Type"] == "Orders-Sales") | (soa["Transaction Type"] == "Refunds-Claims")]

    # get sorep data
    sorep = pd.read_excel(sorep_dir)
    filtered_sorep = sorep[sorep["Name"] == "JABRA PH - ONLINE SALES"]

    sorep_soa_grouped = match_soa_and_sorep(filtered_soa, filtered_sorep)

    sirep = pd.read_excel(sirep_dir)

    # filter to SI Report later
    # => if order already has SI, disregard
    sorep_soa_grouped = check_if_existing_si(sorep_soa_grouped, sirep)
    if (sorep_soa_grouped.shape[0] == 0):
        # empty dataframe
        return False

    #Batching and joining with SO Register

    soreg = pd.read_excel(soreg_dir)

    matched_soreg_sorep_soa = join_soreg_and_sorep_and_soa(sorep_soa_grouped, soreg)

    #print(output_tab)
    matched_soreg_sorep_soa["SalesInvoiceCode"] = "" # iterate thrhu all later
    matched_soreg_sorep_soa["SalesInvoiceDate"] = "" # get invoice date
    matched_soreg_sorep_soa["PostingDate"] = "12-31-2023" # get posting date
    matched_soreg_sorep_soa["OurDONO"] = ""
    matched_soreg_sorep_soa["DueDate"] = "12-31-2023" # set as end of month
    matched_soreg_sorep_soa["GlobalProgressInvoicingRate"] = ""
    matched_soreg_sorep_soa["IsApproved"] = True    
    matched_soreg_sorep_soa["IsDeferredVAT"] = False
    matched_soreg_sorep_soa["TaxDate"] = "12-31-2023"
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
    matched_soreg_sorep_soa["UnitPrice"] = matched_soreg_sorep_soa["Amount"]
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
    print(output_tab)
    # create file
    #matched_soreg_sorep_soa.to_excel('output.xlsx', sheet_name="Sheet 1", index=False)
    #print(output_tab)
    #output_tab.to_excel('output.xlsx', sheet_name="Sheet 1", index=False)
 
# test data   
# create_template('../../SalesInvoiceUploader-data/test-data/Dec2023 - Lazada SOA.xlsx', '../../SalesInvoiceUploader-data/test-data/Dec2023 - Sales Order Report.xlsx', '../../SalesInvoiceUploader-data/test-data/Dec2023 - Sales Invoice Report.xlsx', '../../SalesInvoiceUploader-data/test-data/Dec2023 - Sales Order Register.xlsx')