
import requests,json,sys,os
import pandas as pd
from openpyxl import load_workbook

def keyloader():
    if(os.path.isfile('.key')):
        with open('.key','r') as keyer:
            key = str(keyer.readline()).replace('\n','').split(',')[1]
            return key
    else:
        key = str(input("Please enter your EDGAR API key: ")).replace(' ','')
        with open('.key','w') as keyer:
            print('key,'+key,file=keyer)
        return key

def fetch_financials(ticker):
    ticker = ticker.upper()

    # Grab latest financial results from EDGAR
    key = keyloader()
    api_url = 'http://edgaronline.api.mashery.com/v2/corefinancials/ann?primarysymbols='+ticker+'&appkey='
    response = json.loads(requests.get(api_url+key).text)
    latest_results = response["result"]["rows"][0]["values"]

    # Turn list of field / value dictionaries into single dict
    financials = {}
    for item in latest_results:
        financials[item["field"]] = item["value"]

    # Create dataframe
    data = {
        'labels' : [
            "Statement Date",
            "Total Revenue",
            "Gross Profit",
            "General Expenses",
            "EBIT",
            "Depreciation Amortization",
            "Total Current Assets",
            "Capital Expenditures",
            "Fed State Tax",
            "Return On Equity"
        ],
        'data' : [
            financials["periodenddate"],
            financials["totalrevenue"],
            financials["grossprofit"],
            financials["sellinggeneraladministrativeexpenses"],
            financials["ebit"],
            financials["cfdepreciationamortization"],
            financials["totalcurrentassets"],
            financials["capitalexpenditures"],
            ((financials["incomebeforetaxes"] - financials["netincome"]) / financials["netincome"]),
            financials["netincome"] / financials["totalstockholdersequity"]
        ],
        'data_scale' : [
            (financials["periodenddate"]),
            financials["totalrevenue"]/1000,
            financials["grossprofit"]/1000,
            financials["sellinggeneraladministrativeexpenses"]/1000,
            financials["ebit"]/1000,
            financials["cfdepreciationamortization"]/1000,
            financials["totalcurrentassets"]/1000,
            financials["capitalexpenditures"]/1000,
            ((financials["incomebeforetaxes"] - financials["netincome"]) / financials["netincome"]),
            financials["netincome"] / financials["totalstockholdersequity"]
        ]
    }
    dataframe = pd.DataFrame(data=data)

    return dataframe

def write_vars(data,path_to_sheet="Valuation.xlsx",sheet="10K-Data"):
    book = load_workbook(path_to_sheet)
    writer = pd.ExcelWriter(path_to_sheet, engine='openpyxl')

    writer.book = book
    workbook = writer.book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    data.to_excel(writer,sheet,index=False)
    writer.save()

    return 0

write_vars(fetch_financials("glw"),"Valuation.xlsx")

def main():
    args = sys.argv[1:]
    fin_data = fetch_financials(str(args[0]))
    if(len(args) == 1): write_vars(fin_data)
    elif(len(args) == 2): write_vars(fin_data,args[1])
    elif(len(args) == 3): write_vars(fin_data,args[1],args[2])
    else: print("Error...\nPlease enter the ticker, followed by the Excel workbook path (option) and the sheet name for the data set (optional)")

if __name__ == '__main__':
    main()
