import requests
import math
import yfinance as yf
import matplotlib.pyplot as plt
from datetime import datetime
from datetime import timedelta
import numpy as np
import xlwt
from xlwt import Workbook


apiToken = 'PRODUCTION_API_KEY'
sandBoxToken = 'Tsk_201939d7f09b46d0a0c06fd8e2b80337'

#Returns Stock Beta Value as float
def get_beta(stock):
    data = requests.get('https://cloud.iexapis.com/stable/data-points/'+stock+'/BETA?token='+apiToken).json()
    return data

def get_Price(stockSymbol):
    data = requests.get('https://sandbox.iexapis.com/stable/stock/'+stockSymbol+'/price?token='+sandBoxToken).json()
    return data

def get_sector_performance(sector):
    data = requests.get('https://cloud.iexapis.com/stable/stock/market/sector-performance?token='+apiToken).json()

    for sec in data:
        if(sec['name'] == sector):
            perf = sec['performance']
    return perf

def get_price_to_book_ratio(stock):
    data = requests.get('https://cloud.iexapis.com/stable/time-series/REPORTED_FINANCIALS/'+stock+'?token='+apiToken).json()
    assets = data[0]['Assets']
    close = requests.get('https://cloud.iexapis.com/stable/data-points/'+stock+'/QUOTE-CLOSE?token='+apiToken).json()
    sharesOutstanding = data[0]["CommonStockSharesOutstanding"]
    ratio = (close * sharesOutstanding)/assets
    return ratio

def get_historical_close_prices(stock, startDate, endDate, interval):
    data = yf.download(stock,startDate,endDate, interval=interval)
    plt.xlabel("Date")
    plt.ylabel("USD")
    plt.title(stock)
    return data.Close

def get_trenches(stock,startDate,endDate,interval):
    closingPrices = get_historical_close_prices(stock,startDate,endDate,interval)
    currentPrice = 0
    previousPrice = 0
    lowestPrice = 0
    trenches = []
    i = 0
    while(i < closingPrices.size):
        currentPrice = closingPrices[i]
        if(i == 0):
            lowestPrice = currentPrice
        else:
            previousPrice = closingPrices[i-1]
            if(previousPrice < currentPrice):
                trenches.append(previousPrice)
        i+=1
    return trenches

def get_average_price(stock,startDate,endDate,interval):
    data = get_historical_close_prices(stock, startDate, endDate, interval)
    i = 0
    sum = 0.0
    while(i < data.size):
        sum += data[i]
        i+=1
    return sum/data.size

def get_Gross_Profit_Margin_Percentage(stock):
    stockData = stock[0]['income'][0]
    stockGrossIncome = stockData['grossProfit']
    stockTotalRevenue = stockData['totalRevenue']
    if (stockGrossIncome is None):
        return 'Stock Gross Income None Reported'
    elif (stockTotalRevenue is None):
        return 'Stock Total Revenue None Reported'
    else:
        grossProfitMargin = stockGrossIncome / stockTotalRevenue
        return grossProfitMargin * 100

def get_Operation_Expenses_as_Percentage_of_Gross_Profit(stock):
    stockData = stock[0]['income'][0]
    stockOpExpenses = stockData['operatingExpense']
    stockGrossProfit = stockData['grossProfit']
    if(stockOpExpenses is None):
        return 'Stock Operation Expenses None Reported'
    elif(stockGrossProfit is None):
        return 'Stock Gross Profit None Reported'
    else:
        opVSGrossProfit = stockOpExpenses/stockGrossProfit
        return opVSGrossProfit * 100

def get_SGA_Expenses_as_Percentage_of_Gross_Profit(stock):
    stockData = stock[0]['income'][0]
    stockSGAExpenses = stockData['sellingGeneralAndAdmin']
    stockGrossProfit = stockData['grossProfit']
    if(stockSGAExpenses is None):
        return 'Stock SGA Expenses None Reported'
    elif(stockGrossProfit is None):
        return 'Stock Gross Profit None Reported'
    else:
        SGAVSGrossProfit = stockSGAExpenses/stockGrossProfit
        return SGAVSGrossProfit * 100

def get_RD_Expenses_as_Percentage_of_Gross_Profit(stock):
    stockData = stock[0]['income'][0]
    stockRDExpenses = stockData['researchAndDevelopment']
    stockGrossProfit = stockData['grossProfit']
    if(stockRDExpenses is None):
        return 'Stock R&D Expenses None Reported'
    elif(stockGrossProfit is None):
        return 'Stock Gross Profit None Reported'
    else:
        opVSGrossProfit = stockRDExpenses/stockGrossProfit
        return opVSGrossProfit * 100

def get_Operating_Profit_Margin_Percentage(stock):
    stockData = stock[0]['income'][0]
    stockEBIT = stockData['ebit']
    stockRevenue = stockData['totalRevenue']
    if (stockEBIT is None):
        return 'Stock R&D Expenses None Reported'
    elif (stockRevenue is None):
        return 'Stock Gross Profit None Reported'
    else:
        operatingProfitMargin = stockEBIT / stockRevenue
        return operatingProfitMargin * 100

def get_Interest_Income_as_Percentage_of_Operating_Income(stock):
    stockData = stock[0]['income'][0]
    stockInterestIncome = stockData['interestIncome']
    stockOpIncome = stockData['operatingIncome']
    if(stockInterestIncome is None):
        return 'Stock Interest Income None Reported'
    elif(stockOpIncome is None):
        return 'Stock Operating Income None Reported'
    else:
        interestVSOpIncome = stockInterestIncome/stockOpIncome
        return interestVSOpIncome * 100

def get_Pretax_Income_in_Millions(stock):
    stockData = stock[0]['income'][0]
    return stockData['pretaxIncome']/1000000

def get_taxed_Income_in_Millions(stock):
    stockData = stock[0]['income'][0]
    return (stockData['pretaxIncome'] - stockData['incomeTax'])/1000000

def is_Tax_35Percent_of_Income(stock):
    stockData = stock[0]['income'][0]
    if(stockData['incomeTax'] == (0.35 * stockData['pretaxIncome'])):
        return True
    else:
        return False
def get_Tax_Percent_of_Income(stock):
    stockData = stock[0]['income'][0]
    return (stockData['incomeTax']/stockData['pretaxIncome'])*100

def get_Net_Earnings_as_Percentage_of_Revenue(stock):
    stockData = stock[0]['income'][0]
    netEarning = stockData['netIncome']
    revenue = stockData['totalRevenue']
    percent = netEarning/revenue
    return percent * 100

def get_perShare_Earnings_Current(stock):
    stockData = stock[0]['income'][0]
    netEarning = stockData['netIncomeBasic']
    test = stock[3][0]
    try:
        sharesOutstanding = test['CommonStockSharesOutstanding']
    except(KeyError):
        try:
            sharesOutstanding = test['EntityCommonStockSharesOutstanding']
        except(KeyError):
            return 'Invalid Key'
    return netEarning/sharesOutstanding

def get_Current_Assets_in_Millions(stock):
    stockData = stock[1]['balancesheet'][0]
    return (stockData['currentAssets']/1000000)

def get_Current_Cash_in_Millions(stock):
    stockData = stock[1]['balancesheet'][0]
    return (stockData['currentCash']/1000000)

def get_Return_on_Assets_Percent(stock):
    netIncome = stock[4]['financials'][0]['netIncome']
    totalAssets = stock[1]['balancesheet'][0]['totalAssets']
    return (netIncome / totalAssets) * 100

def get_Total_Assets_in_Millions(stock):
    return stock[1]['balancesheet'][0]['totalAssets']/1000000

def get_ShortTerm_Debt_versus_LongTerm_Debt(stock):
    short = stock[4]['financials'][0]['shortTermDebt']
    long = stock[4]['financials'][0]['longTermDebt']
    return short/long

def get_LongTerm_Debt_as_Percentage_of_NetIncome(stock):
    stockData = stock[0]['income'][0]
    netEarning = stockData['netIncome']
    long = stock[4]['financials'][0]['longTermDebt']
    return (long/netEarning)*100

def get_Time_to_Payoff_LongTerm_Debt_with_NetIncome(stock):
    debtToIncome = get_LongTerm_Debt_as_Percentage_of_NetIncome(stock)
    years = (int)(debtToIncome/100)
    months = (int)(((debtToIncome%100)/100)*12)
    return str(years) +' Years, '+ str(months) + ' Months'

#total liabilities/shareholders equity
def get_Debt_to_Shareholders_Equity(stock):
    totalLiabilities = stock[1]['balancesheet'][0]['totalLiabilities']
    shareHolderEquity = stock[1]['balancesheet'][0]['shareholderEquity']
    return (totalLiabilities/shareHolderEquity)
'''
    total liabilities/shareholders equity+treasuryStock
        financialInstitutions should be below 10
        Everything else below 0.80
'''
def get_Adjusted_Debt_to_Shareholders_Equity(stock):
    totalLiabilities = stock[1]['balancesheet'][0]['totalLiabilities']
    shareHolderEquity = stock[1]['balancesheet'][0]['shareholderEquity']
    treasuryStock =  stock[1]['balancesheet'][0]['treasuryStock']
    if(treasuryStock is None):
        return (totalLiabilities/shareHolderEquity)
    else:
        return (totalLiabilities/(shareHolderEquity+treasuryStock))

def get_Retained_Earnings_in_Millions(stock):
    return (stock[1]['balancesheet'][0]['retainedEarnings'])/1000000

def generate_Excel_Report(stocks, symbols, reportName):
    wb = Workbook()
    report = wb.add_sheet('Stock Analysis')
    bold = xlwt.easyxf('font: bold 1')
    report.write(0,0, 'Company', bold)
    report.write(0, 1, 'Stock Price', bold)
    report.write(0, 2, 'Income After Tax ($ in Millions)', bold)
    report.write(0, 3, 'Tax % on Income', bold)
    report.write(0, 4, 'Net Earnings as % of Revenue', bold)
    report.write(0, 5, 'Operating Profit Margin %', bold)
    report.write(0, 6, 'Time to Pay Off Debt w/ Net Income', bold)
    report.write(0, 7, 'Assets ($ in Millions)', bold)
    report.write(0, 8, '% Return on Assets', bold)
    report.write(0, 9, 'Per Share Earnings', bold)
    report.write(0, 10, 'Retained Earnings ($ in Millions)', bold)
    report.write(0, 11, 'Debt to Shareholders Equity (adjusted, $ in Dollars)', bold)
    index = 1
    for stock in stocks:
        report.write(index,0, symbols[(index -1)])
        report.write(index,1, get_Price(symbols[index-1]))
        report.write(index,2, get_taxed_Income_in_Millions(stock))
        report.write(index,3, get_Tax_Percent_of_Income(stock))
        report.write(index,4, get_Net_Earnings_as_Percentage_of_Revenue(stock))
        report.write(index,5, get_Operating_Profit_Margin_Percentage(stock))
        report.write(index,6, get_Time_to_Payoff_LongTerm_Debt_with_NetIncome(stock))
        report.write(index,7, get_Total_Assets_in_Millions(stock))
        report.write(index,8, get_Return_on_Assets_Percent(stock))
        report.write(index,9, get_perShare_Earnings_Current(stock))
        report.write(index,10, get_Retained_Earnings_in_Millions(stock))
        report.write(index,11,get_Debt_to_Shareholders_Equity(stock))
        index+=1
    wb.save(reportName)


def get_Stock_Information_Current_Annually(stock):
    stockData = []
    balanceSheet = requests.get('https://sandbox.iexapis.com/stable/stock/'+stock+'/balance-sheet?period=annual&token='+sandBoxToken).json()
    cashFlow = requests.get('https://sandbox.iexapis.com/stable/stock/'+stock+'/cash-flow?period=annual&token='+sandBoxToken).json()
    incomeStatement = requests.get('https://sandbox.iexapis.com/stable/stock/'+stock+'/income?period=annual&token='+sandBoxToken).json()
    financialsReported = requests.get('https://sandbox.iexapis.com/stable/time-series/REPORTED_FINANCIALS/'+stock+'/10-K?token='+sandBoxToken).json()
    advancedStats = requests.get('https://sandbox.iexapis.com/stable/stock/'+stock+'/advanced-stats?token='+sandBoxToken).json()
    financials = requests.get('https://sandbox.iexapis.com/stable/stock/'+stock+'/financials?period=annual&token='+sandBoxToken).json()
    stockData.append(incomeStatement)
    stockData.append(balanceSheet)
    stockData.append(cashFlow)
    stockData.append(financialsReported)
    stockData.append(financials)
    stockData.append(advancedStats)
    return stockData


todayDate = datetime.date(datetime.now())
yesterday = todayDate - timedelta(days=1)
weekAgo = todayDate - timedelta(days=7)
monthAgo = todayDate - timedelta(days=30)
interval = '1m'

exxon = 'XOM'
chevron = 'CVX'
devon = 'DVN'

exxonData = get_Stock_Information_Current_Annually(exxon)
chevronData = get_Stock_Information_Current_Annually(chevron)
devonData = get_Stock_Information_Current_Annually(devon)

symbols = [exxon, chevron, devon]
data = [exxonData, chevronData, devonData]

generate_Excel_Report(data,symbols,'Test Report 4')




