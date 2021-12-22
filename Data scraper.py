from selenium import webdriver
import time
import warnings
import statistics
import os.path
from os import path
warnings.filterwarnings('ignore', category=DeprecationWarning)

# excel writer
fname = 'name.csv'
fileexits = 1
if not path.exists(fname):
    fileexits = 0
f = open(fname, 'a', encoding='utf-8')
if fileexits == 0:
    f.write('TICKER,ROC,R.RATE,C.DEBT,WKCAP%,D.RATIO')

# opens tickers for iteration
with open('industry.txt') as ticker:
    tickers = ticker.readline()
ticker.close()



"""CONSTANT INPUTS"""
rf = 3.195
cds = 1.68

"""TICKER LOOPING PROGRAM"""

drvPath = r'+ path + \chromedriver.exe'
browser = webdriver.Chrome(drvPath)

list_tickers = []
[list_tickers.append(x) for x in tickers.split()]

for test in list_tickers:

    """INCOME STATEMENTS"""

    browser.get('http://financials.morningstar.com/income-statement/is.html?t=' + test + '&region=phl&culture=en-US')
    time.sleep(10)

    # revenue
    revenue = []
    for reve in range(1, 6):
        revenue.append(browser.find_element_by_xpath("//div[@id='data_i1']//div[@id='Y_" + str(reve) + "']").text)
    if revenue[0] == '':
        cnv_revenue = [0, 0, 0, 0, 0]
    else:
        revenue = [w.replace('—', '0') for w in revenue]
        revenue = [w.replace('(', '-') for w in revenue]
        revenue = [w.replace(')', '') for w in revenue]
        x11 = []
        [x11.append(i.translate({ord(','): None})) for i in revenue]
        cnv_revenue = list(map(int, x11))

    incb = [] #income before tax
    for a in range(1, 6):
        incb.append(browser.find_element_by_xpath("//div[@id='data_i60']//div[@id='Y_" + str(a) + "']").text)
    if incb[0] == '':
        cnv_incb = [0, 0, 0, 0, 0]
    else:
        incb = [w.replace('—', '0') for w in incb]
        incb = [w.replace('(', '-') for w in incb]
        incb = [w.replace(')', '') for w in incb]
        x11 = []
        [x11.append(i.translate({ord(','): None})) for i in incb]
        cnv_incb = list(map(int, x11))

    prv = []
    for g in range(1, 6):
        prv.append(browser.find_element_by_xpath("//div[@id='data_i61']//div[@id='Y_" + str(g) + "']").text)
    if prv[0] == '':
        cnv_prv = [0, 0, 0, 0, 0]
    else:
        prv = [w.replace('—', '0') for w in prv]
        prv = [w.replace('(', '-') for w in prv]
        prv = [w.replace(')', '') for w in prv]
        x11 = []
        [x11.append(i.translate({ord(','): None})) for i in prv]
        cnv_prv = list(map(int, x11))

    tax = cnv_prv[-1]/cnv_incb[-1]

    """BALANCE SHEET"""
    browser.get('http://financials.morningstar.com/balance-sheet/bs.html?t=' + test + '&region=phl&culture=en-US')
    time.sleep(10)

    stdebt = []
    for st in range(1, 6):
        stdebt.append(browser.find_element_by_xpath("//div[@id='data_i41']//div[@id='Y_" + str(st) + "']").text)
    if stdebt[0] == '':
        cnv_stdebt = [0, 0, 0, 0, 0]
    else:
        stdebt = [w.replace('—', '0') for w in stdebt]
        x5 = []
        [x5.append(i.translate({ord(','): None})) for i in stdebt]
        cnv_stdebt = list(map(int, x5))

    ltdebt = []
    for lt in range(1, 6):
        ltdebt.append(browser.find_element_by_xpath("//div[@id='data_i50']//div[@id='Y_"+str(lt)+"']").text)
    if ltdebt[0] == '':
        cnv_ltdebt = [0, 0, 0, 0, 0]
    else:
        ltdebt = [w.replace('—', '0') for w in ltdebt]
        x6 = []
        [x6.append(i.translate({ord(','): None})) for i in ltdebt]
        cnv_ltdebt = list(map(int, x6))

    equity = []
    for x in range(1, 6):
        equity.append(browser.find_element_by_xpath("//div[@id='data_ttg8']//div[@id='Y_" + str(x) + "']").text)
    if equity == '':
        cnv_equity = [0, 0, 0, 0, 0]
    else:
        equity = [w.replace('—', '0') for w in equity]
        equity = [w.replace('(', '-') for w in equity]
        equity = [w.replace(')', '') for w in equity]
        x3 = []
        [x3.append(i.translate({ord(','): None})) for i in equity]
        cnv_equity = list(map(int, x3))


    cash = []
    for x in range(1, 6):
        cash.append(browser.find_element_by_xpath("//div[@id='data_ttgg1']//div[@id='Y_" + str(x) + "']").text)
    if cash == '':
        cnv_cash = [0, 0, 0, 0, 0]
    else:
        cash = [w.replace('—', '0') for w in cash]
        cash = [w.replace('(', '-') for w in cash]
        cash = [w.replace(')', '') for w in cash]
        x3 = []
        [x3.append(i.translate({ord(','): None})) for i in cash]
        cnv_cash = list(map(int, x3))

    c_assets = []
    for cass in range(1, 6):
        c_assets.append(browser.find_element_by_xpath("//div[@id='data_ttg1']//div[@id='Y_" + str(cass) + "']").text)
    if c_assets[0] == '':
        cnv_cassets = [0, 0, 0, 0, 0]
    else:
        c_assets = [w.replace('—', '0') for w in c_assets]
        x1 = []
        [x1.append(i.translate({ord(','): None})) for i in c_assets]
        cnv_cassets = list(map(int, x1))

    c_lia = []
    for cl in range(1, 6):
        c_lia.append(browser.find_element_by_xpath("//div[@id='data_ttgg5']//div[@id='Y_"+str(cl)+"']").text)
    if c_lia[0] == '':
        cnv_clia = [0, 0, 0, 0, 0]
    else:
        c_lia = [w.replace('—', '0') for w in c_lia]
        c_lia = [w.replace('(', '-') for w in c_lia]
        c_lia = [w.replace(')', '') for w in c_lia]
        x2 = []
        [x2.append(i.translate({ord(','): None})) for i in c_lia]
        cnv_clia = list(map(int, x2))

    tot_debt = cnv_stdebt[-1] + cnv_ltdebt[-1]
    nwkcap = cnv_cassets[-1] - cnv_clia[-1]

    """CASH FLOW"""
    browser.get('http://financials.morningstar.com/cash-flow/cf.html?t=' + test + '&region=phl&culture=en-US')
    time.sleep(10)
    net_income = []
    for net_i in range(1, 6):  # (1, 6) is the year from 1-4
        net_income.append(browser.find_element_by_xpath(
            "//div[@id='data_i1']//div[@id='Y_" + str(net_i) + "']").text)  # str(net_i) loops 4yrs
    if net_income[0] == '':
        cnv_net_income = [0, 0, 0, 0, 0]
    else:
        net_income = [w.replace('—', '0') for w in net_income]
        net_income = [w.replace('(', '-') for w in net_income]
        net_income = [w.replace(')', '') for w in net_income]
        x1 = []
        [x1.append(i.translate({ord(','): None})) for i in net_income]
        cnv_net_income = list(map(int, x1))

    wkcap = []
    for x in range(1, 6):  # (1, 6) is the year from 1-4
        wkcap.append(browser.find_element_by_xpath(
            "//div[@id='data_i15']//div[@id='Y_" + str(x) + "']").text)
    if wkcap[0] == '':
        cnv_wkcap = [0, 0, 0, 0, 0]
    else:
        wkcap = [w.replace('—', '0') for w in wkcap]
        wkcap = [w.replace('(', '') for w in wkcap]
        wkcap = [w.replace(')', '') for w in wkcap]
        x4 = []
        [x4.append(i.translate({ord(','): None})) for i in wkcap]
        cnv_wkcap = list(map(int, x4))

    capex = []
    for net_i in range(1, 6):  # (1, 6) is the year from 1-4
        capex.append(browser.find_element_by_xpath(
            "//div[@id='data_i96']//div[@id='Y_" + str(net_i) + "']").text)  # str(net_i) loops 4yrs
    if capex[0] == '':
        cnv_capex= [0, 0, 0, 0, 0]
    else:
        capex = [w.replace('—', '0') for w in capex]
        capex = [w.replace('(', '') for w in capex]
        capex = [w.replace(')', '') for w in capex]
        x4 = []
        [x4.append(i.translate({ord(','): None})) for i in capex]
        cnv_capex = list(map(int, x4))

    da = []
    for d in range(1, 6):
        da.append(browser.find_element_by_xpath("//div[@id='data_i2']//div[@id='Y_"+str(d)+"']").text)
    if da[0] == '':
        cnv_da = [0, 0, 0, 0, 0]
    else:
        da = [w.replace('—', '0') for w in da]
        da = [w.replace('(', '-') for w in da]
        da = [w.replace(')', '') for w in da]
        x15 = []
        [x15.append(i.translate({ord(','): None})) for i in da]
        cnv_da = list(map(int, x15))

    netcapex = cnv_capex[-1] - cnv_da[-1]
    ebit_tax = cnv_net_income[-1] * (1-tax)

    """KEY RATIOS"""
    browser.get('http://financials.morningstar.com/ratios/r.html?t=' + test + '&region=phl&culture=en-US')
    time.sleep(10)

    intcov = []
    for yr in range(10, 11):
        intcov.append(browser.find_element_by_xpath(
            "/html[1]/body[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div[4]/table[2]/tbody[1]/tr[16]/td[" + str(yr) + "]").text)
    intcov = [w.replace('—', '0') for w in intcov]
    cnv_intcov = [float(i) for i in intcov]
    if cnv_intcov[-1] >= 12.5:
        cove = 0.69
    elif 9.5 <= cnv_intcov[-1] < 12.5:
        cove = 0.85
    elif 7.5 <= cnv_intcov[-1] < 9.5:
        cove = 1.07
    elif 6 <= cnv_intcov[-1] < 7.5:
        cove = 1.18
    elif 4.5 <= cnv_intcov[-1] < 6:
        cove = 1.33
    elif 4 <= cnv_intcov[-1] < 4.5:
        cove = 1.71
    elif 3.5 <= cnv_intcov[-1] < 4:
        cove = 2.31
    elif 3 <= cnv_intcov[-1] < 3.5:
        cove = 2.77
    elif 2.5 <= cnv_intcov[-1] < 3:
        cove = 4.05
    elif 2 <= cnv_intcov[-1] < 2.5:
        cove = 4.86
    elif 1.5 <= cnv_intcov[-1] < 2:
        cove = 5.94
    elif 1.25 <= cnv_intcov[-1] < 1.5:
        cove = 9.47
    elif 0.8 <= cnv_intcov[-1] < 1.25:
        cove = 9.97
    elif 0.5 <= cnv_intcov[-1] < 0.8:
        cove = 13.09
    elif cnv_intcov[-1] < 0.5:
        cove = 17.44
    else:
        cove = 0

    """CALC"""
    try:
        roc = str(ebit_tax/(tot_debt + (cnv_equity[-1] - cnv_cash[-1])) * 100)
        rrate = str((netcapex + cnv_wkcap[-1])/ebit_tax * 100)
        wkrev = str(nwkcap/cnv_revenue[-1] * 100)
        de = str(tot_debt/cnv_equity[-1] * 100)
        cdebt = str(rf + cds + cove)

        f.write('\n' + test + ',' + roc + ',' + rrate + ',' + cdebt + ',' + wkrev + ',' + de + ',')
    except ZeroDivisionError:
        continue

browser.quit()