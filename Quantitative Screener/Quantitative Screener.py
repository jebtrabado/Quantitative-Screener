from selenium import webdriver
import time
import warnings
import statistics
import os.path
from os import path
warnings.filterwarnings('ignore', category=DeprecationWarning)

start = time.time()

# excel writer
fname = 'Quantitative Screener.csv'
fileexits = 1
if not path.exists(fname):
    fileexits = 0
f = open(fname, 'a', encoding='utf-8')
if fileexits == 0:
    f.write('TICKER,STA,RANK,SNOA,RANK,,DSOI,GMI,AQI,SGI,DEPI,SGAI,LVGI,TATA,'
            'PROBM,RANK,,ROA,VALUE,ROIC,VALUE,CFOA,RANK,CAGR,RANK,STAB.,RANK,MAX,,FS_ROA,'
            'FS_FCFTA,ACCRUAL,FS_LEVER,FS_LIQUID,FS_NEQISS,FS_ROA,FS_FCFTA,FS_MGN,FS_TURN,FS_SCORE,RANK')

"""TICKER LOOPING PROGRAM"""

# opens tickers for iteration
with open('tickers.txt') as ticker:
    tickers = ticker.readline()
ticker.close()

# stores original companies for deletion
list_tickers = []
[list_tickers.append(x) for x in tickers.split()]

"""USER INPUT LOOP COUNT"""

# asks for amount of loops
answer = False
user_input = 5
while not answer:
    user = input('Input the number of companies you want to screen for now. ')
    try:
        if int(user) > 0:
            answer = True
            user_input = int(user)
        else:
            print('invalid input')
    except ValueError:
        print('invalid input')

"""OPEN BROWSER"""
drvPath = r'+ CHROME DRIVER FILE PATH + \chromedriver.exe'
browser = webdriver.Chrome(drvPath)


# creates a list to be deleted in the original
company = []
loop = 0
while loop < user_input:
    comp = tickers.split()[loop]
    company.append(comp)
    loop += 1

    """PROGRAM STARTS"""

    for test in company:

        """CASHFLOW STATEMENTS"""

        browser.get('http://financials.morningstar.com/cash-flow/cf.html?t=' + test + '&region=phl&culture=en-US')
        time.sleep(15)

        # net_income
        net_income = []
        for net_i in range(1, 6): # (1, 6) is the year from 1-4
            net_income.append(browser.find_element_by_xpath("//div[@id='data_i1']//div[@id='Y_"+str(net_i)+"']").text) # str(net_i) loops 4yrs
        if net_income[0] == '':
            cnv_net_income = [0, 0, 0, 0, 0]
        else:
            net_income = [w.replace('—', '0') for w in net_income]
            net_income = [w.replace('(', '-') for w in net_income]
            net_income = [w.replace(')', '') for w in net_income]
            x1 = []
            [x1.append(i.translate({ord(','): None})) for i in net_income]
            cnv_net_income = list(map(int, x1))

        # cfos
        cfos = []
        for cfo in range(1, 6):
            cfos.append(browser.find_element_by_xpath("//div[@id='data_tts1']//div[@id='Y_"+str(cfo)+"']").text)
        if cfos[0] == '':
            cnv_cfos = [0, 0, 0, 0, 0]
        else:
            cfos = [w.replace('—', '0') for w in cfos]
            cfos = [w.replace('(', '-') for w in cfos]
            cfos = [w.replace(')', '') for w in cfos]
            x2 = []
            [x2.append(i.translate({ord(','): None})) for i in cfos]
            cnv_cfos = list(map(int, x2))

        # cfis
        cfis = []
        for cfi in range(1, 6):
            cfis.append(browser.find_element_by_xpath("//div[@id='data_tts2']//div[@id='Y_"+str(cfi)+"']").text)
        if cfis[0] == '':
            cnv_cfis = [0, 0, 0, 0, 0]
        else:
            cfis = [w.replace('—', '0') for w in cfis]
            cfis = [w.replace('(', '-') for w in cfis]
            cfis = [w.replace(')', '') for w in cfis]
            x18 = []
            [x18.append(i.translate({ord(','): None})) for i in cfis]
            cnv_cfis = list(map(int, x18))

        # D&A CASHFLOW
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

        """"BALANCE SHEETS"""

        browser.get('http://financials.morningstar.com/balance-sheet/bs.html?t=' + test + '&region=phl&culture=en-US')
        time.sleep(15)

        # tot_assets
        tot_assets = []
        for tassets in range(1, 6):
            tot_assets.append(browser.find_element_by_xpath("//div[@id='data_tts1']//div[@id='Y_"+str(tassets)+"']").text)
        if tot_assets == '':
            cnv_tot = [0, 0, 0, 0, 0]
        else:
            tot_assets = [w.replace('—', '0') for w in tot_assets]
            x3 = []
            [x3.append(i.translate({ord(','): None})) for i in tot_assets]
            cnv_tot = list(map(int, x3))

        # cash equivalence
        cash_equi = []
        for cash in range(1, 6):
            cash_equi.append(browser.find_element_by_xpath("//div[@id='data_i1']//div[@id='Y_"+str(cash)+"']").text)
        if cash_equi == '':
            cnv_equi = [0, 0, 0, 0, 0]
        else:
            cash_equi = [w.replace('—', '0') for w in cash_equi]
            cash_equi = [w.replace('(', '-') for w in cash_equi]
            cash_equi = [w.replace(')', '') for w in cash_equi]
            x4 = []
            [x4.append(i.translate({ord(','): None})) for i in cash_equi]
            cnv_equi = list(map(int, x4))

        # short term debt
        stdebt = []
        for st in range(1, 6):
            stdebt.append(browser.find_element_by_xpath("//div[@id='data_i41']//div[@id='Y_"+str(st)+"']").text)
        if stdebt[0] == '':
            cnv_stdebt = [0, 0, 0, 0, 0]
        else:
            stdebt = [w.replace('—', '0') for w in stdebt]
            x5 = []
            [x5.append(i.translate({ord(','): None})) for i in stdebt]
            cnv_stdebt = list(map(int, x5))

        # long term debt
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

        # minority interest
        min = []
        for mi in range(1, 6):
            min.append(browser.find_element_by_xpath("//div[@id='data_i56']//div[@id='Y_"+str(mi)+"']").text)
        if min[0] == '':
            cnv_min = [0, 0, 0, 0, 0]
        else:
            min = [w.replace('—', '0') for w in min]
            min = [w.replace('(', '-') for w in min]
            min = [w.replace(')', '') for w in min]
            x7 = []
            [x7.append(i.translate({ord(','): None})) for i in min]
            cnv_min = list(map(int, x7))

        # preferred stock
        pref = []
        for pre in range(1, 6):
            pref.append(browser.find_element_by_xpath("//div[@id='data_i81']//div[@id='Y_"+str(pre)+"']").text)
        if pref[0] == '':
            cnv_pref = [0, 0, 0, 0, 0]
        else:
            pref = [w.replace('—', '0') for w in pref]
            x8 = []
            [x8.append(i.translate({ord(','): None})) for i in pref]
            cnv_pref = list(map(int, x8))

        # equity
        equity = []
        for eq in range(1, 6):
            equity.append(browser.find_element_by_xpath("//div[@id='data_ttg8']//div[@id='Y_"+str(eq)+"']").text)

        equity = [w.replace('—', '0') for w in equity]
        x9 = []
        [x9.append(i.translate({ord(','): None})) for i in equity]
        cnv_equity = list(map(int, x9))

        # receivables for day sales outstanding
        resiv = []
        for res in range(1, 6):
            resiv.append(browser.find_element_by_xpath("//div[@id='data_i3']//div[@id='Y_"+str(res)+"']").text)
        if resiv[0] == '':
            cnv_resiv = [0, 0, 0, 0, 0]
        else:
            resiv = [w.replace('—', '0') for w in resiv]
            resiv = [w.replace('(', '-') for w in resiv]
            resiv = [w.replace(')', '') for w in resiv]
            x10 = []
            [x10.append(i.translate({ord(','): None})) for i in resiv]
            cnv_resiv = list(map(int, x10))

        # current assets
        c_assets = []
        for cass in range(1, 6):
            c_assets.append(browser.find_element_by_xpath("//div[@id='data_ttg1']//div[@id='Y_"+str(cass)+"']").text)
        if c_assets[0] == '':
            cnv_cassets = [0, 0, 0, 0, 0]
        else:
            c_assets = [w.replace('—', '0') for w in c_assets]
            x13 = []
            [x13.append(i.translate({ord(','): None})) for i in c_assets]
            cnv_cassets = list(map(int, x13))

        # net ppe
        n_ppe = []
        for nppe in range(1, 6):
            n_ppe.append(browser.find_element_by_xpath("//div[@id='data_ttgg2']//div[@id='Y_"+str(nppe)+"']").text)
        if n_ppe[0] == '':
            cnv_nppe = [0, 0, 0, 0, 0]
        else:
            n_ppe = [w.replace('—', '0') for w in n_ppe]
            x14 = []
            [x14.append(i.translate({ord(','): None})) for i in n_ppe]
            cnv_nppe = list(map(int, x14))

        # current liabilities
        c_lia = []
        for cl in range(1, 6):
            c_lia.append(browser.find_element_by_xpath("//div[@id='data_ttgg5']//div[@id='Y_"+str(cl)+"']").text)
        if c_lia[0] == '':
            cnv_clia = [0, 0, 0, 0, 0]
        else:
            c_lia = [w.replace('—', '0') for w in c_lia]
            c_lia = [w.replace('(', '-') for w in c_lia]
            c_lia = [w.replace(')', '') for w in c_lia]
            x17 = []
            [x17.append(i.translate({ord(','): None})) for i in c_lia]
            cnv_clia = list(map(int, x17))

        """INCOME STATEMENTS"""

        browser.get('http://financials.morningstar.com/income-statement/is.html?t=' + test + '&region=phl&culture=en-US')
        time.sleep(15)

        # revenue
        revenue = []
        for reve in range(1, 6):
            revenue.append(browser.find_element_by_xpath("//div[@id='data_i1']//div[@id='Y_"+str(reve)+"']").text)
        if revenue[0] == '':
            cnv_revenue = [0, 0, 0, 0, 0]
        else:
            revenue = [w.replace('—', '0') for w in revenue]
            revenue = [w.replace('(', '-') for w in revenue]
            revenue = [w.replace(')', '') for w in revenue]
            x11 = []
            [x11.append(i.translate({ord(','): None})) for i in revenue]
            cnv_revenue = list(map(int, x11))

        # cost of revenue
        c_revenue = []
        for creve in range(1, 6):
            c_revenue.append(browser.find_element_by_xpath("//div[@id='data_i6']//div[@id='Y_"+str(creve)+"']").text)
        if c_revenue[0] == '':
            cnv_crevenue = [0, 0, 0, 0, 0]
        else:
            c_revenue = [w.replace('—', '0') for w in c_revenue]
            c_revenue = [w.replace('(', '-') for w in c_revenue]
            c_revenue = [w.replace(')', '') for w in c_revenue]
            x12 = []
            [x12.append(i.translate({ord(','): None})) for i in c_revenue]
            cnv_crevenue = list(map(int, x12))

        # Sales General Administrative
        sga = []
        for sg in range(1, 6):
            sga.append(browser.find_element_by_xpath("//div[@id='data_i12']//div[@id='Y_"+str(sg)+"']").text)
        if sga[0] == '':
            cnv_sga = [0, 0, 0, 0, 0]
        else:
            sga = [w.replace('—', '0') for w in sga]
            sga = [w.replace('(', '-') for w in sga]
            sga = [w.replace(')', '') for w in sga]
            x16 = []
            [x16.append(i.translate({ord(','): None})) for i in sga]
            cnv_sga = list(map(int, x16))

        """CALCULATIONS"""
        """FRAUD AND MANIPULATORS"""


        STA = []
        # STA calculations
        for sta in range(5):
            STA.append(100*((cnv_net_income[sta] - cnv_cfos[sta])/(cnv_tot[sta])))
        sta_output = STA[-1]


        SNOA = []
        for sno in range(5):
            snoa1 = (cnv_tot[sno]-cnv_equi[sno])
            snoa2 = (cnv_tot[sno]-cnv_stdebt[sno]-cnv_ltdebt[sno]-cnv_min[sno]-cnv_pref[sno]-cnv_equity[sno])
            snoa3 = snoa1 - snoa2
            SNOA.append(snoa3/cnv_tot[sno]*100)
        snoa_output = SNOA[-1]

        """PROBM"""

        # days sale outstanding
        daysout = [0]
        for days in range(4): # 4 year from latest
            ave_rec_tn1 = (cnv_resiv[days] + cnv_resiv[days + 1]) / 2 # average receivables from t & t+1
            rec_turnover = cnv_revenue[days + 1] / ave_rec_tn1
            daysout.append(365 / rec_turnover)
        DSO = round((daysout[4] / daysout[3]), 2)
        dso_ouptut = DSO

        # gross margin index
        gm = [] # 5 year results
        for days in range(5):
            gm1 = cnv_revenue[days] - cnv_crevenue[days]
            gm.append(gm1 / cnv_revenue[days] * 100)
        GMI = round((gm[-2] / gm[-1]), 2)
        gmi_output = GMI

        # asset quality index
        aqi = [] # 5 year results
        for days in range(5):
            aqi1 = (cnv_cassets[days] + cnv_nppe[days])
            aqi.append(1 - (aqi1 / cnv_tot[days]))
        AQI = round((aqi[-1] / aqi[-2]), 2)
        aqi_output = AQI

        # sales growth index
        sgi = [0]
        for days in range(1, 5): # gives 4 years sgi
            sgi.append(cnv_revenue[days] / cnv_revenue[days - 1])
        SGI = round(sgi[-1], 2)
        sgi_output = SGI

        # depreciation index
        depi = []
        for days in range(1, 5):
            deno = cnv_da[days] + cnv_nppe[days]
            depi.append(cnv_da[days] / deno)
        DEPI = round((depi[-2] / depi[-1]), 2)
        depi_output = DEPI

        # sales, general and administrative
        sgai = []
        for days in range(1, 5):
            sgai.append(cnv_sga[days] / cnv_revenue[days])
        SGAI = round((sgai[-1] / sgai[-2]), 2)
        sgai_output = SGAI

        # leverage index
        lvgi = []
        for days in range(1, 5):
            num = cnv_ltdebt[days] + cnv_clia[days]
            lvgi.append(num / cnv_tot[days])
        LVGI = round((lvgi[-1] / lvgi[-2]), 2)
        lvgi_output = LVGI

        # otal accruals to total assets
        tata = []
        for days in range(1, 5):
            num = cnv_net_income[days] - cnv_cfos[days] - cnv_cfis[days]
            tata.append(num / cnv_tot[days])
        TATA = round(tata[-1], 2)
        tata_output = TATA

        # PROBM
        MSCORE = -4.84 + (0.92*DSO) + (0.528*GMI) + (0.404*AQI) + (0.892*SGI) + (0.115*DEPI) - (0.172*SGAI) + (4.679*TATA) - (0.327*LVGI)
        mscore_output = MSCORE

        """KEY RATIOS"""

        browser.get('http://financials.morningstar.com/ratios/r.html?t=' + test + '&region=phl&culture=en-US')
        time.sleep(15)

        """FRANCHISE POWER"""

        # 8yr geometric ROA
        ROA = []
        for yr in range(3, 11):
            ROA.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div[4]/table[2]/tbody[1]/tr[8]/td["+str(yr)+"]").text)
        ROA = [w.replace('—', '0') for w in ROA]
        cnv_roa = [float(i) for i in ROA]
        geo_roa = []
        g_mean_roa = 1
        [geo_roa.append(1 + (r / 100)) for r in cnv_roa]
        for x in geo_roa:
            g_mean_roa *= x
        ans_roa = round((100 * ((g_mean_roa ** (1 / len(geo_roa))) - 1)), 2)

        # volatility of roa
        vol_roa = []
        for z in cnv_roa:
            if z != 0:
                k = round((ans_roa / z), 2)
                vol_roa.append(k)
            else:
                vol_roa.append(0)
        vol_counter = 0
        for v in vol_roa:
            if v < 0.5 or v > 1.5:
                vol_counter += 1
        # OUTPUT
        if vol_counter > 5:
            roa_val = 0
        else:
            roa_val = 1

        # 8yr geometric ROIC
        ROIC = []
        for yr in range(3, 11):
            ROIC.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div[4]/table[2]/tbody[1]/tr[14]/td["+str(yr)+"]").text)
        ROIC = [w.replace('—', '0') for w in ROIC]
        cnv_roic = [float(i) for i in ROIC]
        geo_roic = []
        g_mean_roic = 1
        [geo_roic.append(1 + (r / 100)) for r in cnv_roic]
        for x in geo_roic:
            g_mean_roic *= x
        ans_roic = round((100 * ((g_mean_roic ** (1 / len(geo_roic))) - 1)), 2)

        # volatility of roic
        vol_roic = []
        for z in cnv_roic:
            if z != 0:
                k = round((ans_roic / z), 2)
                vol_roic.append(k)
            else:
                vol_roic.append(0)
        vol_counter2 = 0
        for v in vol_roic:
            if v < 0.5 or v > 1.5:
                vol_counter2 += 1
        # OUTPUT
        if vol_counter2 > 5:
            roic_val = 0
        else:
            roic_val = 1

        # cash flow to total assets
        CF = []
        for yr in range(3, 11):
            CF.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[26]/td["+str(yr)+"]").text)
        cfa = []
        CF = [w.replace('—', '0') for w in CF]
        [cfa.append(i.translate({ord(','): None})) for i in CF]
        cnv_cfot = [int(i) for i in cfa]
        sum_cfo = sum(cnv_cfot)
        CFOA = round((sum_cfo / cnv_tot[-1] * 100), 2)
        cfoa_output = CFOA

        # 8yr profit margins
        NM = []
        for yr in range(3, 11):
            NM.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div[4]/table[2]/tbody[1]/tr[4]/td["+str(yr)+"]").text)
        NM = [w.replace('—', '0') for w in NM]
        cnv_nm = [float(i) for i in NM]

        # check if there's too many zero's
        zero = 0
        for check in cnv_nm:
            if check == 0:
                zero += 1
        if zero <= 3:
            loop1 = 0
            while cnv_nm[loop1] <= 4:  # check if oldest margin is zero
                loop1 += 1  # if zero checks the next
            loop2 = -1
            while cnv_nm[loop2] <= 4:  # checks if newest is zero
                loop2 -= 1
            ans_nm = ((cnv_nm[loop2] / cnv_nm[loop1]) ** (1 / (8 - loop1 + loop2)) - 1) * 100

            # stability of margins
            ans_nm_stab = round((statistics.mean(cnv_nm) / statistics.stdev(cnv_nm)), 2)
        else:
            ans_nm = 0
            ans_nm_stab = 100

        # fs_neqiss
        fs_neqiss = []
        for yr in range(3, 11):
            fs_neqiss.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[18]/td["+str(yr)+"]").text)
        fs_neq = []
        [fs_neq.append(i.translate({ord(','): None})) for i in fs_neqiss]
        cnv_neqiss = [int(i) for i in fs_neq]
        fs_neqiss2 = 1 if (cnv_neqiss[-1] - cnv_neqiss[-2]) < 0 else 0

        # gross margin growth
        gmg = []
        for yr in range(3, 11):
            gmg.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div[4]/table[1]/tbody[1]/tr[6]/td["+str(yr)+"]").text)
        gmg = [w.replace('—', '0') for w in gmg]
        cnv_gmg = [float(i) for i in gmg]
        fs_gmg2 = 1 if (cnv_gmg[-1] - cnv_gmg[-2]) > 0 else 0

        # asset turnover growth
        ast = []
        for yr in range(3, 11):
            ast.append(browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[2]/div[2]/div[1]/div[4]/table[2]/tbody[1]/tr[6]/td["+str(yr)+"]").text)
        ast = [w.replace('—', '0') for w in ast]
        cnv_ast = [float(i) for i in ast]
        fs_ast2 = 1 if (cnv_ast[-1] - cnv_ast[-2]) > 0 else 0

        # fs_roa
        fs_roa = 1 if cnv_roa[-1] > 0 else 0
        # fs_fcfta
        fs_fcfta = 1 if (cnv_cfot[-1] / cnv_tot[-1]) > 0 else 0
        # accrual
        fs_acc = 1 if ((cnv_cfot[-1] / cnv_tot[-1]) * 100 - cnv_roa[-1]) > 0 else 0
        # fs_lever
        fs_lever = 1 if ((cnv_ltdebt[-1] / cnv_tot[-1]) - (cnv_ltdebt[-2] / cnv_tot[-2])) > 0 else 0
        # fs_liquid
        fs_liquid = 1 if ((cnv_cassets[-1] / cnv_clia[-1]) - (cnv_cassets[-2] / cnv_clia[-2])) > 0 else 0
        # roa growth
        fs_roa_g = 1 if (cnv_roa[-1] - cnv_roa[-2]) > 0 else 0
        # fcfta growth
        fs_fcfta_g = 1 if ((cnv_cfot[-1] / cnv_tot[-1]) - (cnv_cfot[-2] / cnv_tot[-2])) > 0 else 0

        FS_SCORE = fs_ast2 + fs_gmg2 + fs_fcfta_g + fs_roa + fs_neqiss2 + fs_lever + fs_liquid + fs_fcfta + fs_roa_g + fs_acc

        """EXCEL WRITER"""

        # excel output
        f.write('\n' + test + ',' + str(sta_output) + ',,' + str(snoa_output) + ',,,' + str(dso_ouptut) + ',' + str(gmi_output)
                + ',' + str(aqi_output) + ',' + str(sgi_output) + ',' + str(depi_output) + ',' + str(sgai_output) + ',' + str(lvgi_output)
                + ',' + str(tata_output) + ',' + str(mscore_output) + ',,,' + str(ans_roa) + ',' + str(roa_val) + ',' + str(ans_roic) + ',' + str(roic_val)
                + ',' + str(cfoa_output) + ',,' + str(ans_nm) + ',,' + str(ans_nm_stab) + ',,,,' + str(fs_roa) + ',' + str(fs_fcfta)
                + ',' + str(fs_acc) + ',' + str(fs_lever) + ',' + str(fs_liquid) + ',' + str(fs_neqiss2) + ',' + str(fs_roa_g) + ',' + str(fs_fcfta_g) +
                ',' + str(fs_gmg2) + ',' + str(fs_ast2) + ',' + str(FS_SCORE))

# timer
end = time.time()
print(f"Runtime of the program is {end - start}")

# end of scraping
browser.quit()

for y in company:
    list_tickers.remove(y) if y in list_tickers else None

# gives a new list to be used next
edit = ' '.join(map(str, list_tickers))
edit_txt = open('tickers.txt', 'w')
edit_txt.writelines(edit)
edit_txt.close()