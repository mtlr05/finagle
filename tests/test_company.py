import finagle as cmp
import pytest
import pandas as pd
import os
import copy

def test_value():
    #Sample Problem 2
    #value straight from fcfe
    #Value=2929.108 - note: the calculation doesn't use the first value of fcfe, since this is supposed to be a ttm number.

    #initilizers
    re = 0.0557+0.0214
    gt = 0.0214
    year = 6 #number of years to forescast; i.e. not including ttm (baseline) year

    #bypass financials
    financials ={'date' : '2021-12-31'}
    fcfe=[0, 155.76,161.20,166.84,172.67,178.71,182.53]
    abc = cmp.company(financials = financials, ticker = 'abc', year = year, fcfe = fcfe, re = re, gt = gt)
    
    result = pd.DataFrame(abc.value())
    
    filename=os.path.join(os.path.dirname(__file__), 'value.pkl')
    #answer = pd.read_pickle("./value.pkl")
    answer=pd.read_pickle(filename)
    os.remove('abc.log')
    pd.testing.assert_frame_equal(answer, result)

def test_fcf_from_earnings():
    #Sample Problem 1
    #value calculated using fcf_from_earnings()
    #PV = 8.7

    #initializers
    re = 0.12
    gt = 0.01
    shares = 1
    year = 5 #number of years to forescast; i.e. not including ttm (baseline) year

    #company input data
    financials ={
    'date' : '2021-12-31',
    'e': 1,    
    }

    #method specific inputs
    payout = 0.5
    gf = [0.05,0.04]
    roe = 0.15

    xyz = cmp.company(financials = financials, ticker = 'xyz',re = re, gt = gt, year = year, shares = shares)
    xyz.fcf_from_earnings(payout,gf,roe)
    
    result = pd.DataFrame(xyz.value())
    
    filename=os.path.join(os.path.dirname(__file__), 'fcf_from_earnings.pkl')
    answer=pd.read_pickle(filename)
    #answer = pd.read_pickle("./fcf_from_earnings.pkl")
    os.remove('xyz.log')
    pd.testing.assert_frame_equal(answer, result)

def test_fcf_from_ebitda():
    #Sample Problem 3
    #full ebitda based calculation
    #includes captial allocation to debt and buybacks
    #PV = 74.0

    #initializers
    rd = 0.065
    re = 0.10
    t = 0.21
    shares = 2.3 
    gt = 0.02
    roict=0.75 #this could be quite high if there is significant uncapitalized R&D
    year = 10 #number of years to forescast; i.e. not including ttm (baseline) year

    #company input data
    financials = {
    'date' : '2021-12-31',
    'revenue' :[0],
    'ebitda' : [13.7,13.8,14.0,14.7,16.0,18.1,19.6,21.4,23.7,26.4,29.7],
    'capex' :  [1.4,1.5,1.6,1.6,1.8,2.0,2.0,2.1,2.2,2.4,2.5],
    'sbc' : [0,0,0,0,0,0,0,0,0,0,0],    
    'dwc' : [0,0,0,0,0,0,0,0,0,0,0],
    'tax' : [0.68],
    'da' : [6.8,7.0,7.3,8.0,9.0], #don't input for all years since the terminal year should be calculated from capex and ROIC
    'debt' :  [51.7,43.3,34.6,36.6,40.0,45.2,48.9,53.6,59.3,74.3,80], #the terminal year should not increase from the prior year

    'interest' : [3.6],
    'cash' : 0,
    'nol' : 0,
    'noa' : 0,
    }

    DISCK = cmp.company(financials = financials,ticker = 'DISCK',rd = rd,re = re,t = t,shares = shares,gt = gt,roict = roict,year = year)
    DISCK.fcf_from_ebitda()
    DISCK.fcf_to_debt(leverage=2.5)
    DISCK.fcf_to_buyback(price=28.22,dp = 'proportional')

    result = pd.DataFrame(DISCK.value())
    filename=os.path.join(os.path.dirname(__file__), 'fcf_from_ebitda.pkl')
    answer=pd.read_pickle(filename)
    #answer = pd.read_pickle("./fcf_from_ebitda.pkl")
    os.remove('DISCK.log')
    pd.testing.assert_frame_equal(answer, result)
    
def test_forecast_ebitda():
    #Sample Problem 4
    #FRG
    #full ebitda based calculation, using forecast_ebitda() and load_data() method
    #PV = 114

    #initializers
    rd = 0.065
    re = 0.10
    t = 0.21
    shares = 0.744 + 40.295 
    gt = 0.02
    roict=0.15 #this could be quite high if there is significant uncapitalized R&D
    year = 10 #number of years to forescast; i.e. not including ttm (baseline) year

    #company input data
    financials = {
    'date' : '2021-12-26',
    'revenue' :[0],
    'ebitda' : [330], #only the ttm year, other years are calculated using the forecast_ebitda() method
    'capex' :  [40,46.0,50.6,55.7,61.2,67.3,73.2,78.6,83.3,87.2,90.1], #year+1 required
    'dwc' : [0,0,0,0,0,0,0,0,0,0,0], #year+1 required
    'sbc' : [0,0,0,0,0,0,0,0,0,0,0],    
    'tax' : [0],
    'da' : [51.847,54.4], #don't input for all years since the terminal year should be calculated from capex and ROIC
    'debt' :  [1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9], 

    'interest' : [70],
    'cash' : 159.72, 
    'nol' : 127.4,
    'noa' : 55.86, #0 if nothing, this is for the Nexpoint equity
    }

    FRG = cmp.company(ticker = 'FRG',rd = rd,re = re,t = t,shares = shares,gt = gt,roict = roict,year = year)
    FRG.forecast_ebitda(330,[0.15,0.1,0.1,0.1,0.1], financials)
    FRG.load_financials(financials = financials.copy())
    FRG.fcf_from_ebitda()
    FRG.fcf_to_debt(leverage=2.5)
    FRG.fcf_to_buyback(price=43,dp = 'proportional')

    result = pd.DataFrame(FRG.value())
    filename=os.path.join(os.path.dirname(__file__), 'forecast_ebitda.pkl')
    answer=pd.read_pickle(filename)
    #answer = pd.read_pickle("./forecast_ebitda.pkl")
    os.remove('FRG.log')
    pd.testing.assert_frame_equal(answer, result)

def test_forecast_ebitda_keeps_supplied_revenue_and_interpolates_on_ebitda():
    year = 8
    me = 0.25
    mc = 0.40
    gsnext = 0.10
    gf = [0.10, 0.10, 0.05, 0.05]
    length = len(gf)
    financials = {
        'revenue': [400, 450],
        'ebitda': [100],
    }
    snapshot = copy.deepcopy(financials)
    company = cmp.company(
        ticker='REV', rd=0.05, re=0.1, t=0.21, shares=10,
        gt=0.02, roict=0.15, year=year, dividend=[0])
    ebitda = company.forecast_ebitda(
        100, gf, financials, me=me, mc=mc, gsnext=gsnext)
    assert financials == snapshot

    revenue = company._forecasts['revenue']
    assert len(revenue) == year + 1
    assert revenue[0] == 400
    assert revenue[1] == 450
    assert all(value is not None for value in revenue)

    # gs[length] is gsnext: it is the sales-growth knot at the end of gf.
    first = length + 1
    assert revenue[first] == pytest.approx(ebitda[length] / me * (1 + gsnext))

    span = ebitda[first] - ebitda[1]
    for i in range(2, first):
        expected = 450 + (revenue[first] - 450) * (ebitda[i] - ebitda[1]) / span
        assert revenue[i] == pytest.approx(expected)

    os.remove('REV.log')

def test_fcf_to_acquire():
    #Sample 5
    #FRG
    #full ebitda based calculation, using forecast_ebitda() and load_data() method
    #includes fcf_to_acquire()
    #pv=126.4

    #initializers
    rd = 0.065
    re = 0.10
    t = 0.21
    shares = 0.744 + 40.295 
    gt = 0.02
    roict=0.15 #this could be quite high if there is significant uncapitalized R&D
    year = 10 #number of years to forescast; i.e. not including ttm (baseline) year

    #company input data
    financials = {
    'date' : '2021-12-26',
    'revenue' :[0],
    'ebitda' : [330], #only the ttm year, other years are calculated using the forecast_ebitda() method
    'capex' :  [40,46.0,50.6,55.7,61.2,67.3,73.2,78.6,83.3,87.2,90.1], #year+1 required
    'dwc' : [0,0,0,0,0,0,0,0,0,0,0], #year+1 required
    'sbc' : [0,0,0,0,0,0,0,0,0,0,0], #year+1 required
    'tax' : [0],
    'da' : [51.847,54.4], #don't input for all years since the terminal year should be calculated from capex and ROIC
    'debt' :  [1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9,1072.9], 

    'interest' : [70],
    'cash' : 159.72, 
    'nol' : 127.4,
    'noa' : 55.86, #None if nothing, this is for the Nexpoint equity
    }

    FRG = cmp.company(ticker = 'FRG',rd = rd,re = re,t = t,shares = shares,gt = gt,roict = roict,year = year)
    FRG.forecast_ebitda(330,[0.15,0.1,0.1,0.1,0.1], financials)
    FRG.load_financials(financials = financials.copy())
    FRG.fcf_from_ebitda()
    FRG.fcf_to_acquire(adjust_cash = True, year_a = 0, ebitda_frac = 0.2696, multiple = 6.52, leverage = 6.52, gnext = 0.1, cap_frac = 0.22)
    FRG.fcf_to_debt(leverage=2.5)
    FRG.fcf_to_buyback(price=45,dp = 'proportional')

    result = pd.DataFrame(FRG.value())
    filename=os.path.join(os.path.dirname(__file__), 'fcf_to_acquire.pkl')
    answer=pd.read_pickle(filename)
    #answer = pd.read_pickle("./fcf_to_acquire.pkl")
    os.remove('FRG.log')
    pd.testing.assert_frame_equal(answer, result)
    
def test_fcf_to_allocate():
    # Sample 6
    # PL, post-q2 2023 (quarter ending july 31,2022)
    # full ebitda based calculation, using forecast_ebitda() and load_data() method
    # testing fcf_to_allocate
    # PV = 26.1, PV_DDM = 21.6

    #initializers
    rd = 0.075 #no debt
    re = 0.10
    t = 0.21
    shares = 270 + 12.8 + 21.459 + 8 # 270 shares + 12.8 warrant+ 21.459 earn-out shares not included in sbc + 8 estimated increase in H2
    gt = 0.02
    roict=0.15 #this could be quite high if there is significant uncapitalized R&D
    year = 15 #number of years to forescast; i.e. not including ttm (baseline) year

    #company input data
    financials = {
    'date' : '2023-01-31',
    'revenue' :[0],
    'ebitda' :[-64,-35,42,162,242,354,502,694,933,1219,1548,1907,2278,2637,2956,3205], #only the ttm year, other years are calculated using the forecast_ebitda() method
    'capex' :[26,38,50,61,87,120,161,210,268,332,400,468,531,584,622,641], #year+1 required
    'dwc' :[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], #year+1 required
    'sbc' :[100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100],
    'tax' :[0],
    'da' : [60], #don't input for all years since the terminal year should be calculated from capex and ROIC
    'debt' :[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], 
        
    'interest' : [0],
    'cash' : 598 - 45 , #includes $450 + $148 from warrants - $45 for H2 burn-rate 
    'nol' : 114.6,
    'noa' : 0, 
    }

    PL = cmp.company(ticker = 'PL',rd = rd,re = re,t = t,shares = shares,gt = gt,roict = roict,year = year, dividend = 0)
    PL.forecast_sbc(sbc_f=[80,80,80,80,80],sbc_rate_t = 0.17,financials=financials)
    PL.load_financials(financials = financials.copy())
    PL.fcf_from_ebitda()
    PL.fcf_to_debt(leverage=1, year_d=4)
    PL.fcf_to_allocate(price = 10,dp = 'constant', buybacks = 0)
    
    result = pd.DataFrame(PL.value())
    filename=os.path.join(os.path.dirname(__file__), 'fcf_to_allocate.pkl')
    answer=pd.read_pickle(filename)
    #answer = pd.read_pickle("./fcf_to_acquire.pkl")
    os.remove('PL.log')
    pd.testing.assert_frame_equal(answer, result)

def test_forecast_leaves_caller_dict_and_log_and_rejects_missing_column():
    year = 10
    financials = {
        'date': '2021-12-26',
        'revenue': [0],
        'ebitda': [330],
        'capex': [40, 46.0],
        'dwc': [0] * (year + 1),
        'sbc': [0],
        'tax': [0],
        'da': [51.847],
        'debt': [10] * (year + 1),
        'interest': [1],
        'cash': 5,
        'nol': 0,
        'noa': 0,
    }
    snapshot = copy.deepcopy(financials)
    company = cmp.company(
        ticker='LOGME', rd=0.05, re=0.1, t=0.21, shares=10,
        gt=0.02, roict=0.15, year=year, dividend=[0])
    company.forecast_ebitda(330, [0.1], financials)
    company.forecast_capex(financials['capex'], financials)
    company.forecast_sbc(financials['sbc'], financials, sbc_rate_t=None)
    assert financials == snapshot

    with_note = copy.deepcopy(financials)
    with_note['notes'] = 'ignored extra key'
    company.load_financials(financials=with_note)
    assert company.data_for_ebitda is True
    company.fcf_from_ebitda()
    company.value()
    with open('LOGME.log', encoding='utf-8') as handle:
        text = handle.read()
    os.remove('LOGME.log')
    assert 'LOGME' in text
    assert 'fcf_from_ebitda() method complete' in text

    missing = copy.deepcopy(snapshot)
    del missing['sbc']
    incomplete = cmp.company(
        ticker='BADCO', rd=0.05, re=0.1, t=0.21, gt=0.02, year=year)
    with pytest.raises(ValueError, match='sbc'):
        incomplete.load_financials(missing)
    os.remove('BADCO.log')


def _statements(year=4):
    return {
        'date': '2021-12-31',
        'revenue': [200, 220, 240, 260, 280],
        'ebitda': [100, 110, 120, 130, 140],
        'capex': [20],
        'sbc': [5, 5, 5, 5, 5],
        'dwc': [0, 0, 0, 0, 1],
        'tax': [1],
        'da': [10, 12],
        'debt': [10] * (year + 1),
        'interest': [1],
        'cash': 0,
        'nol': 0,
        'noa': 0,
    }


def _capex_model(ticker, year=4, gt=0.02, roict=0.20, t=0.25):
    return cmp.company(
        ticker=ticker, rd=0.05, re=0.10, t=t, shares=1,
        gt=gt, roict=roict, year=year, dividend=[0])


def _read_log(ticker):
    with open(ticker + '.log', encoding='utf-8') as handle:
        return handle.read()


def test_method_1_solves_terminal_capex_and_rounds_trip_depreciation():
    year = 4
    financials = _statements(year)
    snapshot = copy.deepcopy(financials)
    company = _capex_model('M1', year)
    capex = company.forecast_capex(financials['capex'], financials)
    assert financials == snapshot

    C = 0.02 / 0.20 * (1 - 0.25)
    da_t = 12 / 220 * 280
    capex_t = da_t * (1 - C) - 1 + C * (140 - 5)
    revenue = financials['revenue']
    expected = [20]
    for i in range(1, year + 1):
        expected.append(20 + (capex_t - 20) * (revenue[i] - revenue[0]) / (revenue[year] - revenue[0]))
    assert capex == pytest.approx(expected)
    assert capex[-1] == pytest.approx(capex_t)

    company.load_financials(financials.copy())
    company.fcf_from_ebitda()
    da = list(company.fin['da'])
    assert da[0] == pytest.approx(10)
    assert da[1] == pytest.approx(12)
    for i in range(1, year + 1):
        assert da[i] == pytest.approx(12 / 220 * revenue[i])
    os.remove('M1.log')


def test_method_2_holds_the_capex_to_ebitda_ratio_when_asked():
    year = 4
    financials = _statements(year)
    financials['capex'] = [20, 22]
    company = _capex_model('M2', year)
    capex = company.forecast_capex(financials['capex'], financials, method=2)
    ebitda = financials['ebitda']
    assert capex[0:2] == [20, 22]
    for i in range(2, year + 1):
        assert capex[i] == pytest.approx(22 / ebitda[1] * ebitda[i])
    os.remove('M2.log')


def test_capex_method_defaults_from_revenue():
    year = 4
    with_revenue = _statements(year)
    company = _capex_model('DEF1', year)
    default = company.forecast_capex(with_revenue['capex'], with_revenue)
    explicit = _capex_model('DEF1B', year).forecast_capex(
        with_revenue['capex'], with_revenue, method=1)
    assert default == pytest.approx(explicit)

    bare = _statements(year)
    bare['revenue'] = [0]
    bare['ebitda'] = [100, 110, 121, 133.1, 146.41]
    bare['capex'] = [40, 46]
    company = _capex_model('DEF2', year)
    company.forecast_ebitda(100, [0.1], bare)
    capex = company.forecast_capex(bare['capex'], bare)
    ebitda = company._forecasts['ebitda']
    assert capex[0:2] == [40, 46]
    for i in range(2, year + 1):
        assert capex[i] == pytest.approx(46 / ebitda[1] * ebitda[i])
    os.remove('DEF1.log')
    os.remove('DEF1B.log')
    os.remove('DEF2.log')


def test_method_1_without_revenue_does_not_store_capex():
    year = 4
    financials = _statements(year)
    financials['revenue'] = [0]
    company = _capex_model('NOREV', year)
    company.forecast_ebitda(100, [0.1], financials)
    assert company.forecast_capex(financials['capex'], financials, method=1) is None
    assert 'capex' not in company._forecasts
    assert 'method 1 needs revenue' in _read_log('NOREV')
    os.remove('NOREV.log')


def test_flat_ebitda_depreciation_falls_back_to_a_time_line():
    year = 2
    financials = {
        'date': '2021-12-31',
        'revenue': [0],
        'ebitda': [100],
        'capex': [10, 10, 10],
        'dwc': [0, 0, 0],
        'sbc': [0, 0, 0],
        'tax': [0],
        'da': [20],
        'debt': [0, 0, 0],
        'interest': [0],
        'cash': 0,
        'nol': 0,
        'noa': 0,
    }
    company = _capex_model('FLAT', year, gt=0, roict=0.15, t=0.21)
    company.forecast_ebitda(100, [0], financials)
    company.load_financials(financials.copy())
    company.fcf_from_ebitda()
    da = list(company.fin['da'])
    assert da[0] == pytest.approx(20)
    assert da[1] == pytest.approx((da[0] + da[2]) / 2)
    assert 'EBITDA does not change between the last supplied depreciation' in _read_log('FLAT')
    os.remove('FLAT.log')


def test_negative_terminal_capex_is_a_warning():
    year = 2
    financials = _statements(year)
    financials['revenue'] = [100, 110, 120]
    financials['ebitda'] = [50, 55, 60]
    financials['sbc'] = [0, 0, 0]
    financials['da'] = [1]
    financials['dwc'] = [0, 0, 80]
    financials['capex'] = [5]
    company = _capex_model('NEGC', year)
    capex = company.forecast_capex(financials['capex'], financials, method=1)
    assert capex[-1] < 0
    assert 'negative capex in terminal year' in _read_log('NEGC')
    os.remove('NEGC.log')


def test_terminal_sbc_uses_the_stored_forecast():
    year = 4
    financials = _statements(year)
    financials['sbc'] = [5]
    first = _capex_model('SBC1', year)
    projected = first.forecast_capex(financials['capex'], financials, method=1)
    assert 'forecast_sbc has not run' in _read_log('SBC1')

    second = _capex_model('SBC2', year)
    second.forecast_sbc(financials['sbc'], financials, sbc_rate_t=0.2)
    stated = second.forecast_capex(financials['capex'], financials, method=1)
    assert projected[-1] != pytest.approx(stated[-1])
    assert 'forecast_sbc has not run' not in _read_log('SBC2')
    os.remove('SBC1.log')
    os.remove('SBC2.log')