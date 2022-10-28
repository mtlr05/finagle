import finagle as cmp
import pytest
import pandas as pd
import os

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
    answer = pd.read_pickle("./fcf_from_earnings.pkl")
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
    answer = pd.read_pickle("./fcf_from_ebitda.pkl")
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
    answer = pd.read_pickle("./forecast_ebitda.pkl")
    os.remove('FRG.log')
    pd.testing.assert_frame_equal(answer, result)

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
    answer = pd.read_pickle("./fcf_to_acquire.pkl")
    os.remove('FRG.log')
    pd.testing.assert_frame_equal(answer, result)