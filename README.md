# Company project

This project contains a class, a notebook and an excel template. This is a class is used to model the financials of a publicly traded company. The notebook provides a convinient way to use the class and has various examples which may be modified as a starting point. The excel spreasheet is a template file which is used for creating a report for documentation and display purposes.   

## Company Class

The python code 'company.py' contains a which allows modelling a company financially. With financial forecasts in various forms or levels of precision, one can input company financial foecasts. These can then form the basis for the calculation of free-cash flows (FCFE or FCFF, collectively referred to here FCF's). Once the FCF's have been calculated using the appropriate method, they can then be used for various purposes, specifically they can pay down debt, buyback shares, or even make an acquisition.

It has a number of methods which in general should be run in order once initialization of the company has occured:
- `forecast_X()`, to forecast financial data, where a forecast was not included in the financial dictionary.
- `load_financials()`, to load financial data into the company if it was not included during initialization.
- `fcf_from_X()`, there are several of these methods. They are all meant to be used for calculating fcf
- `fcf_to_X()`, these are methods which are used to model how the FCF is to be used: paying down debt, buying back shares
- `value()`, this method is use for calculating a DCF from the cashflows, This shold be run only once other methods described have been run
- `display_fin()`, this method is used to process the financials. It should be used once all modelling is completed

You may refer to the docstring in the class for more details on its use or the notebook for examples on its us

## Valuation notebook

The jupyter notebook 'Valuation notebook.ipynb' is a convenient way of using the company class. It contains various examples which can be modified for your own use. The examples make it fairly intuitive (I think), and can be self-interpreted so that you can use the class with little or no explanation.

## Company template 

The company template is a file which is called with the `xyz.display_fin()` method which is used to create a spreadsheet report

## Example

```
import company as cmp

#initializers
rd = 0.045
re = 0.10
t = 0.21
shares = 46 
gt = 0.02
roict=0.15 #mostly used for calculating terminal D&A
year = 10 #number of years to forescast; i.e. not including ttm (baseline) year

#company input data
financials = {
'date' : '2021-9-30',
'ebitda' : [881], #898 - stock based comp. only the ttm year, other years are calculated using the forecast_ebitda() method
'capex' :  [64,85],
'dwc' : [0,0,0,0,0,0,0,0,0,0,0],
'tax' : [192],
'da' : [79], #don't input for all years since the terminal year should be calculated from capex and ROIC
'debt' :  [758,758,758,758,765,765,765,765,765,765,765], 

'interest' : [33],
'cash' : 576,
'nol' : 0,
'noa' : 0,
}

ATKR = cmp.company(ticker = 'ATKR',rd = rd,re = re,t = t,shares = shares,gt = gt,roict = roict,year = year)
ATKR.forecast_ebitda(881,[-0.19,-0.15,0.10,0.10], financials) #only input 'organic' EBITDA
ATKR.forecast_capex(financials['capex'],financials)
ATKR.load_financials(financials = financials.copy())
ATKR.fcf_from_ebitda()
ATKR.fcf_to_acquire(year_a = 0, ebitda_frac = 0.008, multiple = 6.5, leverage = 0, gnext = 0.1, cap_frac = 0.12, adjust_cash = False)
ATKR.fcf_to_acquire(year_a = 1, ebitda_frac = 0.03, multiple = 6.5, leverage = 0, gnext = 0.1, cap_frac = 0.12, adjust_cash = False)
ATKR.fcf_to_acquire(year_a = 2, ebitda_frac = 0.1, multiple = 6.5, leverage = 0, gnext = 0.1, cap_frac = 0.12, adjust_cash = False)
ATKR.fcf_to_debt(leverage=2, year_d=3)
ATKR.fcf_to_buyback(price=120,dp = 'proportional')
ATKR.value()
ATKR.display_fin()
```


