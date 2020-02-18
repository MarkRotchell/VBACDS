import math

from datetime import date
from datetime import datetime as dt
from datetime import timedelta as td
from dateutil.relativedelta import relativedelta as rd
from typing import Callable

SATURDAY = 4
MONTHS_IN_YEAR = 12
ANNUAL = 1
SEMI_ANNUAL = 2
QUARTERLY = 4
MONTHLY = 12

def _fol(start_date: date) -> date:
	assert isinstance(start_date,date)
	while start_date.weekday() > SATURDAY:
		start_date += td(days=1)
	return start_date

def _prev(start_date: date) -> date:
	assert isinstance(start_date,date)
	while start_date.weekday() > SATURDAY:
		start_date += td(days=-1)
	return start_date

def _modfol(start_date: date) -> date:
	assert isinstance(start_date, date)
	f = _fol(start_date)
	if f.month != start_date.month:
		return _prev(start_date)
	else:
		return f

def _act360(start_date: date,
		    end_date:   date) -> float:
	assert isinstance(start_date,date)
	assert isinstance(end_date,date)
	DAYS_IN_YEAR = 360
	return float((end_date - start_date).days / DAYS_IN_YEAR)

def _act365(start_date: date,
		    end_date:   date) -> float:
	assert isinstance(start_date,date)
	assert isinstance(end_date,date)
	DAYS_IN_YEAR = 365
	return float((end_date - start_date).days / DAYS_IN_YEAR)


def _30360(start_date: date,
		   end_date:   date) -> float:
	assert isinstance(start_date,date)
	assert isinstance(end_date,date)
	DAYS_IN_YEAR = 360
	DAYS_IN_MONTH = 30
	d1 = start_date.day
	m1 = start_date.month
	y1 = start_date.year
	d2 = end_date.day
	m2 = end_date.month
	y2 = end_date.year
	if d1 > DAYS_IN_MONTH:
		d1 = DAYS_IN_MONTH
	if d1 == DAYS_IN_MONTH and d2 > DAYS_IN_MONTH:
		d2 = DAYS_IN_MONTH
	return ((y2-y1) * DAYS_IN_YEAR + (m2-m1)*DAYS_IN_MONTH + (d2-d1))/DAYS_IN_YEAR

def _tenor_to_date(start_date: date, 
	  			   tenor: str \
				   ) -> date:
	assert isinstance(start_date,date)
	assert isinstance(tenor,str)
	assert len(tenor) >=2
	periods, period = tenor[:-1], tenor[-1].upper()
	assert period in ["M","Y"], "tenor expect to end with M or Y"
	assert periods.isdecimal()==True
	periods = int(periods)
	if period.upper() == "M":
		return start_date + rd(months = periods)
	elif period.upper() == "Y":
		return start_date + rd(years = periods)

def _tenor_to_months(tenor: str) -> date:
	assert isinstance(tenor,str)
	assert len(tenor) >=2
	periods, period = tenor[:-1], tenor[-1].upper()
	assert period in ["M","Y"], "tenor expect to end with M or Y"
	assert periods.isdecimal()==True
	periods = int(periods)
	if period.upper() == "M":
		return periods
	elif period.upper() == "Y":
		return int(periods * MONTHS_IN_YEAR)

def _generate_schedule_from_start_date(start_date: date,
									   frequency: int, #number of coupons per year
									   months: int,    #months from start date to last coupon
									   bad_day_convention: Callable = _modfol \
									   ) -> list:
	
	assert isinstance(frequency,int)
	assert frequency in [ANNUAL,SEMI_ANNUAL,QUARTERLY,MONTHLY]
	months_in_period = int(MONTHS_IN_YEAR / frequency)

	assert isinstance(months,int)
	assert months > 0
	assert months % months_in_period == 0

	periods = int(months/months_in_period)
	rolls = [start_date + rd(months=i*months_in_period) for i in range(periods+1)]
	return list(map(bad_day_convention,rolls))



def _depo_discount_factor(curve_date: date,
						  maturity: date,
						  rate: float,
						  day_count_convention: Callable) ->float:
	year_fraction = day_count_convention(curve_date,maturity)
	return 1 / (1+rate*year_fraction)

def _depo_point(curve_date: date,
			    tenor: str,
			    rate: float,
			    depo_day_count_convention: Callable,
			    depo_bad_day_convention: Callable, 
				curve_zero_convention: Callable,
				curve_day_count_convention: Callable) -> float:
	maturity = _tenor_to_date(curve_date,tenor)
	maturity = depo_bad_day_convention(maturity)
	discount_factor = _depo_discount_factor(curve_date,maturity,rate,depo_day_count_convention)
	zero = curve_zero_convention(curve_date,maturity,discount_factor,curve_day_count_convention)
	return [maturity,zero]



def _cc_zero_from_discount_factor(curve_date: date, 
								  maturity: date,
								  discount_factor: float,
								  day_count_convention: Callable = _act365) -> float:

	assert isinstance(curve_date,date)
	assert isinstance(maturity, date)
	assert isinstance(discount_factor,float)
	assert isinstance(day_count_convention, Callable)
	assert discount_factor > 0
	return - math.log(discount_factor)/day_count_convention(curve_date,maturity)

def _discount_factor_from_cc_zero(curve_date: date,
								  maturity: date,
								  zero: float,
								  day_count_convention: Callable = _act365) -> float:
	assert isinstance(curve_date,date)
	assert isinstance(maturity, date)
	assert isinstance(zero,float)
	assert isinstance(day_count_convention, Callable)
	return math.exp(-zero * day_count_convention(curve_date,maturity))

def _linear_search_date(target: date,
						dates: list) -> int:
	if target < dates[0]:
		return -1
	elif target >= dates[-1]:
		return len(dates)-1
	else:
		for i in range(1, len(dates)):
			if target < dates[i]:
				return i-1
	
def _interp_zero_constant_forwards(target: date,
								   curve_date: date,
								   curve: list) -> float:
	dates = [d for d,z in curve]
	zeros = [z for d,z in curve]
	i = _linear_search_date(target,dates)

	if i == -1: 
		return zeros[0]
	elif target == dates[i]: 
		return zeros[i]
	elif i == len(dates)-1: 
		i -= 1

	t = float((target-curve_date).days)
	t1 = float((dates[i]-curve_date).days)
	t2 = float((dates[i+1]-curve_date).days)
	tz1 = t1 * zeros[i]
	tz2 = t2 * zeros[i+1]

	return (tz1 + (t - t1)/(t2-t1) * (tz2-tz1))/t

def _value_swap_with_proposed_zero_rate(schedule: list,
	 									coupon_rate: float,
			    						curve_date: date,
			    						zero_guess: float, 
			    						curve: list,
			    						curve_day_count_convention: Callable = _act365,
			    						swap_day_count_convention: Callable = _30360,
									    curve_interpolator: Callable = _interp_zero_constant_forwards,
									    curve_zero_to_DF: Callable = _discount_factor_from_cc_zero \
			    						) -> float:
	
	maturity = schedule[-1]
	curve.append([maturity,zero_guess])
	coupon_dates = schedule[1:]
	dccs = [swap_day_count_convention(i,j) for i,j, in zip(schedule,coupon_dates)]
	zeros = [curve_interpolator(d,curve_date,curve) for d in coupon_dates]
	dfs = [curve_zero_to_DF(curve_date,d,z,curve_day_count_convention) for d,z in zip(coupon_dates,zeros)]
	unit_coupon_pvs = [df * dcc for df, dcc in zip(dfs,dccs)]
	curve.pop()
	return sum(unit_coupon_pvs) * coupon_rate + dfs[-1] -1

def _swap_point(curve_date: date,
				tenor: str,
				coupon_rate: float,
				curve: list,
				swap_frequency: int = SEMI_ANNUAL,
				curve_day_count_convention: Callable = _act365,
				curve_interpolator: Callable = _interp_zero_constant_forwards,
				curve_zero_to_DF: Callable = _discount_factor_from_cc_zero,
				curve_DF_to_zero: Callable = _cc_zero_from_discount_factor,
				swap_day_count_convention: Callable = _30360,
				swap_bad_day_convention: Callable = _modfol) -> float:
	
	SECOND_GUESS_MULTIPLIER = 1.1
	EPSILON = 1E-15

	months = _tenor_to_months(tenor)
	schedule = _generate_schedule_from_start_date(curve_date,swap_frequency,months,swap_bad_day_convention)
	z1 = coupon_rate
	z2 = coupon_rate * SECOND_GUESS_MULTIPLIER
	p1 = _value_swap_with_proposed_zero_rate(schedule,coupon_rate,curve_date,z1,curve,curve_day_count_convention,swap_day_count_convention,curve_interpolator,_discount_factor_from_cc_zero)
	p2 = 0
	while math.fabs((z2-z1)/z1) > EPSILON:
		p2 = _value_swap_with_proposed_zero_rate(schedule,coupon_rate,curve_date,z2,curve,curve_day_count_convention,swap_day_count_convention,curve_interpolator,_discount_factor_from_cc_zero)
		if (p2 - p1) == 0.0: break
		z = z1 + (z2-z1)*(0-p1)/(p2-p1)
		z1 = z2
		z2 = z
		p1 = p2

	return [schedule[-1],z2]

curve_date = date(2009,9,9)
rawdata = [
    ("D","1M",0.2538),
    ("D","2M",0.2650),
    ("D","3M",0.3144),
    ("D","6M",0.7125),
    ("D","9M",1.0263),
    ("D","1Y",1.2881),
    ("S","2Y",1.2957),
    ("S","3Y",1.9209),
	("S","4Y",2.3832),
	("S","5Y",2.7488),
	("S","6Y",3.0217),
	("S","7Y",3.2484),
	("S","8Y",3.4039),
	("S","9Y",3.544),
	("S","10Y",3.6495),
	("S","12Y",3.8139),
	("S","15Y",3.9686),
	("S","20Y",4.0715),
	("S","25Y",4.1134),
	("S","30Y",4.1459),
    ]


CURVE_DCC = _act365
SWAP_DCC =  _30360
DEPO_DCC =  _act360

DEPO_BDC = _modfol
SWAP_BDC = _modfol

CURVE_INTERPOLATOR = _interp_zero_constant_forwards
CURVE_DF_TO_ZERO = _cc_zero_from_discount_factor
CURVE_ZERO_TO_DF = _discount_factor_from_cc_zero

curve_zero_convention = _cc_zero_from_discount_factor

curve = []

for product, tenor, rate in rawdata:
	if product.upper() == "D":
		curve.append(_depo_point(curve_date,
								 tenor,
								 rate/100,
								 DEPO_DCC,
								 DEPO_BDC,
								 CURVE_DF_TO_ZERO,
								 CURVE_DCC))
	elif product.upper() =="S":
		curve.append(_swap_point(curve_date,
						         tenor,
								 rate/100,
								 curve,
								 SEMI_ANNUAL,
								 CURVE_DCC,
								 CURVE_INTERPOLATOR,
								 CURVE_ZERO_TO_DF,
								 CURVE_DF_TO_ZERO,
								 SWAP_DCC,
								 SWAP_BDC))

for d,z in curve: print(d.strftime("%d%b%y"),z)

#dates = [d for d,z in curve]
#zeros = [z for d,z in curve]

#print(dates)

#plt.plot(dates,zeros)
#plt.show()
