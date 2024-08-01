### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.

from .date import getDaysOfMonth


def btround(p_num, p_places=2):

    dec = decimal.Decimal.from_float(float(p_num))
    dec = dec.quantize(
        decimal.Decimal(str(10 ** (-3 - p_places))), rounding=decimal.ROUND_HALF_UP
    )
    dec = dec.quantize(
        decimal.Decimal(str(10 ** (-p_places))), rounding=decimal.ROUND_HALF_UP
    )

    return float(dec)


#


def calcDailyinterest(p_balance, p_intrate, p_startdate, p_enddate):

    if p_intrate <= 0:
        return 0

    if p_balance is None or p_balance < 0:
        p_balance = 0
    if p_startdate is None or p_startdate < 0:
        p_startdate = 0
    if p_enddate is None or p_enddate < 0:
        p_enddate = 0

    daily_int = p_intrate / 365.0 / 100
    daily_periods = p_enddate - p_startdate
    v1 = pow((1 + daily_int), daily_periods)
    balance = p_balance * v1

    return btround(balance - p_balance, 2)


def calcInterest(p_period, p_amount, p_rate, p_usecompound=1):
    interestamount = 0
    if p_rate <= 0:
        return 0
    if p_amount <= 0:
        return 0

    if p_usecompound == 1:
        days = getDaysOfMonth(p_period)
        interestamount = p_amount * days * (float(p_rate) / 100.0 / 365.0)
    elif p_usecompound == 2:
        days = getDaysOfMonth(p_period)
        interestamount = (
            p_amount * pow(1 + (float(p_rate) / 36500.0), days)
        ) - p_amount
    else:
        interestamount = p_amount * float(p_rate) / 100.0 / 12.0
    #

    return btround(interestamount, 2)


#
