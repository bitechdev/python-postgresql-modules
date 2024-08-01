### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.


import datetime


def clarionDateToPeriod(p_int_date):
    days = 0
    if isinstance(p_int_date, int):
        days = int(p_int_date)
    elif isinstance(p_int_date, str):
        if p_int_date.isnumeric():
            days = p_int_date
        else:
            raise ValueError(
                "The given date parameter does not contain a valid integer."
            )
    else:
        days = int(p_int_date)

    if days <= 0:
        return ""

    if days < 36163 or days > 218784:
        raise ValueError(
            "The given date range is incorrect. Acceptable ranges are 1900-01-01 to 2400-01-01"
        )

    daysfrom = datetime.timedelta(days=days)
    clarionstart = datetime.date(1800, 12, 28)
    return str((clarionstart + daysfrom).isoformat())[:7]


def clarionDateToDateStr(p_int_date):
    days = 0
    if isinstance(p_int_date, int):
        days = p_int_date
    elif isinstance(p_int_date, str):
        if p_int_date.isnumeric():
            days = p_int_date
        else:
            raise ValueError(
                "The given date parameter does not contain a valid integer."
            )
    else:
        raise ValueError("The given date parameter is not an integer.")

    if days < 0 and days > -657432:
        daysfrom = datetime.timedelta(days=abs(days))
        clarionstart = datetime.date(1800, 12, 28)
        return str((clarionstart - daysfrom).isoformat())
    elif days > 0 and days < 437565:
        daysfrom = datetime.timedelta(days=days)
        clarionstart = datetime.date(1800, 12, 28)
        return str((clarionstart + daysfrom).isoformat())
    else:
        return "0000-00-00"


def dateToClarionInt(p_date):
    if p_date is None:
        return 0

    days = 0
    if isinstance(p_date, str):
        date = datetime.datetime.strptime(p_date, "%Y-%m-%d").date()
    elif isinstance(p_date, datetime.date):
        date = p_date
    else:
        raise ValueError(
            "The given date parameter is not an string date or datetime.datetime object. [{}]".format(
                type(p_date)
            )
        )

    clarionstart = datetime.date(1800, 12, 28)
    return int((date - clarionstart).days)


def periodToDate(p_str_period):
    days = 0
    if isinstance(p_str_period, str):
        if len(p_str_period) < 8:
            p_str_period = p_str_period + "-01"
    else:
        raise ValueError("The given period parameter is not a string.")

    if p_str_period == "":
        return 0

    return dateToClarionInt(p_date)


def nextPeriod(p_str_period, p_direction=0):

    if p_str_period is None or p_str_period == "":
        raise ValueError("The period given is blank.")
    if isinstance(p_str_period, str):
        if p_str_period.lower() == "none":
            raise ValueError(
                "A null period parameter was given. [" + str(p_str_period) + "]"
            )
        if len(p_str_period) < 8:
            p_str_period = p_str_period + "-01"
        elif p_str_period[8:10] > "31":
            p_str_period = p_str_period[0:7] + "-01"

    else:
        raise ValueError("The given period parameter is not a string.")

    perioddate = datetime.datetime.strptime(p_str_period, "%Y-%m-%d").replace(day=1)
    newdate = perioddate
    if p_direction > 0:
        newdate = newdate - datetime.timedelta(days=2)
    else:
        newdate = newdate + datetime.timedelta(days=32)

    newdate = newdate.replace(day=1)

    return str(datetime.date.isoformat(newdate))


def addPeriod(p_str_period, p_cnt):
    if isinstance(p_str_period, str):
        if len(p_str_period) < 8:
            p_str_period = p_str_period + "-01"
    else:
        raise ValueError("The given period parameter is not a string.")

    direction = 0
    if p_cnt < 0:
        direction = 1
        p_cnt = abs(p_cnt)

    for i in range(1, p_cnt):
        p_str_period = nextPeriod(p_str_period, direction)

    return p_str_period


def getLastWeekDayOfMonth(p_dayofweeknr, p_period):

    if not isinstance(p_period, str):
        raise ValueError("The given period parameter is not a string.")

    month = int(p_period[5:7])
    if month in (4, 6, 9, 11):
        p_period = p_period[:7] + "-30"
    elif month == 2 and (int(p_period[0:4]) % 4) == 0:
        p_period = p_period[:7] + "-29"
    elif month == 2 and (int(p_period[0:4]) % 4) != 0:
        p_period = p_period[:7] + "-28"
    else:
        p_period = p_period[:7] + "-31"

    perioddate = datetime.datetime.strptime(p_period, "%Y-%m-%d")

    for i in range(1, 8):
        if perioddate.weekday() + 1 == p_dayofweeknr:
            return int(perioddate.day)

        perioddate = perioddate - datetime.timedelta(days=1)

    return 0


def getDaysOfMonth(p_period):
    try:
        if not isinstance(p_period, str):
            raise ValueError("The given period parameter is not a string.")
        #

        month = int(p_period[5:7])
        year = int(p_period[0:4])
        if month in (4, 6, 9, 11):
            return 30
        elif month == 2 and (year % 4) == 0:
            return 29
        elif month == 2 and (year % 4) != 0:
            return 28
        else:
            return 31
        #
    except Exception as e:
        raise ValueError(str(e) + " in getDaysOfMonth Input: p_period=" + str(p_period))
    #

    return 0


#


def getPeriodDate(p_str_period, p_day):

    if isinstance(p_str_period, str):
        if len(p_str_period) < 7:
            raise Exception("Invalid period lenght.")
    else:
        raise PaymentPlanDebugError("The given period parameter is not a string.")

    p_str_period = p_str_period.strip(" ")

    month = int(p_str_period[5:7])
    year = int(p_str_period[0:4])

    if int(p_day) == 0 or int(p_day) > 31:
        p_day = 31

    if month in (4, 6, 9, 11) and int(p_day) > 30:
        p_day = 30

    if month == 2 and int(p_day) > 28 and year % 4 == 0:
        p_day = 29
    elif month == 2 and int(p_day) > 28:
        p_day = 28

    perioddate = datetime.datetime(year=year, month=month, day=p_day)
    clarionstart = datetime.datetime(year=1800, month=12, day=28)
    diff = perioddate - clarionstart
    return int(diff.days)
