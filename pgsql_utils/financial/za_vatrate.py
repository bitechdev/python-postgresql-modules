### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.


import datetime

ZA_VAT_TABLE = {
    "1991-09-01": 10,
    "1993-04-01": 14,
    "2018-04-01": 15,
}


def getVatRate(p_date=None):
    date = datetime.datetime.now().isoformat()
    if p_date == "":
        p_date = None

    if isinstance(p_date, str):
        date = p_date
    elif isinstance(p_date, datetime.date):
        date = p_date.isoformat()

    for k, v in ZA_VAT_TABLE.items():
        if p_date >= k:
            return v

    return None
