def isnone(p_var, p_def=None):
    if p_var is None:
        return p_def
    else:
        return p_var


def chkif(p_check, p_trueval, p_falseval):
    if p_check:
        return p_trueval

    return p_falseval


def date_add_mon(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = int(sourcedate.year + month / 12)
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)


#
