def getProp(obj, name, defValue=None):
    prop = defValue
    try:
        prop = obj[name]
    except:
        prop = defValue
    #
    if prop is None and defValue is not None:
        prop = defValue
    #
    return prop


#
