import json


class BJSONEnc(json.JSONEncoder):
    def default(self, o):
        try:
            json.dumps(o)
        except:
            return json.dumps("".format(o))
        #
        return o

    #


def jsoncomplex(dct):
    if "__complex__" in objct:
        return "{}".format(dct)
    #
    return dct


#
