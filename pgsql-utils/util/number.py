

def tryInt(p_str):
  try:
    return int(p_str)
  except Exception as e:
    return 0
  #
#


def isFloat(p_str):
  try:
    if str(p_str).isnumeric():
      val = float(p_str)
      if (val is not None):
        return True
      #
    #
  except Exception as e:
    return False
  #
  return False
#
def tryFloat(p_str, p_def = 0):
  try:
    return float(p_str)
  except Exception as e:
    if p_def is None:
      return p_def
    return float(p_def)
  #
#


def getNumber(p_str, p_def = 0):
  try:
    numbers = ""
    cnt = 0
    firstnum = 0
    for s in p_str:
      cnt += 1
      if s.isdigit() or s == "." or s == "," or (cnt < 2 and s == "-") or s == "+":
        numbers = numbers + s
        if s.isdigit():
          firstnum = cnt
        #
      #
    #

    #to much text and numbers, do not extract
    if firstnum > len(numbers)*2:
      return p_def
    #

    return float(numbers)
  except Exception as e:
    if p_def is None:
      return p_def
    return float(p_def)
  #
#

def nint(p_numstr):
  parts = str(p_numstr).split(".")
  if len(parts) > 1:
    return parts[0]
  parts = str(p_numstr).split(",")
  if len(parts) > 1:
    return parts[0]

  return p_numstr
#


def ndec(p_numstr, p_num = None):
  if p_num is None:
    p_num = 2
  else:
    p_num = int(p_num)
  #

  parts = str(p_numstr).split(".")
  if len(parts) > 1:
    n = tryFloat(parts[1])
    return str(round(n,p_num))[:p_num]
  #

  parts = str(p_numstr).split(",")
  if len(parts) > 1:
    n = tryFloat(parts[1])
    return str(round(n,p_num))[:p_num]
  #

  return p_numstr
#