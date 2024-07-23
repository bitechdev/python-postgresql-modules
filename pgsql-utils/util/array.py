from .number import tryFloat

def sum_any(p_list):
  total = 0.0
  for i, val in enumerate(p_list):
    total = total + tryFloat(val,0.0)
  #
  return total
#

def max_any(p_list):
  maxval = 0.0
  for i, val in enumerate(p_list):
    if tryFloat(val,0.0) >= maxval:
      maxval = val
    #
  #
  return maxval
#

def min_any(p_list):
  minval = 0.0
  for i, val in enumerate(p_list):
    if tryFloat(val,0.0) <= minval:
      minval = val
    #
  #
  return minval
#

def avg_any(p_list):
  cleanlist = []
  for i, val in enumerate(p_list):
    converted =  tryFloat(val,0)
    if converted > 0:
      cleanlist.append(converted)
    #
  #
  return avg(cleanlist)
#
