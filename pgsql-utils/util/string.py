from .number import tryFloat
from num2words  import num2words

def ordinalstr(p_num):
  l = int(str(p_num)[-1:])
  if p_num > 10 and p_num < 20:
    l = 5
  #
  if l == 0: return "{}th".format(p_num)
  if l == 1: return "{}st".format(p_num)
  if l == 2: return "{}nd".format(p_num)
  if l == 3: return "{}rd".format(p_num)
  if l >= 4: return "{}th".format(p_num)

  return str(p_num)
#


def n2w(p_num, p_type="cardinal",  p_dec = 2):
  if p_num is None: return p_num
  #if isinstance(p_num, str): return p_num
  words = ""
  try:
    if isinstance(p_num, str):
      p_num = round(tryFloat(p_num),p_dec)
    if p_num < 0:
      p_num = round(tryFloat(p_num),p_dec)
      words = str(num2words(p_num*-1, lang="en_GB", to=p_type)).replace(",", "")
      return str(num2words(p_num*-1, lang="en_GB", to=p_type)).capitalize()
    #
    p_num = round(tryFloat(p_num),p_dec)
    words = str(num2words(p_num, lang="en_GB", to=p_type)).replace(",", "")
    return words.capitalize()

  except Exception as e:
    #sqlmsg("n2w error, {}: input:{}".format(e,p_num), p_type="local error")
    return str(p_num).capitalize()
  #

#

def istrreplace(p_text, p_old, p_new):
  idx = 0
  while idx < len(p_text):
    index_l = p_text.lower().find(p_old.lower(), idx)
    if index_l == -1:
      return p_text
    p_text = p_text[:index_l] + p_new + p_text[index_l + len(p_old):]
    idx = index_l + len(p_new)
  return p_text
#



def html_replace(p_txt):
  pr = "&"
  pf = ";"
  newtxt = p_txt.replace(pr + "nbsp" + pf," ")
  newtxt = newtxt.replace(pr + "gt" + pf,">")
  newtxt = newtxt.replace(pr +"lt" + pf,"<")
  newtxt = newtxt.replace(pr +"amp" + pf,"&")
  newtxt = newtxt.replace(pr +"#64" + pf,"@")
  newtxt = newtxt.replace(pr +"quot" + pf,"\"")
  newtxt = newtxt.replace(pr +"#34" + pf,"\"")
  newtxt = newtxt.replace(pr +"#61" + pf,"=")

  return newtxt
#


def findWordArticle(pWord : str):
  cleaned = pWord.strip(" ")
  if len(cleaned) == 0:
    return ""
  #

  vowels = "aeiou"
  exceptions = ["exray","u"] ##Sound similiar but not vowel
  if pWord in exceptions:
    return "an"
  #
  if cleaned[0].lower() in vowels:
    return "an"
  #
  return "a"
#