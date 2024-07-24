if 'plpy' not in globals():
    import pgsql_utils.pgproc.fake_plpy as plpy
#
 
def get_err_msg(p_funcname="function_name", p_errmsg="", p_errdetail="", p_context=""):
  errmsg = ""
  rs = plpy.execute("""select get_err_msg({}, {}, {}, {}, '', '') as errmsg;""".format( \
    plpy.quote_literal(p_funcname), plpy.quote_literal(p_errmsg) , plpy.quote_literal(p_context), plpy.quote_literal(p_errdetail) \
  ))
  if rs is not None and rs.get(0) is not None:
    errmsg = str(rs[0]['errmsg'])

  return errmsg
#

def sqlmsg(p_msg,p_title = "function_name",p_type = "local notice"):
  if "error" in p_type:
    plpy.warning(p_title + " -> " + p_msg)
  #
  rs = plpy.execute("""select public.log_event({},{}, public.bt_enum('eventlog'::text, {}));""".format(
    plpy.quote_literal(p_title), plpy.quote_literal(p_msg) , plpy.quote_literal(p_type)))
  #
#
