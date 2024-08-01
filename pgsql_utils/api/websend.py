### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.


import json
import io, sys, os, traceback
import requests
import urllib.parse
import tempfile
import hashlib
import base64
import urllib3


def callWebSend(p_output, p_errmsg, p_retval, p_extradata, p_http_code):

    global p_output, p_errmsg, p_retval, p_extradata, p_http_code
    global m_feedback_cookies, m_noerror

    opts = {}
    p_retval = 0
    p_http_code = 0
    p_errmsg = ""
    p_output = None
    p_extradata = None
    m_sslverify = True
    m_noerror = False
    m_btreval = True

    m_headers = {}
    m_headers["User-Agent"] = "Bitech/8.0"
    # m_headers['Cache-control'] = "no-cache"
    # m_headers['Connection'] = "keep-alive"
    # m_headers['Pragma'] = "no-cache"
    # m_headers['Content-Type'] = "application/json"
    # m_headers['x-api-key'] = "simplicity-dev"

    m_proxies = {}
    m_parms = {}
    m_cookies = {}
    m_url = ""
    m_devflag = False
    m_allowredirect = True
    m_stream = False
    m_cert = None
    m_timeout = 10
    m_readtimeout = 600
    m_action = "get"
    m_username = ""
    m_password = ""
    m_feedback_cookies = None
    m_base64response = False

    try:
        opts = json.loads(p_option)
        if not isinstance(opts, dict):
            raise Exception("No object. Options requires an JSON object.")
    except Exception as e:
        p_retval = 1
        p_errmsg = "Invalid options JSON.".format(e)
        return [p_retval, p_errmsg, p_output, "", 0]
    #

    def get_err_msg(
        p_funcname="pl_web_send", p_errmsg="", p_errdetail="", p_context=""
    ):
        errmsg = ""
        rs = plpy.execute(
            """select get_err_msg({}, {}, {}, {}, '', '') as errmsg;""".format(
                plpy.quote_literal(p_funcname),
                plpy.quote_literal(p_errmsg),
                plpy.quote_literal(p_context),
                plpy.quote_literal(p_errdetail),
            )
        )
        if rs is not None and rs[0] is not None:
            errmsg = str(rs[0]["errmsg"])

        return errmsg

    #

    def sqlmsg(p_msg, p_title="pl_web_send", p_type="local notice"):
        if "error" in p_type:
            plpy.warning(p_title + " -> " + p_msg)
        #
        rs = plpy.execute(
            """select public.log_event({},{}, public.bt_enum('eventlog'::text, {}));""".format(
                plpy.quote_literal(p_title),
                plpy.quote_literal(p_msg),
                plpy.quote_literal(p_type),
            )
        )
        #

    #

    def http_action():
        global p_output, p_http_code
        global m_feedback_cookies

        response = None
        res = None
        m_auth = None
        if len(m_username) > 1:
            # m_auth = {"username": m_username, "password": m_password}
            m_auth = requests.auth.HTTPBasicAuth(m_username, m_password)
        #

        if m_action == "post":
            res = requests.request(
                method=m_action,
                url=m_url,
                data=p_infile.encode("utf8"),
                headers=m_headers,
                timeout=(m_timeout, m_readtimeout),
                proxies=m_proxies,
                verify=m_sslverify,
                params=m_parms,
                cookies=m_cookies,
                allow_redirects=m_allowredirect,
                auth=m_auth,
                stream=m_stream,
                cert=m_cert,
            )
        elif m_action == "patch":
            res = requests.request(
                method=m_action,
                url=m_url,
                data=p_infile.encode("utf8"),
                headers=m_headers,
                timeout=(m_timeout, m_readtimeout),
                proxies=m_proxies,
                verify=m_sslverify,
                params=m_parms,
                cookies=m_cookies,
                allow_redirects=m_allowredirect,
                auth=m_auth,
                stream=m_stream,
                cert=m_cert,
            )
        elif m_action == "delete":
            res = requests.request(
                method=m_action,
                url=m_url,
                headers=m_headers,
                timeout=(m_timeout, m_readtimeout),
                proxies=m_proxies,
                verify=m_sslverify,
                params=m_parms,
                cookies=m_cookies,
                allow_redirects=m_allowredirect,
                auth=m_auth,
                stream=m_stream,
                cert=m_cert,
            )
        elif m_action == "head":
            res = requests.request(
                method=m_action,
                url=m_url,
                headers=m_headers,
                timeout=(m_timeout, m_readtimeout),
                proxies=m_proxies,
                verify=m_sslverify,
                params=m_parms,
                cookies=m_cookies,
                allow_redirects=m_allowredirect,
                auth=m_auth,
                stream=m_stream,
                cert=m_cert,
            )
        elif m_action == "options":
            res = requests.request(
                method=m_action,
                url=m_url,
                headers=m_headers,
                timeout=(m_timeout, m_readtimeout),
                proxies=m_proxies,
                verify=m_sslverify,
                params=m_parms,
                cookies=m_cookies,
                allow_redirects=m_allowredirect,
                auth=m_auth,
                stream=m_stream,
                cert=m_cert,
            )
        else:
            res = requests.request(
                method=m_action,
                url=m_url,
                headers=m_headers,
                timeout=(m_timeout, m_readtimeout),
                proxies=m_proxies,
                verify=m_sslverify,
                params=m_parms,
                cookies=m_cookies,
                allow_redirects=m_allowredirect,
                auth=m_auth,
                stream=m_stream,
                cert=m_cert,
            )
        #

        if res is not None:

            if res.cookies is not None and len(res.cookies.keys()) > 0:
                m_feedback_cookies = res.cookies.items()
            #
            # plpy.notice("Got cookies: Type: {} Items: {} m_feedback_cookies: {}".format(type(res.cookies), str(res.cookies.items()) , str(m_feedback_cookies) ) )
            try:
                p_http_code = int(res.status_code)
            except Exception:
                p_http_code = 0
            #
            if isinstance(res.content, str):
                p_output = res.content
            if isinstance(res.content, bytes):
                if m_base64response:
                    bt = base64.b64encode(res.content)
                    p_output = bt.decode()
                else:
                    p_output = res.content.decode()
            else:
                p_output = str()
                raise Exception("Unknown return type from request.")
            #

            if not res.ok and not m_noerror:
                if m_btreval:
                    rvdata = None
                    try:
                        rvdata = json.loads(p_output)
                        if not (
                            isinstance(rvdata, dict) and rvdata.get("retval", "0") > "0"
                        ):
                            rvdata = None
                        #
                    except Exception as e:
                        rvdata = None
                    #
                    if rvdata is None:
                        raise Exception(
                            "HTTP Error: {}, {} \nDetail: {} \nURL:{} \nAction:{}".format(
                                str(res.status_code),
                                str(res.reason),
                                p_output,
                                m_url,
                                m_action,
                            )
                        )
                    else:
                        raise Exception(
                            "{} \nDetail: \nCode:{}, {}  \nURL:{} \nAction:{}".format(
                                rvdata.get("errmsg", ""),
                                str(res.status_code),
                                str(res.reason),
                                m_url,
                                m_action,
                            )
                        )
                    #
                else:
                    raise Exception(
                        "HTTP Error: {}, {} \nDetail: {} \nURL:{} \nAction:{}".format(
                            str(res.status_code),
                            str(res.reason),
                            p_output,
                            m_url,
                            m_action,
                        )
                    )
                #
            #

            res.close()
        else:
            raise Exception("No response from HTTP.")
        #
        return True

    #

    def pfxRead():
        import OpenSSL.crypto

        keyfile = m_pfxcert.get("file", "")
        pempath = m_pfxcert.get("pem", "")
        if pempath == "":
            key = hashlib.sha224(keyfile.encode("utf8")).hexdigest()
            pempath = os.path.join(tempfile.gettempdir(), "." + key + ".md5")
            # plpy.notice("pempath: {}".format(pempath))
            if os.path.exists(pempath):
                # plpy.notice("using exiting file: {}".format(pempath))
                return pempath
            #
        #

        pfx_password = m_pfxcert.get("password", "")
        pfx_data = None
        if not os.path.exists(keyfile):
            raise Exception("Given keyfile does not exist. {}".format(keyfile))
        #
        with open(keyfile, "rb") as f:
            pfx_data = f.read()
        #
        pemfile = open(pempath, "wb")

        p12 = OpenSSL.crypto.load_pkcs12(pfx_data, pfx_password)
        pemfile.write(
            OpenSSL.crypto.dump_privatekey(
                OpenSSL.crypto.FILETYPE_PEM, p12.get_privatekey()
            )
        )
        pemfile.write(
            OpenSSL.crypto.dump_certificate(
                OpenSSL.crypto.FILETYPE_PEM, p12.get_certificate()
            )
        )
        ca = p12.get_ca_certificates()
        if ca is not None:
            for cert in ca:
                pemfile.write(
                    OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, cert)
                )
            #
        #
        pemfile.close()

        # sqlmsg("Key Saved to: {}".format(pempath))
        return pempath

    #

    try:

        m_timeout = int(opts.get("timeout", "60"))
        m_readtimeout = int(opts.get("readtimeout", m_timeout))

        m_action = str(opts.get("action", "get")).lower()
        m_url = opts.get("url", "")
        m_username = opts.get("username", "")
        m_password = opts.get("password", "")
        if int(opts.get("dev", "0")) > 0:
            m_devflag = True
        if int(opts.get("nosslverify", "0")) > 0:
            m_sslverify = False
        if int(opts.get("disableredirect", "0")) > 0:
            m_allowredirect = False
        if int(opts.get("stream", "0")) > 0:
            m_stream = True
        if int(opts.get("noerror", "0")) > 0:
            m_noerror = True
        if int(opts.get("base64bytes", "0")) > 0:
            m_base64response = True
        for h in opts.get("headers", {}):
            m_headers[h] = opts["headers"][h]
        #

        for h in opts.get("proxies", {}):
            m_proxies[h] = opts["proxies"][h]
        #

        for h in opts.get("cookies", {}):
            m_cookies[h] = opts["cookies"][h]
        #

        for h in opts.get("parms", {}):
            m_parms[h] = opts["parms"][h]
        #

        m_cert = opts.get("cert", None)
        if isinstance(m_cert, dict):
            raise Exception("Please provide a list. e.g. [file.cert, file.key]")
        #
        if isinstance(m_cert, list):
            m_cert = (m_cert[0], m_cert[1])
        #

        m_pfxcert = opts.get("pfxcert", None)
        if m_pfxcert is not None and not isinstance(m_pfxcert, dict):
            raise Exception(
                'Please provide a object. e.g. {"file":"filename","password":"pfx password" }'
            )
        #

        if m_url == "":
            raise AssertionError("No URL specified.")
        if m_action not in ("post", "get", "put", "patch", "delete", "head", "options"):
            raise AssertionError("No URL specified.")

        # plpy.notice('init, m_url={}, m_devflag={}, m_action={}, m_sslverify={}, m_timeout={}, headers={}'.format(m_url, m_devflag, m_action, m_sslverify, m_timeout, opts.get('headers')))
        if m_pfxcert is not None:
            m_cert = pfxRead()
        #

        http_action()

        p_extradata = {}
        if m_feedback_cookies is not None:
            p_extradata["cookies"] = m_feedback_cookies
        #

        p_extradata = json.dumps(p_extradata)

    except requests.exceptions.Timeout as e:
        p_retval = 1
        p_errmsg = get_err_msg(
            p_errmsg="Connection Timeout {}, check the network connection.".format(
                m_url
            ),
            p_errdetail=str(e),
        )

    except urllib3.exceptions.NewConnectionError as e:
        p_retval = 1
        p_errmsg = get_err_msg(
            p_errmsg="Connection Failed {}, check the network connection.".format(
                m_url
            ),
            p_errdetail=str(e),
        )

    except urllib3.exceptions.TimeoutError as e:
        p_retval = 1
        p_errmsg = get_err_msg(
            p_errmsg="Connection Timeout {}, check the network connection.".format(
                m_url
            ),
            p_errdetail=str(e),
        )

    except requests.exceptions.ConnectionError as e:
        p_retval = 1
        p_errmsg = get_err_msg(
            p_errmsg="Connection Timeout {}, check the network connection.".format(
                m_url
            ),
            p_errdetail=str(e),
        )

    except requests.exceptions.RequestException as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_retval = 1
        p_errmsg = get_err_msg(
            p_errmsg=str(e),
            p_errdetail=str(e) + " " + str(traceback.format_tb(exc_tb, 5)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )

    except plpy.SPIError as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_retval = 1
        p_errmsg = get_err_msg(
            p_errmsg=str(e),
            p_errdetail=str(traceback.format_tb(exc_tb, 5)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )

    except Exception as inst:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_errmsg = get_err_msg(
            p_errmsg=str(inst),
            p_errdetail=str(traceback.format_tb(exc_tb, 5)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )
        p_retval = 1
    except:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_errmsg = get_err_msg(
            p_errmsg=str(exc_type),
            p_errdetail=str(traceback.format_tb(exc_tb, 5)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )
        p_retval = 1

    return [p_retval, p_errmsg, p_output, p_extradata, p_http_code]
