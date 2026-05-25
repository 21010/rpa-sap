import win32com.client
from dotenv import dotenv_values
import re

secrets = dotenv_values('.env')
connection_string = secrets["SAP_CONNECTION_STRING"]
match = re.search(r"/H/([^/]+)/S/(32\d{2})", connection_string)
if match:
    ashost = match.group(1)
    sysnr = match.group(2)[2:]
    try:
        sap_functions = win32com.client.Dispatch("SAP.Functions")
        rfc_conn = sap_functions.Connection
        print("Got connection from SAP.Functions!")
        rfc_conn.ApplicationServer = ashost
        rfc_conn.SystemNumber = sysnr
        rfc_conn.Client = secrets["SAP_CLIENT"]
        rfc_conn.User = secrets["SAP_USER_ID"]
        rfc_conn.Password = secrets["SAP_PASSWORD"]
        rfc_conn.UseSAPLogonIni = False
        
        if rfc_conn.Logon(0, True):
            print("Logon successful!")
        else:
            print("Logon failed.")
    except Exception as e:
        print(f"Exception: {e}")
