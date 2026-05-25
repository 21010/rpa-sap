import win32com.client
from dotenv import dotenv_values
from rpa_sap.core.connection import ConnectionManager

secrets = dotenv_values('.env')
cm = ConnectionManager()
session = cm.open_new_session(
    secrets["SAP_CONNECTION_STRING"],
    secrets["SAP_USER_ID"],
    secrets["SAP_PASSWORD"],
    secrets["SAP_CLIENT"],
    secrets["SAP_LANGUAGE"]
)

print("Session info:")
print(session.info)

sap_functions = win32com.client.Dispatch("SAP.Functions")
try:
    print("Trying to assign GUI connection to SAP.Functions.Connection...")
    sap_functions.Connection = session.com_connection
    print("Success! SAP.Functions accepts GUI connection.")
except Exception as e:
    print(f"Failed to assign com_connection: {e}")

try:
    print("Trying to assign GUI session to SAP.Functions.Connection...")
    sap_functions.Connection = session.com_session
    print("Success! SAP.Functions accepts GUI session.")
except Exception as e:
    print(f"Failed to assign com_session: {e}")

cm.close_sap_logon()
