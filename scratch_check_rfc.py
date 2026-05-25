import win32com.client
from dotenv import dotenv_values

secrets = dotenv_values('.env')
print("Trying to connect to SAPGUI and inspect objects...")
sap_gui = win32com.client.GetObject("SAPGUI")
application = sap_gui.GetScriptingEngine
connection = application.Connections[0]
session = connection.Sessions[0]

print("Session info:")
print(session.Info.SystemName)
print(session.Info.ApplicationServer)

sap_functions = win32com.client.Dispatch("SAP.Functions")
try:
    print("Trying to assign GUI connection to SAP.Functions.Connection...")
    sap_functions.Connection = connection
    print("Success! SAP.Functions accepts GUI connection.")
except Exception as e:
    print(f"Failed: {e}")

try:
    print("Trying to assign GUI session to SAP.Functions.Connection...")
    sap_functions.Connection = session
    print("Success! SAP.Functions accepts GUI session.")
except Exception as e:
    print(f"Failed: {e}")

