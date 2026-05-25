import pytest
from dotenv import dotenv_values
from rpa_sap.core.connection import ConnectionManager

def main():
    secrets = dotenv_values('.env')
    cm = ConnectionManager()
    print("Opening session...")
    sap_session = cm.open_new_session(
        secrets["SAP_CONNECTION_STRING"],
        secrets["SAP_USER_ID"],
        secrets["SAP_PASSWORD"],
        secrets["SAP_CLIENT"],
        "EN"
    )
    
    print("Executing RFC 1...")
    try:
        rfc = sap_session.execute_rfc("RFC_READ_TABLE")
        print("RFC 1 success!")
    except Exception as e:
        print(f"RFC 1 failed: {e}")
        
    print("Executing RFC 2...")
    try:
        rfc = sap_session.execute_rfc("RFC_READ_TABLE")
        print("RFC 2 success!")
    except Exception as e:
        print(f"RFC 2 failed: {e}")

    cm.close_all_sessions()

if __name__ == "__main__":
    main()
