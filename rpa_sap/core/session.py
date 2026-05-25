import win32com.client
from ..exceptions import SapSessionError, SapTransactionError
from .ui_automation import ElementInteractor


class SapSession:
    """
    Represents an active SAP GUI session and acts as the main context
    for UI automation and transaction execution.
    """

    def __init__(
        self,
        com_session: win32com.client.CDispatch,
        com_connection: win32com.client.CDispatch = None,
        sap_functions: win32com.client.CDispatch = None,
    ):
        self.com_session = com_session
        self.com_connection = com_connection
        self.sap_functions = sap_functions
        self.interactor = ElementInteractor(self)

    @property
    def info(self) -> dict:
        """Returns basic information about the session."""
        return {
            "user": self.com_session.Info.User.upper(),
            "sid": self.com_session.Info.SystemName.upper(),
            "application_server": self.com_session.Info.ApplicationServer.upper(),
            "client": self.com_session.Info.Client.upper(),
            # "is_active": self.com_session.IsActive,
            "is_active": self.com_session.ActiveWindow.Text != "",
            "is_busy": self.com_session.Busy,
            "session_number": self.com_session.Info.SessionNumber,
            "transaction": self.com_session.Info.Transaction,
        }

    @property
    def id(self) -> str:
        return self.com_session.Id

    def findById(self, field_id: str):
        """Find an element by its ID within the active session."""
        return self.com_session.findById(field_id)

    def run_transaction(self, transaction_code: str):
        """
        Runs SAP transaction.
        There is no need to add "/n" or go back to the start screen.
        """
        self.com_session.StartTransaction(transaction_code)
        status = self.interactor.get_status_bar_message()
        if status.type == "E":
            raise SapTransactionError(f"{status.type} : {status.text}")

    def stop_transaction(self):
        """Stops SAP transaction."""
        self.com_session.EndTransaction()

    def set_active_window(self, index: int):
        """Sets active window for active SAP session."""
        return self.com_session.findById(f"wnd[{index}]")

    def execute_rfc(self, function_name: str):
        """
        Retrieves an RFC function object to be executed.
        
        Args:
            function_name (str): Name of the RFC or BAPI function module.
            
        Returns:
            The RFC function COM object.
            
        Raises:
            SapSessionError: If the RFC connection was not established.
        """
        if not self.sap_functions:
            if not hasattr(self, "_rfc_credentials"):
                raise SapSessionError("RFC connection is not established and no credentials were provided.")
            import re
            match = re.search(r"/H/([^/]+)/S/(32\d{2})", self._rfc_credentials["connection_string"])
            if not match:
                raise SapSessionError("Could not parse connection string for RFC (needs /H/.../S/...)")
            ashost = match.group(1)
            sysnr = match.group(2)[2:]
            
            try:
                import pythoncom
                pythoncom.CoInitialize()
                
                self.sap_functions = win32com.client.Dispatch("SAP.Functions")
                rfc_conn = self.sap_functions.Connection
                rfc_conn.ApplicationServer = ashost
                rfc_conn.SystemNumber = sysnr
                rfc_conn.Client = self._rfc_credentials["client"]
                rfc_conn.User = self._rfc_credentials["user_id"]
                rfc_conn.Password = self._rfc_credentials["password"]
                rfc_conn.UseSAPLogonIni = False
    
                if not rfc_conn.Logon(0, True): # Silent=True
                    raise Exception("Logon method returned False.")
    
            except Exception as e:
                raise SapSessionError(f"Failed to establish dynamic RFC connection: {e}") from e
                
        return self.sap_functions.Add(function_name)

    def read_table(self, table_name: str, fields: list = None, options: list = None) -> list[dict]:
        """
        Reads a transparent table from SAP using RFC_READ_TABLE.
        
        Args:
            table_name (str): Name of the SAP table (e.g., 'T000')
            fields (list, optional): List of field names to extract.
            options (list, optional): List of WHERE clauses (e.g., ["SPRAS = 'E'"]).
            
        Returns:
            list[dict]: A list of dictionaries containing the table rows.
            
        Raises:
            SapSessionError: If the RFC execution fails.
        """
        from ..exceptions import SapRfcError
        
        rfc = self.execute_rfc("RFC_READ_TABLE")
        rfc.Exports("QUERY_TABLE").Value = table_name
        rfc.Exports("DELIMITER").Value = ";"
        
        if options:
            options_table = rfc.Tables("OPTIONS")
            options_table.FreeTable()
            options_data = tuple((opt,) for opt in options)
            options_table.Data = options_data
            
        if fields:
            fields_table = rfc.Tables("FIELDS")
            fields_table.FreeTable()
            fields_data = tuple((f, '000000', '000000', '', '') for f in fields)
            fields_table.Data = fields_data

        if not rfc.Call:
            raise SapRfcError(f"Error calling RFC_READ_TABLE for table {table_name}")
            
        res_fields = rfc.Tables("FIELDS")
        headers = [row[0].strip() for row in res_fields.Data]
        
        results = []
        data_table = rfc.Tables("DATA")
        if data_table.RowCount > 0:
            for row in data_table.Data:
                row_string = row[0]
                row_values = row_string.split(";")
                row_dict = {}
                for j, header in enumerate(headers):
                    if j < len(row_values):
                        row_dict[header] = row_values[j].strip()
                    else:
                        row_dict[header] = ""
                results.append(row_dict)
                
        return results
