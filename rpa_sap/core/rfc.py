import re
import win32com.client
from ..exceptions import SapSessionError, SapRfcError

class RfcConnection:
    """
    Represents a headless SAP RFC connection using SAP.Functions.
    """

    def __init__(self, connection_string: str, user_id: str, password: str, client: str = "900", language: str = "EN"):
        self.connection_string = connection_string
        self.user_id = user_id
        self.password = password
        self.client = client
        self.language = language
        self.sap_functions = None
        self._connect()

    def _connect(self):
        """Initializes the connection via SAP.Functions."""
        try:
            import pythoncom
            pythoncom.CoInitialize()

            self.sap_functions = win32com.client.Dispatch("SAP.Functions")
            rfc_conn = self.sap_functions.Connection

            match = re.search(r"/H/([^/]+)/S/(32\d{2})", self.connection_string)

            if match:
                ashost = match.group(1)
                sysnr = match.group(2)[2:]
                rfc_conn.ApplicationServer = ashost
                rfc_conn.SystemNumber = sysnr
                rfc_conn.UseSAPLogonIni = False
            else:
                rfc_conn.System = self.connection_string
                rfc_conn.UseSAPLogonIni = True

            rfc_conn.Client = self.client
            rfc_conn.User = self.user_id
            rfc_conn.Password = self.password
            rfc_conn.Language = self.language

            if not rfc_conn.Logon(0, True):  # Silent=True
                raise Exception("Logon method returned False.")
                
            self.connection = rfc_conn

        except Exception as e:
            raise SapSessionError(
                f"Failed to establish headless RFC connection: {e}"
            ) from e

    def close(self):
        """Explicitly logs off and closes the RFC connection."""
        if hasattr(self, "connection") and self.connection is not None:
            try:
                self.connection.Logoff()
            except Exception:
                pass

    def execute_rfc(self, function_name: str):
        """
        Retrieves an RFC function object to be executed.

        Args:
            function_name (str): Name of the RFC or BAPI function module.

        Returns:
            The RFC function COM object.
        """
        return self.sap_functions.Add(function_name)

    def read_table(
        self, table_name: str, fields: list = None, options: list = None
    ) -> list[dict]:
        """
        Reads a transparent table from SAP using RFC_READ_TABLE.

        Args:
            table_name (str): Name of the SAP table (e.g., 'T000')
            fields (list, optional): List of field names to extract.
            options (list, optional): List of WHERE clauses (e.g., ["SPRAS = 'E'"]).

        Returns:
            list[dict]: A list of dictionaries containing the table rows.

        Raises:
            SapRfcError: If the RFC execution fails.
        """
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
            fields_data = tuple((f, "000000", "000000", "", "") for f in fields)
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
