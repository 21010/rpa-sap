import win32com.client

class SapRFCClient:
    def __init__(self, ashost, sysnr, client, user, passwd):
        try:
            self.logon_control = win32com.client.Dispatch("SAP.LogonControl.1")
            print("Dispatch LogonControl success.")
            self.connection = self.logon_control.NewConnection()
            print("NewConnection success.")
            
            # Configure connection parameters
            print(f"Setting ApplicationServer: {ashost}")
            self.connection.ApplicationServer = ashost
            print(f"Setting SystemNumber: {sysnr}")
            self.connection.SystemNumber = sysnr
            print(f"Setting Client: {client}")
            self.connection.Client = client
            print("Setting User")
            self.connection.User = user
            print("Setting Password")
            self.connection.Password = passwd
            print("Setting UseSAPLogonIni")
            self.connection.UseSAPLogonIni = False

            print("Attempting Logon...")
            # We set Silent to False so that SAP GUI can pop up the exact error message!
            if not self.connection.Logon(0, False):
                raise Exception("Could not connect to SAP. Check credentials and server details.")
                
            self.sap_functions = win32com.client.Dispatch("SAP.Functions")
            self.sap_functions.Connection = self.connection
        except Exception as e:
            raise Exception(f"Failed to initialize SAP COM connection: {e}")

    def read_table(self, table_name, fields=None, options=None):
        rfc = self.sap_functions.Add("RFC_READ_TABLE")
        
        rfc.Exports("QUERY_TABLE").Value = table_name
        rfc.Exports("DELIMITER").Value = ";"
        
        # Add options (WHERE clause)
        if options:
            options_table = rfc.Tables("OPTIONS")
            options_table.FreeTable()
            # Convert list of strings to tuple of tuples e.g. (("SPRAS = 'PL'",),)
            options_data = tuple((opt,) for opt in options)
            options_table.Data = options_data
                
        # Add fields
        if fields:
            fields_table = rfc.Tables("FIELDS")
            fields_table.FreeTable()
            # Structure: ((FIELDNAME, OFFSET, LENGTH, TYPE, FIELDTEXT), ...)
            fields_data = tuple((f, '000000', '000000', '', '') for f in fields)
            fields_table.Data = fields_data

        # Execute
        if rfc.Call:
            # Extract actual headers returned
            res_fields = rfc.Tables("FIELDS")
            headers = [row[0].strip() for row in res_fields.Data]  # row[0] is FIELDNAME
            
            results = []
            data_table = rfc.Tables("DATA")
            if data_table.RowCount > 0:
                for row in data_table.Data:
                    row_string = row[0] # 'WA' is the first column
                    # Split by delimiter
                    row_values = row_string.split(";")
                    
                    row_dict = {}
                    for j, header in enumerate(headers):
                        if j < len(row_values):
                            row_dict[header] = row_values[j].strip()
                        else:
                            row_dict[header] = ""
                    results.append(row_dict)
                
            return results
        else:
            raise Exception("Error calling RFC_READ_TABLE. Check table name, fields, and permissions.")
