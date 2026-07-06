import pytest
from unittest.mock import MagicMock, patch
from rpa_sap.core.rfc import RfcConnection
from rpa_sap.exceptions import SapRfcError

@patch("rpa_sap.core.rfc.pythoncom.CoInitialize")
@patch("rpa_sap.core.rfc.win32com.client.Dispatch")
def test_call_bapi_success(mock_dispatch, mock_coinit):
    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_functions.Connection.Logon.return_value = True
    
    rfc_conn = RfcConnection("022 BPE Test", "test_user", "test_password")
    
    mock_bapi = MagicMock()
    mock_bapi.Call = True
    
    # Mocking Tables and Exports for the BAPI
    mock_return_table = MagicMock()
    mock_return_table.RowCount = 1
    
    mock_row = MagicMock()
    mock_return_table.Rows.return_value = mock_row
    
    mock_input_table = MagicMock()
    
    def side_effect_tables(name):
        if name == "RETURN":
            return mock_return_table
        if name == "INPUT_TABLE":
            return mock_input_table
        return MagicMock()
        
    mock_bapi.Tables = side_effect_tables
    
    # Mocking Imports
    mock_export_val = MagicMock()
    mock_export_val.Value = "SUCCESS"
    def side_effect_imports(name):
        if name == "E_STATUS":
            return mock_export_val
        return MagicMock()
        
    mock_bapi.Imports = side_effect_imports
    
    # Commit mock
    mock_commit = MagicMock()
    mock_commit.Call = True
    
    def side_effect_add(name):
        if name == "BAPI_TEST":
            return mock_bapi
        if name == "BAPI_TRANSACTION_COMMIT":
            return mock_commit
        return MagicMock()
        
    mock_functions.Add = side_effect_add
    
    import_params = {"I_PARAM": "VALUE"}
    table_params = {"INPUT_TABLE": [{"FIELD1": "VAL1"}]}
    
    result = rfc_conn.call_bapi(
        bapi_name="BAPI_TEST",
        import_params=import_params,
        table_params=table_params,
        extract_tables=["RETURN"],
        extract_imports=["E_STATUS"],
        commit=True
    )
    
    # Verifications
    mock_bapi.Exports.assert_called_with("I_PARAM")
    mock_input_table.FreeTable.assert_called_once()
    mock_input_table.Rows.Add.assert_called_once()
    
    mock_commit.Exports.assert_called_with("WAIT")
    
    assert "EXPORTS" in result
    assert result["EXPORTS"]["E_STATUS"] == "SUCCESS"
    assert "TABLES" in result
    assert result["TABLES"]["RETURN"] == mock_return_table
