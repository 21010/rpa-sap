from unittest.mock import MagicMock, patch
from rpa_sap.core.rfc import RfcConnection


@patch("rpa_sap.core.rfc.pythoncom.CoInitialize")
@patch("rpa_sap.core.rfc.win32com.client.Dispatch")
def test_rfc_connection_dynamic_string(mock_dispatch, mock_coinit):
    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_connection = mock_functions.Connection
    mock_connection.Logon.return_value = True

    rfc = RfcConnection(
        connection_string="/H/my-server/S/3200",
        user_id="test_user",
        password="test_password",
        client="900",
        language="EN",
    )

    mock_dispatch.assert_called_with("SAP.Functions")
    assert mock_connection.ApplicationServer == "my-server"
    assert mock_connection.SystemNumber == "00"
    assert mock_connection.UseSAPLogonIni is False
    assert mock_connection.Client == "900"
    assert mock_connection.User == "test_user"
    assert mock_connection.Password == "test_password"
    assert mock_connection.Language == "EN"
    mock_connection.Logon.assert_called_with(0, True)

    mock_func_obj = MagicMock()
    mock_functions.Add.return_value = mock_func_obj

    res = rfc.execute_rfc("RFC_READ_TABLE")
    mock_functions.Add.assert_called_with("RFC_READ_TABLE")
    assert res == mock_func_obj


@patch("rpa_sap.core.rfc.pythoncom.CoInitialize")
@patch("rpa_sap.core.rfc.win32com.client.Dispatch")
def test_rfc_connection_saplogon_ini_string(mock_dispatch, mock_coinit):
    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_connection = mock_functions.Connection
    mock_connection.Logon.return_value = True

    RfcConnection(
        connection_string="022 BPE Test",
        user_id="test_user",
        password="test_password",
        client="900",
        language="EN",
    )

    assert mock_connection.System == "022 BPE Test"
    assert mock_connection.UseSAPLogonIni is True
    assert mock_connection.Client == "900"
    assert mock_connection.User == "test_user"
    assert mock_connection.Password == "test_password"
    assert mock_connection.Language == "EN"
    mock_connection.Logon.assert_called_with(0, True)


@patch("rpa_sap.core.rfc.pythoncom.CoInitialize")
@patch("rpa_sap.core.rfc.win32com.client.Dispatch")
def test_read_table(mock_dispatch, mock_coinit):
    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_functions.Connection.Logon.return_value = True

    mock_rfc = MagicMock()
    mock_functions.Add.return_value = mock_rfc
    mock_rfc.Call = True

    # Mock FIELDS table
    mock_fields_table = MagicMock()
    mock_fields_table.Data = [("FIELD1",), ("FIELD2",)]

    # Mock DATA table
    mock_data_table = MagicMock()
    mock_data_table.RowCount = 2
    mock_data_table.Data = [("VAL1;VAL2",), ("VAL3;VAL4",)]

    def side_effect(table_name):
        if table_name == "FIELDS":
            return mock_fields_table
        elif table_name == "DATA":
            return mock_data_table
        return MagicMock()

    mock_rfc.Tables = side_effect

    rfc = RfcConnection("022 BPE Test", "test_user", "test_password")
    results = rfc.read_table("T000")

    assert len(results) == 2
    assert results[0] == {"FIELD1": "VAL1", "FIELD2": "VAL2"}
    assert results[1] == {"FIELD1": "VAL3", "FIELD2": "VAL4"}
