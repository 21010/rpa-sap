import pytest
from unittest.mock import MagicMock, patch
from rpa_sap.core.session import SapSession


@pytest.fixture
def mock_com_session():
    return MagicMock()


@pytest.fixture
def sap_session(mock_com_session):
    return SapSession(mock_com_session)


@patch("rpa_sap.core.session.win32com.client.Dispatch")
def test_execute_rfc_with_dynamic_string(mock_dispatch, sap_session):
    sap_session._rfc_credentials = {
        "connection_string": "/H/my-server/S/3200",
        "client": "900",
        "user_id": "test_user",
        "password": "test_password",
    }

    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_connection = mock_functions.Connection
    mock_connection.Logon.return_value = True

    rfc = sap_session.execute_rfc("RFC_READ_TABLE")

    mock_dispatch.assert_called_with("SAP.Functions")
    assert mock_connection.ApplicationServer == "my-server"
    assert mock_connection.SystemNumber == "00"
    assert mock_connection.UseSAPLogonIni is False
    assert mock_connection.Client == "900"
    assert mock_connection.User == "test_user"
    assert mock_connection.Password == "test_password"
    mock_connection.Logon.assert_called_with(0, True)

    mock_functions.Add.assert_called_with("RFC_READ_TABLE")
    assert rfc == mock_functions.Add.return_value


@patch("rpa_sap.core.session.win32com.client.Dispatch")
def test_execute_rfc_with_saplogon_ini_string(mock_dispatch, sap_session):
    sap_session._rfc_credentials = {
        "connection_string": "022 BPE Test",
        "client": "900",
        "user_id": "test_user",
        "password": "test_password",
    }

    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_connection = mock_functions.Connection
    mock_connection.Logon.return_value = True

    rfc = sap_session.execute_rfc("RFC_READ_TABLE")

    assert mock_connection.System == "022 BPE Test"
    assert mock_connection.UseSAPLogonIni is True
    mock_connection.Logon.assert_called_with(0, True)

    mock_functions.Add.assert_called_with("RFC_READ_TABLE")
    assert rfc == mock_functions.Add.return_value


@patch("rpa_sap.core.session.win32com.client.Dispatch")
def test_read_table(mock_dispatch, sap_session):
    # Setup mock execute_rfc by providing valid credentials
    sap_session._rfc_credentials = {
        "connection_string": "022 BPE Test",
        "client": "900",
        "user_id": "test_user",
        "password": "test_password",
    }

    mock_functions = MagicMock()
    mock_dispatch.return_value = mock_functions
    mock_functions.Connection.Logon.return_value = True

    mock_rfc = MagicMock()
    mock_functions.Add.return_value = mock_rfc

    # Mock successful call
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

    results = sap_session.read_table("T000")

    assert len(results) == 2
    assert results[0] == {"FIELD1": "VAL1", "FIELD2": "VAL2"}
    assert results[1] == {"FIELD1": "VAL3", "FIELD2": "VAL4"}
