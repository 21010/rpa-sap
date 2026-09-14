from unittest.mock import Mock
import pytest

from rpa_sap import SQ01
from rpa_sap.lib.SQ01 import SQ01Locators


@pytest.fixture
def mock_sap_session():
    session = Mock()
    session.interactor = Mock()
    return session


def test_start_query_invalid_query_area(mock_sap_session):
    sq01 = SQ01(mock_sap_session)
    with pytest.raises(ValueError, match="Invalid query_area: 'Invalid'. Must be 'Standard' or 'Global'"):
        sq01.start_query("QUERY_NAME", query_area="Invalid")  # type: ignore


def test_start_query_valid_query_area(mock_sap_session):
    sq01 = SQ01(mock_sap_session)
    sq01.start_query("QUERY_NAME", query_area="Global")
    
    mock_sap_session.interactor.select.assert_any_call(SQ01Locators.MENU_QUERY_AREAS)
    mock_sap_session.interactor.select.assert_any_call(SQ01Locators.RAD_GLOBAL_AREA)
    mock_sap_session.interactor.press_button.assert_any_call(SQ01Locators.BTN_CHOOSE)


def test_to_local_file_encoding(mock_sap_session):
    sq01 = SQ01(mock_sap_session)
    
    # Needs to mock status bar to avoid exception
    mock_statusbar = Mock()
    mock_statusbar.text = "Download successful"
    mock_sap_session.interactor.get_status_bar_message.return_value = mock_statusbar
    
    sq01.to_local_file("C:\\temp", "data.csv", file_type="csv")
    
    mock_sap_session.interactor.set_text.assert_any_call(SQ01Locators.CTXT_FILE_ENCODING, "0004")

