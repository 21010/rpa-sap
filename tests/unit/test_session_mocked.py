import pytest
from unittest.mock import MagicMock, patch
from rpa_sap.core.session import SapSession


@pytest.fixture
def mock_com_session():
    return MagicMock()


@pytest.fixture
def sap_session(mock_com_session):
    return SapSession(mock_com_session)


def test_rfc_property(sap_session):
    sap_session._rfc_credentials = {
        "connection_string": "022 BPE Test",
        "client": "900",
        "user_id": "test_user",
        "password": "test_password",
        "language": "EN"
    }

    with patch("rpa_sap.core.session.RfcConnection") as mock_rfc_class:
        rfc = sap_session.rfc
        mock_rfc_class.assert_called_with(**sap_session._rfc_credentials)
        assert rfc == mock_rfc_class.return_value

        # Calling again should return cached instance
        rfc2 = sap_session.rfc
        assert rfc2 == rfc
        mock_rfc_class.assert_called_once()
