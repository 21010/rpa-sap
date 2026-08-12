import pytest

requests_mock = pytest.importorskip("requests_mock")
pd = pytest.importorskip("pandas")
pytest.importorskip("rpa_sap")

from rpa_sap.core.odata import (  # noqa: E402
    ODataClient,
    BasicAuthStrategy,
    OAuth2Strategy,
    NoAuthStrategy,
)
from rpa_sap.exceptions import SapODataError  # noqa: E402

BASE_URL = "https://mysap.example.com/sap/opu/odata/sap/API_USER_SRV"


@pytest.fixture
def odata_client():
    auth = BasicAuthStrategy("testuser", "testpass")
    return ODataClient(BASE_URL, auth)


def test_auth_strategies():
    # Test Basic Auth
    basic = BasicAuthStrategy("user", "pass")
    client1 = ODataClient("http://test", basic)
    assert client1.session.auth == ("user", "pass")

    # Test OAuth2
    oauth = OAuth2Strategy("my-token")
    client2 = ODataClient("http://test", oauth)
    assert client2.session.headers["Authorization"] == "Bearer my-token"

    # Test No Auth
    noauth = NoAuthStrategy()
    client3 = ODataClient("http://test", noauth)
    assert client3.session.auth is None
    assert "Authorization" not in client3.session.headers


def test_fetch_csrf_token(odata_client):
    with requests_mock.Mocker() as m:
        m.get(BASE_URL, headers={"x-csrf-token": "abc-123"})
        token = odata_client.fetch_csrf_token()
        assert token == "abc-123"
        assert odata_client.csrf_token == "abc-123"


def test_fetch_csrf_token_failure(odata_client):
    with requests_mock.Mocker() as m:
        m.get(BASE_URL, status_code=403)
        with pytest.raises(SapODataError, match="Error fetching CSRF token"):
            odata_client.fetch_csrf_token()


def test_get_entity_set(odata_client):
    with requests_mock.Mocker() as m:
        mock_response = {
            "d": {
                "results": [
                    {"UserID": "JOHN", "Name": "John Doe"},
                    {"UserID": "JANE", "Name": "Jane Doe"},
                ]
            }
        }
        m.get(
            f"{BASE_URL}/UserSet?$format=json&$select=UserID,Name&$top=2",
            json=mock_response,
        )

        results = odata_client.get("UserSet", select=["UserID", "Name"], top=2)
        assert len(results) == 2
        assert results[0]["UserID"] == "JOHN"


def test_get_dataframe(odata_client):
    with requests_mock.Mocker() as m:
        mock_response = {"d": {"results": [{"UserID": "JOHN", "Name": "John Doe"}]}}
        m.get(f"{BASE_URL}/UserSet?$format=json", json=mock_response)

        df = odata_client.get_dataframe("UserSet")
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 1
        assert df.iloc[0]["UserID"] == "JOHN"


def test_post_entity(odata_client):
    with requests_mock.Mocker() as m:
        # Mock CSRF fetch
        m.get(BASE_URL, headers={"x-csrf-token": "token-123"})

        # Mock POST
        mock_post_response = {"d": {"UserID": "NEWUSER", "Name": "New User"}}
        m.post(f"{BASE_URL}/UserSet", json=mock_post_response)

        payload = {"UserID": "NEWUSER", "Name": "New User"}
        result = odata_client.post("UserSet", payload)

        assert result["UserID"] == "NEWUSER"
        assert odata_client.csrf_token == "token-123"
        assert m.last_request.headers["X-CSRF-Token"] == "token-123"
