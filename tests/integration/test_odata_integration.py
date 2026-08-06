import os
import pytest


@pytest.fixture
def live_odata_client():
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass

    odata_url = os.getenv("SAP_ODATA_URL")
    odata_user = os.getenv("SAP_ODATA_USER")
    odata_password = os.getenv("SAP_ODATA_PASSWORD")

    if (
        not all([odata_url, odata_user, odata_password])
        or odata_user == "your_comm_user"
    ):
        pytest.skip("OData credentials not properly configured in .env")

    from rpa_sap.core.odata import ODataClient, BasicAuthStrategy
    auth = BasicAuthStrategy(str(odata_user), str(odata_password))
    return ODataClient(str(odata_url), auth)


@pytest.mark.integration
def test_live_get_business_partners(live_odata_client):
    """
    Integration test connecting to a live S/4HANA Cloud environment.
    This assumes SAP_ODATA_URL points to the API_BUSINESS_PARTNER OData service.
    """
    try:
        # We query the A_BusinessPartner entity set
        df = live_odata_client.get_dataframe(
            "A_BusinessPartner",
            select=["BusinessPartner", "BusinessPartnerFullName"],
            top=5,
        )

        assert not df.empty, "Expected at least one Business Partner, but got none."
        assert "BusinessPartner" in df.columns, (
            "BusinessPartner column missing from results."
        )
        assert len(df) <= 5, "Received more than the requested $top=5 records."
    except Exception as e:
        from rpa_sap.exceptions import SapODataError
        if isinstance(e, SapODataError):
            pytest.fail(f"OData integration test failed: {e}")
        raise e
