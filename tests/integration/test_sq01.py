from rpa_sap import SQ01


def test_sq01(sap_session):
    sq01 = SQ01(sap_session)
    sq01.start_query("RPA_SOFI", query_area=None, user_group="RPA")
    assert True


def test_sq01_query_area(sap_session):
    sq01 = SQ01(sap_session)
    sq01.start_query(
        query_name="DEMO_01",
        query_area="Global",
        user_group="QDEMO",
        variant_name="SAP&STANDARD",
    )
