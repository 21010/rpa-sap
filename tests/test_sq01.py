from rpa_sap import SQ01


def test_sq01(sap_session):
    sq01 = SQ01(sap_session)
    sq01.start_query("RPA_SOFI", "RPA")
    assert True
