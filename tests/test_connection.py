import pytest


@pytest.fixture(autouse=True)
def clean_sap_state(connection_manager):
    """Ensure SAP state is clean before each test that doesn't use the sap_session fixture"""
    pass


def test_open_and_close_new_session(connection_manager, secrets):
    connection_manager.close_sap_logon()
    session = connection_manager.open_new_session(
        secrets["connection_string"],
        secrets["user_id"],
        secrets["password"],
        secrets["client"],
        secrets["language"],
    )
    assert session is not None
    connection_manager.close_session(session)
    connection_manager.close_sap_logon()


def test_close_sap_session(sap_session, connection_manager):
    connection_manager.close_session(sap_session)


def test_activate_session(connection_manager, secrets):
    connection_manager.close_sap_logon()
    connection_manager.open_new_session(
        secrets["connection_string"],
        secrets["user_id"],
        secrets["password"],
        secrets["client"],
        secrets["language"],
    )

    # Try to activate the current session
    activated_session = connection_manager.activate_session(0, 0)
    assert activated_session is not None
    assert activated_session.info["is_active"] is True
    connection_manager.close_sap_logon()


def test_check_if_session_exists(connection_manager, secrets):
    connection_manager.close_sap_logon()
    connection_manager.open_new_session(
        secrets["connection_string"],
        secrets["user_id"],
        secrets["password"],
        secrets["client"],
        secrets["language"],
    )
    assert connection_manager.check_if_session_exists() is True
    assert connection_manager.check_if_session_exists(0) is True
    assert connection_manager.check_if_session_exists(1) is False

    connection_manager.close_all_sessions()
    connection_manager.close_sap_logon()
    assert connection_manager.check_if_session_exists() is False
    assert connection_manager.check_if_session_exists(0) is False


def test_get_session_info(sap_session):
    session_info = sap_session.info
    print(session_info)
    assert isinstance(session_info, dict)
    assert session_info["user"] != ""


def test_count_connections(connection_manager, secrets):
    connection_manager.close_sap_logon()
    connection_manager.open_new_session(
        secrets["connection_string"],
        secrets["user_id"],
        secrets["password"],
        secrets["client"],
        secrets["language"],
    )
    assert connection_manager.connections.Count == 1

    connection_manager.open_new_session(
        secrets["connection_string"],
        secrets["user_id"],
        secrets["password"],
        secrets["client"],
        secrets["language"],
    )
    assert connection_manager.connections.Count == 2
    connection_manager.close_sap_logon()


def test_is_active(sap_session):
    import time

    time.sleep(2)
    assert sap_session.info["is_active"] is True


def test_wait_until_object_exists(sap_session):
    sap_session.run_transaction("su3")
    exists = sap_session.interactor.wait_until_object_exists("wnd[0]")
    assert exists is True


def test_close_process(connection_manager):
    connection_manager.close_process(process_name="excel.exe")
    assert True
