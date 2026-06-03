import os
import pytest
from dotenv import load_dotenv
from rpa_sap import ConnectionManager

load_dotenv()


def pytest_runtest_setup(item):
    if os.getenv("CI") == "true" and item.get_closest_marker("integration"):
        pytest.skip(
            "Skipping SAP GUI tests in CI environment because SAP GUI is not installed."
        )


def pytest_collection_modifyitems(items):
    for item in items:
        if "integration" in str(item.fspath):
            item.add_marker(pytest.mark.integration)


@pytest.fixture(scope="session")
def secrets():
    # Return a dictionary of secrets mapping to old json format
    return {
        "connection_string": os.getenv("SAP_CONNECTION_STRING", "/H/server/S/3200"),
        "user_id": os.getenv("SAP_USER_ID", "user"),
        "password": os.getenv("SAP_PASSWORD", "pass"),
        "client": os.getenv("SAP_CLIENT", "900"),
        "language": os.getenv("SAP_LANGUAGE", "EN"),
    }


@pytest.fixture(scope="module")
def connection_manager():
    manager = ConnectionManager()
    yield manager
    try:
        manager.close_all_sessions()
        manager.close_sap_logon()
    except Exception:
        pass


@pytest.fixture
def sap_session(connection_manager, secrets):
    session = connection_manager.open_new_session(
        secrets["connection_string"],
        secrets["user_id"],
        secrets["password"],
        secrets["client"],
        secrets["language"],
    )
    yield session
    try:
        connection_manager.close_session(session)
    except Exception:
        pass
