# SAP GUI & RFC Automation
[![CI](https://github.com/21010/rpa-sap/actions/workflows/ci.yml/badge.svg)](https://github.com/21010/rpa-sap/actions/workflows/ci.yml)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![SAP GUI](https://img.shields.io/badge/SAP-GUI_Scripting-008FD3.svg?logo=sap&logoColor=white)](https://www.sap.com/)
[![Robocorp Ready](https://img.shields.io/badge/Robocorp-Ready-blueviolet.svg)](https://robocorp.com/)
[![RPA Automation](https://img.shields.io/badge/RPA-Automation-brightgreen.svg)](#)

A comprehensive Python module for automating SAP GUI operations (SAP Scripting) and RFC (Remote Function Call) integration. Designed to easily integrate with modern RPA frameworks (e.g., Robocorp) and standard Python automation scripts.

## Business Benefits
- **Increased Efficiency:** Automate repetitive manual data entry, extraction tasks, and transactions in SAP without human intervention.
- **Reduced Errors:** Eliminate human error in routine transactions by relying on precise GUI element interaction and structured RFC calls.
- **Scalability:** Seamlessly integrate SAP processes into larger, multi-system RPA orchestrations.
- **Data Export & Transformation:** Effortlessly extract SAP GridView and TableControl data directly into modern formats like Pandas DataFrames, CSV, and Excel for immediate downstream analytics.

## Architecture & Structure
The project is modularized into core interaction capabilities and higher-level automation utilities:

- `rpa_sap/core/`: Contains the foundational layers for SAP interaction.
  - `connection.py`: Manages SAP Logon connections, active process handling, and explicit logins.
  - `session.py`: Handles individual SAP sessions and RFC execution wrappers.
  - `ui_automation.py`: Provides direct interaction capabilities with standard SAP GUI elements (buttons, fields, trees, dialogs).

- `rpa_sap/lib/`: Advanced handlers for complex SAP controls and specific workflows.
  - `GridView.py`: Utilities for reading, scrolling, and extracting data from SAP GridView components.
  - `GuiTableControl.py`: Utilities for managing standard SAP Table Controls.
  - `SQ01.py`: Automation features for SAP query execution workflows (SQ01).

## Installation
To install the package, run:

```sh
pip install rpa-sap
```

## Requirements & Dependencies
- Python >= 3.10
- SAP GUI Scripting must be enabled on the server and client.
- Dependencies: `pandas`, `pywin32`, `wmi`, `python-dotenv`, `openpyxl`

## How to Use (Examples)

### 1. Opening a New SAPGUI Session
The `connection_string` should be the exact name of the connection as defined in your SAP Logon application (e.g., `"S4HANA Prod"`), or a direct SAP connection string (e.g., `"/H/192.168.1.1/S/3200"`).

```python
from rpa_sap import ConnectionManager

manager = ConnectionManager()
session = manager.open_new_session(
    connection_string="My SAP System", # Exact name from SAP Logon
    user_id="user_id", 
    password="password", 
    client="100", 
    language="EN"
)
```

### 2. Interacting with UI Elements
The `session.interactor` exposes an `ElementInteractor` which provides high-level, robust methods to manipulate SAP GUI objects easily.

```python
# Assuming you have an active session
# Set a transaction code (e.g., MM03) in the command field
session.interactor.set_text("wnd[0]/tbar[0]/okcd", "MM03")

# Press the Enter key
session.interactor.press_enter()

# Press a specific button by its ID
session.interactor.press_button("wnd[0]/tbar[1]/btn[8]")

# Check or uncheck a checkbox
session.interactor.set_checkbox_state("wnd[0]/usr/chk[1,1]", True)

# FindElemendById is also available via session object.
session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
session.findById("wnd[0]").sendVKey(0)

# You can also use the transaction context manager to ensure safe cleanup
with session.transaction("ME32L"):
    session.interactor.set_text("wnd[0]/usr/ctxtRM06E-EBELN", "4500000001")
    session.interactor.press_enter()
```

### 3. Extracting Data from GridView
```python
from rpa_sap.lib.GridView import GridView
from rpa_sap.lib.data_extractors import GridViewExtractor

# Initialize the GridView helper with the active session
grid = GridView(session)
extractor = GridViewExtractor(grid)

# Extract data directly to a Pandas DataFrame by providing the element ID
df = extractor.to_dataframe("wnd[0]/usr/cntlGRID1/shellcont/shell")
print(df.head())
```

### 4. Reading a Transparent Table via RFC (Using active session)
```python
# Assuming you have an active session
# The session object provides an embedded RFC connection via the `.rfc` property
results = session.rfc.read_table(
    table_name="T000",
    fields=["MANDT", "MTEXT"],
    options=["SPRAS = 'E'"]
)
print(results)
```

### 5. Fully Headless RFC Connection
For scenarios where you do not need an active SAP GUI session and want to interact purely via RFC, you can use the `RfcConnection` class directly.

```python
from rpa_sap import RfcConnection

# Initialize a headless RFC connection using a context manager
with RfcConnection(
    connection_string="My SAP System",
    user_id="user_id",
    password="password",
    client="100",
    language="EN"
) as rfc:
    # Read a transparent table
    results = rfc.read_table("T000")
    print(results)
    # The connection is automatically closed when the block exits
```

### 6. Executing a BAPI via RFC
You can execute BAPI functions cleanly and extract exactly the parameters or tables you need.

```python
from rpa_sap import RfcConnection

with RfcConnection(
    connection_string="My SAP System",
    user_id="user_id",
    password="password"
) as rfc:
    results = rfc.call_bapi(
        bapi_name="BAPI_USER_GET_DETAIL",
        import_params={"USERNAME": "USERNAME"},
        extract_imports=["ADDRESS"],
        extract_tables=["ACTIVITYGROUPS"]
    )

    # Access the returned structures and tables
    address_data = results.get("ADDRESS")
    roles = results.get("ACTIVITYGROUPS")

    print(f"User Full Name: {address_data.get('FULLNAME')}")
    print(f"Number of roles: {len(roles)}")
```

### 7. Working with Table Controls
```python
from rpa_sap.lib.GuiTableControl import GuiTableControl
from rpa_sap.lib.data_extractors import GuiTableControlExtractor

# Initialize the TableControl helper
table = GuiTableControl(session)
extractor = GuiTableControlExtractor(table)

# Extract the entire Table Control to a Pandas DataFrame
df = extractor.to_dataframe("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY")
print(df.head())

# Set a specific cell value
table.set_cell_value(
    field_id="wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY", 
    value="100", 
    absolute_row_index=0, 
    column_title="Order Quantity"
)
```

### 8. Automating SAP Queries (SQ01)
```python
from rpa_sap.lib.SQ01 import SQ01

# Initialize SQ01 helper
sq01 = SQ01(session)

# Navigate to the query, providing a user group and variant
sq01.start_query(query_name="MY_QUERY", user_group="MY_GROUP", variant_name="DEFAULT")

# Execute the query
sq01.execute_query()

# Export the results directly to a local file
sq01.to_local_file(folder_path="C:\\Exports", file_name="query_results.xls", file_type="xls")
```

### 9. Integrating via SAP OData (REST)
For modern SAP systems (like S/4HANA), OData is the preferred integration method. RPA-SAP provides an `ODataClient` with various authentication strategies (Basic, OAuth2, etc.) and handles CSRF tokens automatically for state-changing operations.

```python
from rpa_sap.core.odata import ODataClient, BasicAuthStrategy

# 1. Choose your authentication strategy
auth = BasicAuthStrategy("USERNAME", "password")
# Or use OAuth2: auth = OAuth2Strategy("my-bearer-token")

# 2. Initialize the client
client = ODataClient("https://mysap.example.com/sap/opu/odata/sap/API_USER_SRV", auth)

# 3. Query an EntitySet and get a Pandas DataFrame
df = client.get_dataframe("UserSet", select=["UserID", "FullName"], top=50)
print(df.head())

# 4. Create a new Entity (CSRF token is fetched automatically)
new_user = {"UserID": "NEW_RPA", "FullName": "RPA Bot User"}
response = client.post("UserSet", payload=new_user)
print("Created:", response)
```

## Testing

The project uses `pytest` and features a two-tier testing strategy:

1. **Unit Tests (Fast & CI-Friendly):** Located in `tests/unit/`. These tests mock the SAP GUI COM objects and run without requiring a live SAP installation or active connection.
   ```sh
   uv run pytest tests/unit
   ```

2. **Integration Tests (Live SAP Environment):** Located in `tests/integration/`. These tests connect to a live SAP environment and require SAP GUI to be installed, running, and accessible. They are marked with `@pytest.mark.integration`.
   ```sh
   uv run pytest -m "integration"
   ```

## Changelog & Recent Updates
- **Stability Fix**: Fixed `Windows Fatal COM exceptions (0x80010108, 0x800706ba)` that occurred on session closure. The library now explicitly detaches COM proxies and forces garbage collection before terminating the SAP logon process.
- **Type Safety**: Full codebase audit using `pyrefly`. Resolved all static analysis errors, standardizing type hints and dependency injection (`BaseMixin`) across core components.
- **Test Discoverability**: Improved pytest integration with IDEs (like VSCode) by gracefully skipping module imports (`pytest.importorskip`) during test collection when dependencies are missing.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing
Contributions are welcome! Please read the [contributing guidelines](CONTRIBUTING.md) for more details.

## Contact
For any questions or suggestions, feel free to open an issue on the [GitHub repository](https://github.com/21010/rpa-sap).
