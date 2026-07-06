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
```

### 3. Extracting Data from GridView
```python
from rpa_sap.lib.GridView import GridView

# Initialize the GridView helper with the active session
grid = GridView(session)

# Extract data directly to a Pandas DataFrame by providing the element ID
df = grid.to_DataFrame("wnd[0]/usr/cntlGRID1/shellcont/shell")
print(df.head())
```

### 4. Reading a Transparent Table via RFC (Using active session)
```python
# Assuming you have an active session
# The session object provides a built-in method to read SAP transparent tables via RFC
results = session.read_table(
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

# Initialize a headless RFC connection
rfc = RfcConnection(
    connection_string="My SAP System",
    user_id="user_id",
    password="password",
    client="100",
    language="EN"
)

# Read a transparent table
results = rfc.read_table("T000")
print(results)

# Always close the connection when done
rfc.close()
```

### 5. Working with Table Controls
```python
from rpa_sap.lib.GuiTableControl import GuiTableControl

# Initialize the TableControl helper
table = GuiTableControl(session)

# Extract the entire Table Control to a Pandas DataFrame
df = table.to_DataFrame("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY")
print(df.head())

# Set a specific cell value
table.set_cell_value(
    field_id="wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY", 
    value="100", 
    absolute_row_index=0, 
    column_title="Order Quantity"
)
```

### 6. Automating SAP Queries (SQ01)
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

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing
Contributions are welcome! Please read the [contributing guidelines](CONTRIBUTING.md) for more details.

## Contact
For any questions or suggestions, feel free to open an issue on the [GitHub repository](https://github.com/21010/rpa-sap).
