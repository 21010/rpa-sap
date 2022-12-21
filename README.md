# RPA_SAP
Python module delivers some actions to automate SAPGUI operations (Sap Scripting)
The module is compatibile with the Robocorp.

## Installation
To install the package run:

```
pip install rpa-sap
```

## Example
### Open new SAPGUI session
```
from rpa_sap import SapGui

sapgui = SapGui()

sapgui.open_new_session(connection_string, user_id, password, client, language)
```

### Dependencies
Python packages: pandas >= 1.4.4, pywin32 >= 303
