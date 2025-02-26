# RPA_SAP
Python module delivers some actions to automate SAPGUI operations (Sap Scripting). The module is compatible with Robocorp.

## Installation
To install the package, run:

```sh
pip install rpa-sap
```

## Example
### Open new SAPGUI session
```python
from rpa_sap import SapGui

sapgui = SapGui()

sapgui.open_new_session(connection_string, user_id, password, client, language)
```

## Features
- Open and close SAP sessions
- Run SAP transactions
- Interact with SAP GUI elements (buttons, text fields, checkboxes, etc.)
- Handle SAP GridView and TableControl
- Export data to various formats (CSV, Excel)

## Dependencies
- pandas >= 1.4.4
- pywin32 >= 303
- wmi >= 1.5.1

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing
Contributions are welcome! Please read the [contributing guidelines](https://github.com/21010/rpa-sap/blob/main/CONTRIBUTING.md) for more details.

## Contact
For any questions or suggestions, feel free to open an issue on the [GitHub repository](https://github.com/21010/rpa-sap).
