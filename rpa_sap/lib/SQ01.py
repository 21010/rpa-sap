from ..core.session import SapSession
from ..exceptions import SapElementNotFoundError


class SQ01Locators:
    # User Group selection
    BTN_CHANGE_USER_GROUP = "wnd[0]/tbar[1]/btn[19]"
    BTN_FILTER = "wnd[1]/tbar[0]/btn[29]"
    SHELL_FILTER_CONTAINER = (
        "wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell"
    )
    BTN_ADD_TO_SELECTION = "wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING"
    BTN_SET_VALUE = "wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON"
    CTXT_USER_GROUP_NAME = (
        "wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW"
    )
    BTN_SEARCH_USER_GROUP = "wnd[3]/tbar[0]/btn[0]"
    SHELL_USER_GROUP_GRID = "wnd[1]/usr/cntlGRID1/shellcont/shell"
    BTN_USER_GROUP_OK = "wnd[1]/tbar[0]/btn[0]"

    # Query execution
    CTXT_QUERY_NAME = "wnd[0]/usr/ctxtRS38R-QNUM"
    BTN_EXECUTE_WITH_VARIANTS = "wnd[0]/tbar[1]/btn[17]"
    CTXT_VARIANT_NAME = "wnd[1]/usr/ctxtRS38R-VARIANT"
    BTN_VARIANT_OK = "wnd[1]/tbar[0]/btn[0]"
    BTN_EXECUTE = "wnd[0]/tbar[1]/btn[8]"

    # Exporting
    BTN_EXPORT = "wnd[0]/tbar[1]/btn[45]"
    RAD_EXPORT_SPREADSHEET = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
    BTN_EXPORT_OK = "wnd[1]/tbar[0]/btn[0]"
    CTXT_DIRECTORY_PATH = "wnd[1]/usr/ctxtDY_PATH"
    CTXT_FILE_NAME = "wnd[1]/usr/ctxtDY_FILENAME"
    CTXT_FILE_ENCODING = "wnd[1]/usr/ctxtDY_FILE_ENCODING"
    BTN_SAVE_FILE = "wnd[1]/tbar[0]/btn[11]"


class SQ01:
    def __init__(self, sap_session: SapSession):
        self.sap_session = sap_session
        self.interactor = sap_session.interactor

    def start_query(
        self, query_name: str, user_group: str = None, variant_name: str = None
    ):
        # Navigate to the transaction code for SQ01
        self.sap_session.run_transaction("SQ01")

        # Change User Group if needed
        if user_group:
            self.interactor.press_button(SQ01Locators.BTN_CHANGE_USER_GROUP)
            self.interactor.press_button(SQ01Locators.BTN_FILTER)
            self.interactor.set_property(
                SQ01Locators.SHELL_FILTER_CONTAINER,
                "selectedRows",
                0,
            )
            self.interactor.press_button(SQ01Locators.BTN_ADD_TO_SELECTION)
            self.interactor.press_button(SQ01Locators.BTN_SET_VALUE)
            self.interactor.set_text(
                SQ01Locators.CTXT_USER_GROUP_NAME,
                user_group,
            )
            self.interactor.press_button(SQ01Locators.BTN_SEARCH_USER_GROUP)

            try:
                self.interactor.set_property(
                    SQ01Locators.SHELL_USER_GROUP_GRID, "selectedRows", 0
                )
            except Exception as ex:
                raise SapElementNotFoundError(
                    f"User Group {user_group} not found. Error: {ex}"
                )

            self.interactor.press_button(SQ01Locators.BTN_USER_GROUP_OK)

        # Enter the query name
        self.interactor.set_text(SQ01Locators.CTXT_QUERY_NAME, query_name)

        if variant_name:
            self.interactor.press_button(SQ01Locators.BTN_EXECUTE_WITH_VARIANTS)
            self.interactor.set_text(SQ01Locators.CTXT_VARIANT_NAME, variant_name)
            self.interactor.press_button(SQ01Locators.BTN_VARIANT_OK)

            # Check if no variant found error message is displayed
            statusbar = self.interactor.get_status_bar_message()
            if f"Variant {variant_name} does not exist" in statusbar.text:
                raise SapElementNotFoundError(
                    f"Error: Variant {variant_name} does not exist"
                )
        else:
            self.interactor.press_button(SQ01Locators.BTN_EXECUTE)

        # Verify statusbar
        statusbar = self.interactor.get_status_bar_message()
        if statusbar.type == "E":
            raise Exception(f"Error: {statusbar.text}")

    def execute_query(self):
        self.interactor.press_button(SQ01Locators.BTN_EXECUTE)

    def to_local_file(self, folder_path: str, file_name: str, file_type: str = "xls"):
        self.interactor.press_button(SQ01Locators.BTN_EXPORT)
        self.interactor.select(SQ01Locators.RAD_EXPORT_SPREADSHEET)
        self.interactor.press_button(SQ01Locators.BTN_EXPORT_OK)

        self.interactor.set_text(SQ01Locators.CTXT_DIRECTORY_PATH, folder_path)
        self.interactor.set_text(SQ01Locators.CTXT_FILE_NAME, file_name)

        encoding = (
            "0000" if file_type == "xls" else "0004" if file_type == "csv" else "0000"
        )
        self.interactor.set_text(SQ01Locators.CTXT_FILE_ENCODING, encoding)

        self.interactor.press_button(SQ01Locators.BTN_SAVE_FILE)

        # Verify status bar
        statusbar = self.interactor.get_status_bar_message()
        if "Download" not in statusbar.text:
            raise Exception("Data has not been exported successfully")
