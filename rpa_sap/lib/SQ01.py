
# from SapGui import SapGui

class SQ01():
    def __init__(self, sap_gui):
        self.sap_gui = sap_gui
    
    def start_query(self, query_name: str, user_group: str = None, variant_name: str = None):
        # Navigate to the transaction code for SQ01
        self.sap_gui.run_transaction('SQ01')

        # Change User Group if needed
        if user_group:
            # press button Change User Group
            self.sap_gui.press_button('wnd[0]/tbar[1]/btn[19]')

            # press Filter button
            self.sap_gui.press_button('wnd[1]/tbar[0]/btn[29]')

            # Select Name
            self.sap_gui.set_property('wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell', 'selectedRows', 0)

            # Press < button
            self.sap_gui.press_button("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING")

            # Press Set Value button
            self.sap_gui.press_button("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON")

            # Enter User Group name
            self.sap_gui.set_text("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW", user_group)

            #Press Search Button
            self.sap_gui.press_button("wnd[3]/tbar[0]/btn[0]")

            # Select User Group row
            try:
                self.sap_gui.set_property("wnd[1]/usr/cntlGRID1/shellcont/shell", 'selectedRows', 0)
            except Exception as ex:
                raise Exception(f"User Group {user_group} not found. Error: {ex}")

            # Press OK button
            self.sap_gui.press_button("wnd[1]/tbar[0]/btn[0]")

        # Enter the query name
        self.sap_gui.set_text('wnd[0]/usr/ctxtRS38R-QNUM', query_name)

        if variant_name:
            # Press Execute with Variants button
            self.sap_gui.press_button('wnd[0]/tbar[1]/btn[17]')
            # Enter variant name
            self.sap_gui.set_text('wnd[1]/usr/ctxtRS38R-VARIANT', variant_name)
            # Press OK button
            self.sap_gui.press_button('wnd[1]/tbar[0]/btn[0]')
            # Check if no variant found error message is displayed
            statusbar = self.sap_gui.get_status_bar_message()
            if statusbar.text.contains(f'Variant {variant_name} does not exist'):
                raise Exception(f'Error: Variant {variant_name} does not exist')
        else:
            # Press Execute button
            self.sap_gui.press_button('wnd[0]/tbar[1]/btn[8]')
        
        # Verify statusbar
        statusbar = self.sap_gui.get_status_bar_message()
        if statusbar.type == 'E':
            raise Exception(f'Error: {statusbar.text}')

    def execute_query(self, time_limit: int = 300):
        self.sap_gui.press_button('wnd[0]/tbar[1]/btn[8]')

    def to_local_file(self, folder_path: str, file_name: str, file_type: str = 'xls'):
        # Press Export button
        self.sap_gui.press_button('wnd[0]/tbar[1]/btn[45]')
        # Select export type
        self.sap_gui.select('wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]')
        # Press OK button
        self.sap_gui.press_button('wnd[1]/tbar[0]/btn[0]')
        # enter folder path
        self.sap_gui.set_text('wnd[1]/usr/ctxtDY_PATH', folder_path)
        # enter file name
        self.sap_gui.set_text('wnd[1]/usr/ctxtDY_FILENAME', file_name)
        # set code pate
        self.sap_gui.set_text('wnd[1]/usr/ctxtDY_FILE_ENCODING', '0000')
        # Press OK button
        self.sap_gui.press_button('wnd[1]/tbar[0]/btn[11]')
        # Verify status bar
        statusbar = self.sap_gui.get_status_bar_message()
        if not statusbar.text.contains('Download'):
            raise Exception(f'Data has not been exported sucessully')
