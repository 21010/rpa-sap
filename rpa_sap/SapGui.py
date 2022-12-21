"""
RPA toolchain to automate SapGui (Sap Scripting)
"""

from os import getlogin
import subprocess
from time import sleep
import re
from collections import namedtuple
import win32com.client
from pandas import DataFrame

StatusBar = namedtuple('StatusBar', ['text', 'type'])

class SapGui:
    """
    Allows to automate SAP Gui operations.

    Example:
        from rpa_sap import SapGui\n\r
        sap = SapGui()\n\r
        sap.open_new_session('/H/server/S/3200', 'user', 'pass', '900', 'EN')\n\r
        sap.run_transaction('sq01')\n\r
        sap.set_text('wnd[0]/tbar[0]/okcd', 'zxy')\n\r
        sap.invoke_method('wnd[0]/tbar[0]/btn[71]', 'press')\n\r
        sap.press_button('wnd[0]/tbar[0]/btn[73]')\n\r
        sap.close_session()\n\r
        sap.close_sap_logon()\n\r
    """

    __sap_gui: win32com.client.CDispatch = None
    __application: win32com.client.CDispatch = None
    active_connection: win32com.client.CDispatch = None
    active_session: win32com.client.CDispatch = None
    active_window: win32com.client.CDispatch = None
    active_object: win32com.client.CDispatch = None
    active_gridview: win32com.client.CDispatch = None

    @property
    def connections(self):
        """ Returns: Collection of all SAP connections """
        return self.__application.Connections

    @property
    def sessions(self):
        """ Returns: Collection of all sessions from active connection """
        return self.active_connection.Sessions

    # Session

    def open_new_session(self, connection_string: str, user_id: str, password: str, client: str = '900', language: str = "EN", timeout: int = 10) -> bool:
        """
        Opens and logs in to new SAP session.

        Args:
            connection_string (str): Connection string for SAP system
            user_id (str): SAP user id
            password (str): SAP password
            client (str, optional): SAP client. Defaults to '900'.
            language (str, optional): Language. Defaults to "EN".
            timeout (int, optional): Defines how many seconds wait until SAPGUI is opened. Default to 10.

        Usage:
            sap.open_new_session('/H/server/S/3200', 'user', 'pass', '900', 'EN')
            sap.open_new_session('/H/server/S/3250', 'user', 'pass')
        """
        # Run sapgui.exe with connection string as a parameter
        try:
            subprocess.check_call(['C:/Program Files (x86)/SAP/FrontEnd/SAPgui/SAPgui.exe', connection_string])
        except (subprocess.CalledProcessError, subprocess.SubprocessError) as ex:
            raise Exception from ex
        # Wait to make sure the SAPGUI is opened
        sleep(timeout)
        # Connect to the SAP session by checking the SID, Client and the User ID
        self.__sap_gui = win32com.client.GetObject('SAPGUI')
        self.__application = self.__sap_gui.GetScriptingEngine
        #self.active_connection = self.__application.Connections[self.__application.Connections.Count - 1]
        self.active_connection = self.connections[self.connections.Count - 1]
        #self.active_session = self.active_connection.Sessions[self.active_connection.Sessions.Count - 1]
        self.active_session = self.sessions[self.sessions.Count - 1]
        self.active_window = self.active_session.findById('wnd[0]')
        # Maximize window
        self.active_window.maximize()
        # Log to the SAP system
        self.active_session.findById('wnd[0]/usr/txtRSYST-BNAME').Text = user_id
        self.active_session.findById('wnd[0]/usr/pwdRSYST-BCODE').Text = password
        self.active_session.findById('wnd[0]/usr/txtRSYST-MANDT').Text = client
        self.active_session.findById('wnd[0]/usr/txtRSYST-LANGU').Text = language
        self.active_window.SendVKey(0)
        # Check if "License Information for Multiple Logon" popsup
        if self.check_if_object_exists('wnd[1]'):
            if self.check_if_object_exists('wnd[1]/usr/radMULTI_LOGON_OPT2'):
                self.select('wnd[1]/usr/radMULTI_LOGON_OPT2')
                self.press_button('wnd[1]/tbar[0]/btn[0]')
        # Read status bar type and message
        status = self.get_status_bar_message()
        if status.type == 'E':
            raise Exception(f'{status.type} : {status.text}')
        return True

    def activate_session(self, connection_index: int | None = None, session_index: int | None = None, user_id: str | None = None, sid: str | None = None, application_server: str | None = None, client: str | None = None):
        """
        Activates existing SAP session by connection index and session index or connections details.\n\r
        To use this method, the SAP session must be already established.\n\r

        
        Args:
            connection_index (int, optional): Connection index. Defaults to None.
            session_index (int, optional): Session index. Defaults to None.
            user_id (str, optional): Logged user id. Defaults to None.
            sid (str, optional): SID. Defaults to None.
            application_server (str, optional): Application server. Defaults to None.
            client (str, optional): Client number. Defaults to None.

        Usage (session index):
            sap.activate_session(session_index=0)

        Usage (connection index and session index):
            sap.activate_session(connection_index=0, session_index=0)\n\r

        Usage (connection details):
            sap.activate_session('user', 'prt', 'autoprt', '900')\n\r


        Raises:
            Exception: Cannot connect to SAPGUI session. SAPGUI seems to be not opened.
        """
        try:
            self.__sap_gui = win32com.client.GetObject('SAPGUI')
            self.__application = self.__sap_gui.GetScriptingEngine
        except Exception as ex:
            raise Exception('Cannot connect to SAPGUI session. SAPGUI seems to be not opened.') from ex
        try:
            # case connection index and session index is passed
            if connection_index is not None and session_index is not None:
                #self.active_connection = self.__application.Connections[connection_index]
                #self.active_session = self.active_connection.Sessions[session_index]
                self.active_connection = self.connections[connection_index]
                self.active_session = self.sessions[session_index]

            # case connection index is None and session index is passed
            if connection_index is None and session_index is not None:
                if self.active_connection is None:
                    self.active_connection = self.connections[self.connections.Count - 1]
                self.active_session = self.sessions[session_index]

            if connection_index is None and session_index is None:
                # case connection details are used
                if user_id is not None and sid is not None and application_server is not None and client is not None:
                    for connection in self.__application.Connections:
                        for session in connection.Sessions:
                            if session.Info.SystemName == sid.upper() and session.Info.Client == client and session.Info.User == user_id.upper() and session.Info.ApplicationServer.upper() == application_server.upper():
                                self.active_connection = connection
                                self.active_session = session
                # case when latest connection and session is used
                else:
                    #self.active_connection = self.__application.Connections[self.__application.Connections.Count - 1]
                    #self.active_session = self.active_connection.Sessions[self.active_connection.Sessions.Count - 1]
                    self.active_connection = self.connections[self.connections.Count - 1]
                    self.active_session = self.sessions[self.sessions.Count - 1]

            self.active_window = self.active_session.Children[self.active_session.Children.Count - 1]
        except Exception as ex:
            raise Exception('Cannot activate session. Please verify provided properties are correct.') from ex

    def check_if_session_exists(self, connection_index: int | None = None, session_index: int | None = None) -> bool:
        """
        Checks if SAPGUI session exists and return True or False.

        Args:
            connection_index (int | None, optional): Connection Index. Defaults to None.
            session_index (int | None, optional): Session Index. Defaults to None.

        Returns:
            bool: True if session exists, False if not.
        """
        try:
            self.__sap_gui = win32com.client.GetObject('SAPGUI')
            self.__application = self.__sap_gui.GetScriptingEngine
            con_index = connection_index if connection_index is not None else self.__application.Connections.Count - 1
            ses_index = session_index if session_index is not None else self.__application.Connections[con_index].Sessions.Count - 1
            obj = self.__application.Connections[con_index].Sessions[ses_index]
            # return True if obj is not None else False
            return obj is not None
        except Exception as ex:
            return False

    def close_session(self, connection_index: int = None, session_index: int = None):
        """
        Closes SAP session.
        If connection index is not provided the active connection will be used.
        If session index is not provided, the active session will be used.

        Args:
            connection_index (int, optional): Connection index. Defaults to None.
            session_index (int, optional): Session index. Defaults to None.
        """
        connection = self.active_connection if connection_index is None else self.connections[connection_index]
        session = self.active_session if session_index is None else connection.Sessions[session_index]
        connection.CloseSession(session.Id)

    def close_all_sessions(self):
        """
        Closes all opened SAP sessions for all opened connections.
        """
        try:
            self.__sap_gui = win32com.client.GetObject('SAPGUI')
            self.__application = self.__sap_gui.GetScriptingEngine
            for connection in self.__application.Connections:
                for session in connection.Sessions:
                    connection.CloseSession(session.Id)
        except:
            pass

    def get_session_info(self, connection_index: int = None, session_index: int = None) -> dict:
        """
        Return information about the session.

        Args:
            connection_index (int, optional): Connection index. Defaults to None.
            session_index (int, optional): Session index. Defaults to None.

        Returns:
            dict: 'is active', 'is busy', 'connection index', 'session index', 'Application Server', 'Code Page',\n
                    'Group', 'GuiCodepage', 'IsLowSpeedConnection', 'Language', 'MessageServer', 'ResponseTime',\n
                    'ScreenNumber', 'SessionNumber', 'SystemNumber', 'SystemSessionId', 'System Name', 'Client',\n
                    'User ID', 'Program', 'Transaction'
        """
        connection = self.active_connection if connection_index is None else self.connections[connection_index]
        session = self.active_session if session_index is None else connection.Sessions[session_index]
        return {
            'is active': session.IsActive,
            'is busy': session.Busy,
            'connection index': connection.Id,
            'session index': session.Id,
            'Application Server': session.Info.ApplicationServer,
            'Code Page': session.Info.Codepage,
            'Group': session.Info.Group,
            'GuiCodepage': session.Info.GuiCodepage,
            'IsLowSpeedConnection': session.Info.IsLowSpeedConnection,
            'Language': session.Info.Language,
            'MessageServer': session.Info.MessageServer,
            'ResponseTime': session.Info.ResponseTime,
            'ScreenNumber': session.Info.ScreenNumber,
            'SessionNumber': session.Info.SessionNumber,
            'SystemNumber': session.Info.SystemNumber,
            'SystemSessionId': session.Info.SystemSessionId,
            'System Name': session.Info.SystemName,
            'Client': session.Info.Client,
            'User ID': session.Info.User,
            'Program': session.Info.Program,
            'Transaction': session.Info.Transaction
        }

    def get_connection_index(self) -> int:
        """
        Returns the index of active connection

        Returns:
            int: number value
        """
        return int(re.search("[0-9]", self.active_connection.Id)[0])

    def get_session_index(self) -> int:
        """
        Returns the index of active session

        Returns:
            int: number value
        """
        return int(re.findall("[0-9]", self.active_session.Id)[-1])

    def count_connections(self) -> int:
        """
        Returns the number of opened connections.

        Returns: 
            int: number value
        """
        return self.__application.Connections.Count

    def count_sessions(self, connection_index: int = None) -> int:
        """
        Count sessions for the SAP connection.
        If connection_index is not provided the active connection is used.

        Args:
            connection_index (int, optional): Connection index. Defaults to None.

        Returns:
            int: number value
        """
        return self.active_connection.Sessions.Count if connection_index is None else self.connections[connection_index].Sessions.Count

    def is_session_busy(self, connection_index: int | None = None, session_index: int | None = None) -> bool:
        """
        Checks if SAP session is busy.

        Args:
            connection_index (int | None, optional): Connection index. Defaults to None.
            session_index (int | None, optional): Session index. Defaults to None.

        Returns:
            bool: True if session is busy, False if not.
        """
        if connection_index is None and session_index is None:
            return self.active_session.Busy
        if connection_index is not None and session_index is None:
            return self.connections[connection_index].Sessions[self.connections[connection_index].Sessions.Count - 1].Busy
        if connection_index is None and session_index is not None:
            return self.active_connection.Sessions[session_index].Busy
        return self.connections[connection_index].Sessions[session_index].Busy

    def is_session_active(self, connection_index: int | None = None, session_index: int | None = None) -> bool:
        """
        Checks if the SAP session is active.

        Args:
            connection_index (int | None, optional): Connection index. Defaults to None.
            session_index (int | None, optional): Session index. Defaults to None.

        Returns:
            bool: True if the session is active, False if not.
        """
        if connection_index is None and session_index is None:
            return self.active_session.IsActive
        if connection_index is not None and session_index is None:
            return self.connections[connection_index].Sessions[self.connections[connection_index].Sessions.Count - 1].IsActive
        if connection_index is None and session_index is not None:
            return self.active_connection.Sessions[session_index].IsActive
        return self.connections[connection_index].Sessions[session_index].IsActive

    def get_connection(self, connection_index: int) -> win32com.client.CDispatch:
        """
        Returns the connection object

        Args:
            connection_index (int): Connection Index

        Returns:
            win32com.client.CDispatch: GuiConnection
        """
        return self.connections[connection_index]

    def get_session(self, connection_index: int = None, session_index: int = None) -> win32com.client.CDispatch:
        """
        Returns the SAP session object.

        Args:
            connection_index (int, optional): Connection index. Defaults to None.
            session_index (int, optional): Session index. Defaults to None.

        Returns:
            win32com.client.CDispatch: GuiSession
        """
        connection = self.active_connection if connection_index is None else self.connections.Connections[connection_index]
        return self.active_session if connection_index is None and session_index is None else connection.Sessions[session_index]

    def set_active_window(self, index: int):
        """
        Sets active window for active Sap session

        Args:
            index (int): windows index
        """
        self.active_window = self.active_session.findById(f'wnd[{index}]')

    # Sap Logon

    def close_sap_logon(self):
        """
        Closes Sap Logon application.
        """
        try:
            subprocess.check_call(f'taskkill /F /IM saplogon.exe /T /FI "USERNAME eq {getlogin()}"', stdout=subprocess.PIPE, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
        except:
            pass

    # Objects

    def get_object(self, field_id: str):
        """
        Returns the SAp object by field id.

        Args:
            field_id (str): Field Id

        Returns:
            win32com.client.CDispatch:  Any 
        """
        return self.__get_object(field_id)

    def get_object_type(self, field_id: str) -> str:
        """
        Returns the Sap object's type.

        Args:
            field_id (str): Field id

        Returns:
            str: type name
        """
        return self.__get_object(field_id).Type

    def check_if_object_exists(self, field_id: str) -> bool:
        """
        Checks if the Sap object exists.

        Args:
            field_id (str): Field Id

        Raises:
            Exception: COM object exceptions

        Returns:
            bool: True is object exists, False if not
        """
        try:
            self.active_session.findById(field_id)
            return True
        except Exception as ex:
            if "The control could not be found by id" in str(ex):
                return False
            else:
                raise Exception from ex

    # Common actions

    def send_v_key(self, key: int, window_index: int = 0):
        """
        Sends the SAP virtual key to the active window.

        Args:
            key (int): Sap Virtual Key Value (without "V"; just a number)
            window_index (int): The index of Sap Window; defaults to 0

        Full list of vitual keys: https://experience.sap.com/files/guidelines/References/nv_fkeys_ref2_e.htm
        """
        window = self.__get_object(f'wnd[{window_index}]')
        window.SendVKey(key)

    def set_focus(self, field_id: str):
        """
        Sets focus on the Sap object.

        Args:
            field_id (str): Field Id
        """
        self.__get_object(field_id).SetFocus()

    def run_transaction(self, transaction_code: str):
        """
        Runs Sap transaction.
        There is no need to add "/n" or go back to the start screen.

        Args:
            transaction_code (str): transaction code
        """
        self.active_session.StartTransaction(transaction_code)
        status = self.get_status_bar_message()
        if status.type == 'E':
            raise Exception(f'{status.type} : {status.text}')

    def stop_transaction(self):
        """
        Stops Sap transaction.
        """
        self.active_session.EndTransaction()

    def get_status_bar_message(self, window_index: int = 0) -> StatusBar:
        """
        Returns the type and the text of Sap statusbar message.

        Args:
            window_index (int, optional): Windows index. Defaults to 0.

        Returns:
            StatusBar: StatusBar(text, type)
        """
        status_bar = self.__get_object(f'wnd[{window_index}]/sbar')
        return StatusBar(status_bar.Text, status_bar.MessageType)

    def get_text(self, field_id: str) -> str:
        """
        Returns value of Text property of Sap object.

        Args:
            field_id (str): Field Id

        Returns:
            str: text value
        """
        return self.__get_object(field_id).Text

    def set_text(self, field_id: str, text: str):
        """
        Set value of Text property of Sap object.

        Args:
            field_id (str): Field id
            text (str): text value
        """
        self.__get_object(field_id).Text = text

    def select(self, field_id: str):
        """
        Select Sap object

        Args:
            field_id (str): Field id
        """
        self.__get_object(field_id).Select()

    def select_combobox_item(self, field_id: str, key_id: str):
        """
        Select ComboBox item.

        Args:
            field_id (str): ComboBox object field id
            key_id (str): key id of the item
        """
        self.__get_object(field_id).Key = key_id

    def check_checkbox(self, field_id: str):
        """
        Mark checkbox field as checked.

        Args:
            field_id (str): Field id
        """
        self.__get_object(field_id).Selected = True

    def uncheck_checkbox(self, field_id: str):
        """
        Mark checkbox field as unchecked.

        Args:
            field_id (str): Field id
        """
        self.__get_object(field_id).Selected = False

    def select_context_menu_item(self, field_id: str, item_id: str):
        """
        Select context menu item.

        Args:
            field_id (str): Context menu field id.
            item_id (str): Menu item id.
        """
        self.__get_object(field_id).SelectContextMenuItem(item_id)

    def press_context_menu_item(self, field_id: str, item_id: str):
        """
        Press context menu item.

        Args:
            field_id (str): Context menu field id
            item_id (str): Menu item id
        """
        self.__get_object(field_id).PressContextButton(item_id)

    def press_button(self, field_id: str):
        """
        Press button

        Args:
            field_id (str): Field id
        """
        self.__get_object(field_id).press()

    # Custom properties and methods

    def set_property(self, field_id: str, property_name: str, property_value):
        """
        Set value of custom property.

        Args:
            field_id (str): Field id.
            property_name (str): Property name
            property_value (_type_): Value to be set
        """
        setattr(self.__get_object(field_id), property_name, property_value)

    def get_property(self, field_id: str, property_name: str) -> object:
        """
        Get value of custom property.

        Args:
            field_id (str): Field id.
            property_name (str): Property name.

        Returns:
            object: Value of the property.
        """
        return getattr(self.__get_object(field_id), property_name)

    def invoke_method(self, field_id: str, method_name: str, *args) -> object:
        """
        Execute custom method.

        Args:
            field_id (str): Field id.
            method_name (str): Method name.
            *args (Any, optional): comma separated values of arguments passed to the method.

        Returns:
            object: Value returned by the method.
        """
        return getattr(self.__get_object(field_id), method_name)(*args)

    # GridView

    def count_gridview_rows(self, grid_view_id: str) -> int:
        """
        Count row of GridView object.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number of rows
        """
        grid_view = self.__get_object(grid_view_id)
        return grid_view.RowCount

    def count_gridview_columns(self, grid_view_id: str) -> int:
        """
        Count columns of GridView object.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number of columns
        """
        grid_view = self.__get_object(grid_view_id)
        return grid_view.ColumnCount

    def get_current_gridview_cell_value(self, grid_view_id: str) -> object:
        """
        Return the value of current GridView cell.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            object: Value of GridView cell
        """
        grid_view = self.__get_object(grid_view_id)
        return grid_view.GetCellValue(grid_view.CurrentCellRow, grid_view.CurrentCellColumn)

    def get_current_gridview_cell(self, grid_view_id: str):
        """
        Return row index and column index of current GridView cell.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            GridViewCell['row', 'column']: object with row and column attributes.
        """
        grid_view = self.__get_object(grid_view_id)
        GridViewCell = namedtuple('GridViewCell', ['row', 'column'])
        return GridViewCell(grid_view.CurrentCellRow, self.__get_gridview_column_index__(grid_view, grid_view.CurrentCellColumn))

    def set_current_gridview_cell(self, grid_view_id: str, row_index: int, column_index: int):
        """
        Set current cell of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            row_index (int): Row index.
            column_index (int): Column index.
        """
        grid_view = self.__get_object(grid_view_id)
        grid_view.SetCurrentCell(
            row_index, self.__get_gridview_column_name__(grid_view, column_index))

    def get_current_gridview_column_name(self, grid_view_id: str) -> str:
        """
        Return the name of current column of the GridView object.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            str: column name
        """
        grid_view = self.__get_object(grid_view_id)
        return grid_view.CurrentCellColumn

    def set_current_gridview_column_name(self, grid_view_id: str, column_name: str):
        """
        Set current column of the GridView by column name

        Args:
            grid_view_id (str): GridView field id.
            column_name (str): Column name.
        """
        grid_view = self.__get_object(grid_view_id)
        grid_view.CurrentCellColumn = column_name

    def get_current_gridview_column_index(self, grid_view_id: str) -> int:
        """
        Return index of current GridView column.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number value.
        """
        grid_view = self.__get_object(grid_view_id)
        for column_index in range(0, grid_view.ColumnOrder.Count):
            if grid_view.ColumnOrder[column_index] == grid_view.CurrentCellColumn:
                return column_index

    def set_current_gridview_column_index(self, grid_view_id: str, column_index: int):
        """
        Set the index of current column of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            column_index (int): Column Index
        """
        grid_view = self.__get_object(grid_view_id)
        grid_view.CurrentCellColumn = self.__get_gridview_column_name__(
            grid_view, column_index)

    def get_current_gridview_row_index(self, grid_view_id: str) -> int:
        """
        Return the index of current GridView row.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number value
        """
        grid_view = self.__get_object(grid_view_id)
        return grid_view.CurrentCellRow

    def set_current_gridview_row_index(self, grid_view_id: str, row_index: int):
        """
        Set the index of current row of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            row_index (int): Row index.
        """
        grid_view = self.__get_object(grid_view_id)
        grid_view.CurrentCellRow = row_index

    def get_selected_gridview_rows(self, grid_view_id: str) -> list:
        """
        Return indexes of selected GridView rows.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            list: list of selected row indexes
        """
        grid_view = self.__get_object(grid_view_id)
        selected_rows: str = str(grid_view.SelectedRows)
        if selected_rows == "":
            return None
        rows_list: list = []
        for row in selected_rows.split(','):
            if '-' in row:
                index_range: list[str] = row.split('-')
                for index in range(index_range[0], index_range[1]):
                    rows_list.append(index)
            rows_list.append(int(row))
        return rows_list

    def set_selected_gridview_rows(self, grid_view_id: str, row_indexes: list[int] | str):
        """
        Set selected rows of GridView object.

        Args:
            grid_view_id (str): GridView field id
            row_indexes (list[int] | str): can be a str, ex. "1", or "1,2" or "1-3" if you want to select a range, or the list of int ex. [1,2,3]
        """
        selected_rows: str
        if isinstance(row_indexes, str):
            selected_rows = row_indexes
        if isinstance(row_indexes, list[int]):
            selected_rows = ','.join([str(item) for item in row_indexes])

        grid_view = self.__get_object(grid_view_id)
        grid_view.SelectedRows(selected_rows)

    # GridView methods
    def clear_gridview_selection(self, grid_view_id: str):
        """
        Clear row selection of the GridView object.

        Args:
            grid_view_id (str): GridView field id.
        """
        grid_view = self.__get_object(grid_view_id)
        grid_view.ClearSelection()

    def double_click_gridview_cell(self, grid_view_id: str, row_index: int = None, column_index: int = None):
        """
        Double click the cell of GridView object.

        Args:
            grid_view_id (str): GridView field id
            row_index (int, optional): _description_. Defaults to None.
            column_index (int, optional): _description_. Defaults to None.
        """
        grid_view = self.__get_object(grid_view_id)
        if row_index is not None or column_index is not None:
            column_name: str = self.__get_gridview_column_name__(grid_view, self.get_current_gridview_column_index(grid_view_id)) if column_index is None else self.__get_gridview_column_name__(grid_view, column_index)
            row_index = self.get_current_gridview_row_index(grid_view_id) if row_index is None else row_index
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index
            grid_view.selectedRows = row_index
        grid_view.DoubleClickCurrentCell()

    def click_gridview_cell(self, grid_view_id: str, row_index: int = None, column_index: int = None):
        grid_view = self.__get_object(grid_view_id)
        if row_index is None and column_index is None:
            grid_view.ClickCurrentCell()
        else:
            column_name = self.__get_gridview_column_name__(
                grid_view, column_index)
            grid_view.currentCellRow = row_index
            grid_view.selectedRows = row_index
            grid_view.Click(row_index, column_name)

    def convert_gridview_column_index_to_name(self, grid_view_id: str, column_name: str) -> int:
        grid_view = self.__get_object(grid_view_id)
        column_index: int
        for column_index in range(0, grid_view.ColumnCount):
            if column_name == grid_view.ColumnOrder[column_index]:
                return column_index

    def get_gridview_cell_address_by_cell_value(self, grid_view_id: str, cell_value: str) -> list:
        grid_view = self.__get_object(grid_view_id)
        indexes = self.__get_gridview_cell_address_by_value__(
            grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(
                f'The GridView row not found for the value: {cell_value}')
        return indexes

    def get_gridview_cell_state(self, grid_view_id: str, row_index: int = None, column_index: int = None) -> str:
        grid_view = self.__get_object(grid_view_id)
        r_index = row_index if row_index is not None else self.get_current_gridview_row_index
        c_index = column_index if column_index is not None else self.get_current_gridview_column_index
        return grid_view.GetCellState(r_index, self.__get_gridview_column_name__(grid_view, c_index))

    def get_gridview_cell_value(self, grid_view_id: str, row_index: int = None, column_index: int = None) -> object:
        grid_view = self.__get_object(grid_view_id)
        r_index = row_index if row_index is not None else self.get_current_gridview_row_index
        c_index = column_index if column_index is not None else self.get_current_gridview_column_index
        return grid_view.GetCellValue(r_index, self.__get_gridview_column_name__(grid_view, c_index))

    def press_gridview_toolbar_button(self, grid_view_id: str, button_id: str):
        grid_view = self.__get_object(grid_view_id)
        grid_view.pressToolbarButton(button_id)

    def press_gridview_toolbar_context_button(self, grid_view_id: str, button_id: str):
        grid_view = self.__get_object(grid_view_id)
        grid_view.pressToolbarContextButton(button_id)

    def press_gridview_toolbar_context_button_and_select_context_menu_item(self, grid_view_id: str, button_id: str, function_code: str):
        grid_view = self.__get_object(grid_view_id)
        grid_view.pressToolbarContextButton(button_id)
        sleep(1)
        grid_view.selectContextMenuItem(function_code)
        grid_view.ActiveWindow.setFocus()

    def select_gridview_all_cells(self, grid_view_id: str):
        grid_view = self.__get_object(grid_view_id)
        grid_view.SelectAll()

    def select_gridview_column(self, grid_view_id: str, column_index: int):
        grid_view = self.__get_object(grid_view_id)
        grid_view.SelectColumn(
            self.__get_gridview_column_name__(grid_view, column_index))

    def select_gridview_context_menu_item(self, grid_view_id: str, function_code: str):
        grid_view = self.__get_object(grid_view_id)
        grid_view.selectContextMenuItem(function_code)

    def select_gridview_rows_by_cell_value(self, grid_view_id: str, cell_value: object):
        grid_view = self.__get_object(grid_view_id)
        indexes = self.__get_gridview_cell_address_by_value__(
            grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception('The GridView row not found for the value: %s' % cell_value)

        for row_index, column_index in indexes:
            column_name = self.__get_gridview_column_name__(grid_view, column_index)
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index

        grid_view.selectedRows = ','.join([str(r) for r, c in indexes])

    def set_gridview_current_cell_by_cell_value(self, grid_view_id: str, cell_value: object):
        grid_view = self.__get_object(grid_view_id)
        indexes = self.__get_gridview_cell_address_by_value__(
            grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(
                f'The GridView row not found for the value: {cell_value}')

        for row_index, column_index in indexes:
            column_name = self.__get_gridview_column_name__(
                grid_view, column_index)
            grid_view.SetCurrentCell(row_index, column_name)

    def gridview_to_array(self, grid_view_id: str) -> list:
        grid_view = self.__get_object(grid_view_id)
        return [self.__get_gridview_headers__(grid_view), *self.__get_gridview_body__(grid_view)]

    def gridview_to_dict(self, grid_view_id: str) -> dict:
        grid_view = self.__get_object(grid_view_id)
        return {'columns': self.__get_gridview_headers__(grid_view), 'data': self.__get_gridview_body__(grid_view)}

    def gridview_to_dataframe(self, grid_view_id: str) -> DataFrame:
        grid_view = self.__get_object(grid_view_id)
        return DataFrame(data=self.__get_gridview_body__(grid_view), columns=self.__get_gridview_headers__(grid_view))

    def gridview_to_csv(self, grid_view_id: str, path_or_buf: str):
        grid_view = self.__get_object(grid_view_id)
        self.gridview_to_dataframe(grid_view).to_csv(
            path_or_buf=path_or_buf, index=False)

    def gridview_to_xlsx(self, grid_view_id: str, file_path: str):
        grid_view = self.__get_object(grid_view_id)
        self.gridview_to_dataframe(grid_view).to_excel(file_path, index=False)

    # Magic methods
    def __get_object(self, field_id: str):
        if not self.__is_object__(field_id):
            raise Exception(f'Cannot find the field: {field_id}.')
        return self.active_session.findById(field_id)

    def __is_object__(self, field_id: str):
        try:
            self.active_session.findById(field_id)
            return True
        except:
            return False

    def __get_session_info__(self, session: win32com.client.CDispatch = None) -> dict:
        ses = self.active_session if session is None else session
        return {
            'user': ses.Info.User.upper(),
            'sid': ses.Info.SystemName.upper(),
            'application_server': ses.Info.ApplicationServer.upper(),
            'client': ses.Info.Client.upper()
        }

    # Magic methods - Grid View
    def __get_gridview_column_index__(self, grid_view: win32com.client.dynamic.CDispatch, column_name: str):
        for column_index in range(0, grid_view.ColumnOrder.Count):
            return column_index if column_name == grid_view.ColumnOrder[column_index] else None

    def __get_gridview_column_name__(self, grid_view: win32com.client.dynamic.CDispatch, column_index: int) -> str:
        return grid_view.ColumnOrder[column_index]

    def __get_gridview_cell_address_by_value__(self, grid_view: win32com.client.dynamic.CDispatch, cell_value: object) -> list:
        Cell_Address = namedtuple('Cell_Address', 'Row_Index Column_Index')
        results = []
        for row_index in range(0, grid_view.RowCount):
            for column_index in range(0, grid_view.ColumnCount):
                if cell_value == grid_view.GetCellValue(row_index, grid_view.ColumnOrder[column_index]):
                    results.append(Cell_Address(row_index, column_index))
        return results

    def __get_gridview_headers__(self, grid_view: win32com.client.dynamic.CDispatch) -> list:
        return [grid_view.GetColumnTitles(column_name)[0] for column_name in grid_view.ColumnOrder]

    def __get_gridview_body__(self, grid_view: win32com.client.dynamic.CDispatch) -> list:
        body = []
        for row_index in range(0, grid_view.RowCount):
            row = []
            for column_index in range(0, grid_view.ColumnCount):
                row.append(grid_view.GetCellValue(
                    row_index, self.__get_gridview_column_name__(grid_view, column_index)))
            body.append(row)
        return body
