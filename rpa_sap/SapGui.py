""" RPA toolchain to automate SapGui (Sap Scripting) """

from os import getlogin
import subprocess
from time import sleep
import datetime
import re
import win32com.client
from .lib.GridView import GridView
from .lib.common import GuiObject, StatusBar

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
    def __init__(self):
        self.__sap_gui: win32com.client.CDispatch = None
        self.__application: win32com.client.CDispatch = None
        self.active_connection: win32com.client.CDispatch = None
        self.active_session: win32com.client.CDispatch = None
        self.active_window: win32com.client.CDispatch = None
        self.active_objects: list[GuiObject] = []
        self.grid_view = GridView(self)

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
        except Exception:
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
        Returns the Sap object by field id.

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

    def wait_until_object_exists(self, field_id: str, timeout: int | datetime.timedelta = 30, ignore_timeout: bool = True) -> bool:
        """ Wait until the object (by field id) exists for given time (timeout).

        Args:
            field_id (str): Field Id
            timeout (int | datetime.timedelta, optional): timeout in seconds or timedelta object. Defaults to 30.
            ignore_timeout (bool, optional): ignore if object does not exists and not throw the error. Defaults to True.

        Returns:
            bool: True if object appears, False if not or if the timeout has been reached.
        """
        _time = datetime.datetime.now()
        _time += datetime.timedelta(seconds=timeout) if isinstance(timeout, int) else timeout
        while datetime.datetime.now() < _time and self.check_if_object_exists(field_id) is False:
            sleep(1)
        if ignore_timeout is False:
            raise Exception(f'Sap object {field_id} couldn\'t be found.')

        return self.check_if_object_exists(field_id) is True

    # Common actions

    def send_v_key(self, key: int, window_index: int = 0):
        """
        Sends the SAP virtual key to the window.

        Args:
            key (int): Sap Virtual Key Value (without "V"; just a number)
            window_index (int): The index of Sap Window; defaults to 0

        Full list of vitual keys: https://experience.sap.com/files/guidelines/References/nv_fkeys_ref2_e.htm
        """
        window = self.__get_object(f'wnd[{window_index}]')
        window.SendVKey(key)

    def press_enter(self, window_index: int = 0):
        """
        Sends the SAP virtual key (ENTER) to the window.

        Args:
            window_index (int, optional): The index of Sap Window; Defaults to 0.
        """
        window = self.__get_object(f'wnd[{window_index}]')
        window.SendVKey(0)

    def press_F2(self, window_index: int = 0):
        """ Press F2 button

        Args:
            window_index (int, optional): window id. Defaults to 0.
        """
        window = self.__get_object(f'wnd[{window_index}]')
        window.SendVKey(2)

    def press_F3(self, window_index: int = 0):
        """ Press F3 button

        Args:
            window_index (int, optional): window id. Defaults to 0.
        """
        window = self.__get_object(f'wnd[{window_index}]')
        window.SendVKey(3)

    def press_F8(self, window_index: int = 0):
        """ Press F8 key

        Args:
            window_index (int, optional): window id. Defaults to 0.
        """
        window = self.__get_object(f'wnd[{window_index}]')
        window.SendVKey(8)

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

    def set_checkbox_state(self, field_id: str, state: bool):
        """
        Mark checkbox field as checked or unchecked based on state value

        Args:
            field_id (str): Field id
            state (bool): True (checked) or False (unchecked)
        """
        self.__get_object(field_id).Selected = state

    def get_checkbox_state(self, field_id: str) -> bool:
        """
        Returns a state of the checkbox field.

        Args:
            field_id (str): Field Id

        Returns:
            bool: True if checkbox is checked or False if not checked
        """
        return self.__get_object(field_id).Selected

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

    def double_click(self, field_id: str):
        """
        Double click field

        Args:
            field_id (str): Field id.
        """
        self.__get_object(field_id).doubleClick()


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

    def get_property(self, field_id: str, property_name: str):
        """
        Get value of custom property.

        Args:
            field_id (str): Field id.
            property_name (str): Property name.

        Returns:
            object: Value of the property.
        """
        return getattr(self.__get_object(field_id), property_name)

    def invoke_method(self, field_id: str, method_name: str, *args):
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
