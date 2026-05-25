from os import getlogin
import subprocess  # nosec B404
import time
import win32com.client
import win32api
import win32con
import warnings
import wmi
from ..exceptions import SapConnectionError, SapProcessError
from .session import SapSession

class ConnectionManager:
    """
    Manages SAP GUI processes, connections, and sessions.
    """
    def __init__(self):
        self.sap_gui = None
        self.application = None

    def _initialize_engine(self):
        try:
            self.sap_gui = win32com.client.GetObject("SAPGUI")
            self.application = self.sap_gui.GetScriptingEngine
        except Exception as ex:
            raise SapConnectionError("Cannot connect to SAPGUI session. SAPGUI seems to be not opened.") from ex

    @property
    def connections(self):
        """Returns: Collection of all SAP connections"""
        if not self.application:
            self._initialize_engine()
        return self.application.Connections

    def open_new_session(
        self,
        connection_string: str,
        user_id: str,
        password: str,
        client: str = "900",
        language: str = "EN",
        timeout: int = 10,
        step: int = 1,
    ) -> SapSession:
        """
        Opens and logs in to a new SAP session.

        Args:
            connection_string (str): Connection string for SAP system
            user_id (str): SAP user id
            password (str): SAP password
            client (str, optional): SAP client. Defaults to '900'.
            language (str, optional): Language. Defaults to "EN".
            timeout (int, optional): Defines how many seconds wait until SAPGUI is opened. Default to 10.
        """
        # Run sapgui.exe with connection string as a parameter
        try:
            subprocess.check_call([  # nosec B603
                "C:/Program Files (x86)/SAP/FrontEnd/SAPgui/SAPgui.exe",
                connection_string,
            ])
            time.sleep(step)
        except (subprocess.CalledProcessError, subprocess.SubprocessError) as ex:
            self.close_sap_logon()
            raise SapProcessError("Failed to start saplogon.exe") from ex

        end_time = time.time() + timeout
        connection = None
        session = None
        while time.time() < end_time:
            try:
                self._initialize_engine()
                connection = self.connections[self.connections.Count - 1]
                session = connection.Sessions[connection.Sessions.Count - 1]
                break
            except Exception:
                time.sleep(step)

        if session is None:
            raise SapConnectionError("Timeout while waiting for SAP session to open.")

        sap_session = SapSession(session, connection)
        sap_session._rfc_credentials = {
            "connection_string": connection_string,
            "user_id": user_id,
            "password": password,
            "client": client
        }
        
        active_window = session.findById("wnd[0]")
        active_window.maximize()
        
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = user_id
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
        session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = client
        session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = language
        active_window.SendVKey(0)
        
        # Check if "License Information for Multiple Logon" pops up
        if sap_session.interactor.check_if_object_exists("wnd[1]"):
            if sap_session.interactor.check_if_object_exists("wnd[1]/usr/radMULTI_LOGON_OPT2"):
                sap_session.interactor.select("wnd[1]/usr/radMULTI_LOGON_OPT2")
                sap_session.interactor.press_button("wnd[1]/tbar[0]/btn[0]")
                
        status = sap_session.interactor.get_status_bar_message()
        if status.type == "E":
            raise SapConnectionError(f"{status.type} : {status.text}")
            
        return sap_session

    def activate_session(
        self,
        connection_index: int | None = None,
        session_index: int | None = None,
        user_id: str | None = None,
        sid: str | None = None,
        application_server: str | None = None,
        client: str | None = None,
    ) -> SapSession:
        """
        Activates an existing SAP session by connection index and session index or connection details.
        """
        self._initialize_engine()
        
        active_connection = None
        active_session = None

        try:
            if connection_index is not None and session_index is not None:
                active_connection = self.connections[connection_index]
                active_session = active_connection.Sessions[session_index]
            elif connection_index is None and session_index is not None:
                active_connection = self.connections[self.connections.Count - 1]
                active_session = active_connection.Sessions[session_index]
            elif connection_index is None and session_index is None:
                if user_id and sid and application_server and client:
                    for connection in self.connections:
                        for session in connection.Sessions:
                            if (
                                session.Info.SystemName == sid.upper()
                                and session.Info.Client == client
                                and session.Info.User == user_id.upper()
                                and session.Info.ApplicationServer.upper() == application_server.upper()
                            ):
                                active_connection = connection
                                active_session = session
                else:
                    active_connection = self.connections[self.connections.Count - 1]
                    active_session = active_connection.Sessions[active_connection.Sessions.Count - 1]

            if not active_session:
                raise ValueError("Matching session not found.")
                
            return SapSession(active_session, active_connection)
        except Exception as ex:
            raise SapConnectionError("Cannot activate session. Please verify provided properties are correct.") from ex

    def check_if_session_exists(
        self, connection_index: int | None = None, session_index: int | None = None
    ) -> bool:
        try:
            self._initialize_engine()
            con_index = connection_index if connection_index is not None else self.connections.Count - 1
            ses_index = session_index if session_index is not None else self.connections[con_index].Sessions.Count - 1
            obj = self.connections[con_index].Sessions[ses_index]
            return obj is not None
        except Exception:
            return False

    def close_all_sessions(self):
        """
        Closes all opened SAP sessions for all opened connections.
        """
        try:
            self._initialize_engine()
            for connection in self.connections:
                for session in connection.Sessions:
                    connection.CloseSession(session.Id)
        except Exception:
            pass  # nosec B110

    def close_session(self, sap_session: SapSession):
        """Closes a specific SAP session."""
        if sap_session.com_connection:
            sap_session.com_connection.CloseSession(sap_session.com_session.Id)

    def close_sap_logon(self, username: str = None):
        """
        Closes Sap Logon application opened by the specific user.
        """
        try:
            self.application = None
            self.sap_gui = None
        except Exception:
            pass  # nosec B110

        self.close_process("saplogon.exe", username)

    def close_process(self, process_name: str, username: str = None):
        """
        Closes Windows process opened by the specific user.
        """
        try:
            c = wmi.WMI()
            username = getlogin() if username is None else username

            for process in c.Win32_Process(name=process_name):
                process_owner = process.GetOwner()
                if process_owner and (
                    username.upper() in str(process_owner).upper()
                ):
                    handle = win32api.OpenProcess(
                        win32con.PROCESS_TERMINATE, 0, process.ProcessId
                    )
                    win32api.TerminateProcess(handle, -1)
                    win32api.CloseHandle(handle)
        except Exception as e:
            warnings.warn(f"Process {process_name} not found or could not be closed. {e}", UserWarning)
