from os import getlogin
import subprocess  # nosec B404
import time
import win32com.client
import warnings
import psutil
from ..exceptions import SapConnectionError, SapProcessError
from .session import SapSession


class _ProcessManager:
    """Helper class to manage Windows OS processes using psutil."""

    @staticmethod
    def is_sap_running(username: str = None) -> bool:
        username = (username or getlogin()).upper()
        try:
            for proc in psutil.process_iter(["name", "username"]):
                if not proc.is_running() or proc.status() in (
                    psutil.STATUS_ZOMBIE,
                    psutil.STATUS_DEAD,
                ):
                    continue
                name = proc.info.get("name")
                if name and name.lower() in ("saplogon.exe", "sapgui.exe"):
                    proc_user = proc.info.get("username")
                    if proc_user and username in proc_user.upper():
                        return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
        return False

    @staticmethod
    def close_process(process_name: str, username: str = None):
        username = (username or getlogin()).upper()
        try:
            for proc in psutil.process_iter(["name", "username"]):
                name = proc.info.get("name")
                if name and name.lower() == process_name.lower():
                    proc_user = proc.info.get("username")
                    if proc_user and username in proc_user.upper():
                        proc.kill()
                        try:
                            proc.wait(timeout=3)
                        except psutil.TimeoutExpired:
                            pass
        except Exception as e:
            warnings.warn(
                f"Process {process_name} not found or could not be closed. {e}",
                UserWarning,
            )


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
            raise SapConnectionError(
                "Cannot connect to SAPGUI session. SAPGUI seems to be not opened."
            ) from ex

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
        password: str | None = None,
        client: str = "900",
        language: str = "EN",
        timeout: int = 15,
        step: int = 1,
    ) -> SapSession:
        """
        Opens and logs in to a new SAP session. Support for both SSO (password=None) and standard login.
        """
        sap_running = _ProcessManager.is_sap_running()
        sapgui_path = "C:/Program Files (x86)/SAP/FrontEnd/SAPgui/SAPgui.exe"

        if not sap_running:
            try:
                subprocess.Popen([sapgui_path, connection_string])
            except Exception as ex:
                self.close_sap_logon()
                raise SapProcessError(f"Failed to start {sapgui_path}") from ex
            time.sleep(4)

        end_time = time.time() + timeout
        connection = None
        session = None

        while time.time() < end_time:
            try:
                self._initialize_engine()
                if self.application:
                    if not sap_running:
                        # Find the connection that SAPgui.exe automatically opened
                        if self.connections.Count > 0:
                            connection = self.connections[self.connections.Count - 1]
                            if connection.Children.Count > 0:
                                session = connection.Children(0)
                                break
                    else:
                        # Actively open a new connection on the running engine
                        connection = self.application.OpenConnectionByConnectionString(
                            connection_string, True
                        )
                        session = connection.Children(0)
                        break
            except Exception:
                # If we were expecting a running SAP but it's stale/crashing:
                if sap_running:
                    self.close_sap_logon()
                    time.sleep(2)
                    sap_running = False
                    try:
                        subprocess.Popen([sapgui_path, connection_string])
                    except Exception as ex:
                        raise SapProcessError(f"Failed to start {sapgui_path}") from ex
                    time.sleep(4)

            time.sleep(step)
        else:
            raise SapConnectionError("Timeout while waiting for SAP session to open.")

        sap_session = SapSession(session, connection)

        # NOTE: Keeping password in memory can be a security issue, but kept for compatibility.
        sap_session._rfc_credentials = {
            "connection_string": connection_string,
            "user_id": user_id,
            "password": password,
            "client": client,
        }

        self._perform_login(sap_session, user_id, password, client, language)
        return sap_session

    def _perform_login(
        self,
        sap_session: SapSession,
        user_id: str,
        password: str | None,
        client: str,
        language: str,
    ):
        """Automates the SAP GUI UI login process."""
        session = sap_session.com_session

        try:
            active_window = session.findById("wnd[0]")
            active_window.maximize()
        except Exception:
            pass

        # Wait up to 5 seconds for the login screen to render. If it doesn't, assume SSO bypassed it.
        user_field = None
        for _ in range(5):
            try:
                user_field = session.findById("wnd[0]/usr/txtRSYST-BNAME")
                break
            except Exception:
                time.sleep(1)

        if user_field:
            try:
                user_field.Text = user_id
                if password:
                    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
                session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = client
                session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = language
                active_window.SendVKey(0)
            except Exception as e:
                warnings.warn(f"Error during login automation: {e}", UserWarning)

        # Check if "License Information for Multiple Logon" pops up
        if sap_session.interactor.check_if_object_exists("wnd[1]"):
            if sap_session.interactor.check_if_object_exists(
                "wnd[1]/usr/radMULTI_LOGON_OPT2"
            ):
                sap_session.interactor.select("wnd[1]/usr/radMULTI_LOGON_OPT2")
                sap_session.interactor.press_button("wnd[1]/tbar[0]/btn[0]")

        status = sap_session.interactor.get_status_bar_message()
        if status.type == "E":
            raise SapConnectionError(f"{status.type} : {status.text}")

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

                    def _match_session(s):
                        i = s.Info
                        return (
                            i.SystemName == sid.upper()
                            and i.Client == client
                            and i.User == user_id.upper()
                            and i.ApplicationServer.upper()
                            == application_server.upper()
                        )

                    for connection in self.connections:
                        for session in connection.Sessions:
                            if _match_session(session):
                                active_connection = connection
                                active_session = session
                                break
                        if active_session:
                            break
                else:
                    active_connection = self.connections[self.connections.Count - 1]
                    active_session = active_connection.Sessions[
                        active_connection.Sessions.Count - 1
                    ]

            if not active_session:
                raise ValueError("Matching session not found.")

            return SapSession(active_session, active_connection)
        except Exception as ex:
            raise SapConnectionError(
                "Cannot activate session. Please verify provided properties are correct."
            ) from ex

    def check_if_session_exists(
        self, connection_index: int | None = None, session_index: int | None = None
    ) -> bool:
        try:
            self._initialize_engine()
            con_index = (
                connection_index
                if connection_index is not None
                else self.connections.Count - 1
            )
            ses_index = (
                session_index
                if session_index is not None
                else self.connections[con_index].Sessions.Count - 1
            )
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
                    try:
                        connection.CloseSession(session.Id)
                    except Exception as e:
                        warnings.warn(f"Failed to close session: {e}", UserWarning)
        except Exception:
            pass  # nosec B110

    def close_session(self, sap_session: SapSession):
        """Closes a specific SAP session."""
        if sap_session.com_connection:
            try:
                sap_session.com_connection.CloseSession(sap_session.com_session.Id)
            except Exception as e:
                warnings.warn(f"Failed to close session: {e}", UserWarning)
            # Remove COM references to avoid GC issues / RPC crashes
            sap_session.com_session = None
            sap_session.com_connection = None

    def close_sap_logon(self, username: str = None):
        """
        Closes Sap Logon application opened by the specific user.
        """
        # Clear engine COM references before terminating process
        self.application = None
        self.sap_gui = None
        _ProcessManager.close_process("saplogon.exe", username)
        _ProcessManager.close_process("sapgui.exe", username)
        time.sleep(2)  # Give Windows time to clean up ROT entry for dead processes

    def close_process(self, process_name: str, username: str = None):
        """
        Closes Windows process opened by the specific user.
        """
        _ProcessManager.close_process(process_name, username)
