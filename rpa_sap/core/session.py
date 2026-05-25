import win32com.client
from ..exceptions import SapSessionError, SapTransactionError
from .ui_automation import ElementInteractor


class SapSession:
    """
    Represents an active SAP GUI session and acts as the main context
    for UI automation and transaction execution.
    """

    def __init__(
        self,
        com_session: win32com.client.CDispatch,
        com_connection: win32com.client.CDispatch = None,
    ):
        self.com_session = com_session
        self.com_connection = com_connection
        self.interactor = ElementInteractor(self)

    @property
    def info(self) -> dict:
        """Returns basic information about the session."""
        return {
            "user": self.com_session.Info.User.upper(),
            "sid": self.com_session.Info.SystemName.upper(),
            "application_server": self.com_session.Info.ApplicationServer.upper(),
            "client": self.com_session.Info.Client.upper(),
            # "is_active": self.com_session.IsActive,
            "is_active": self.com_session.ActiveWindow.Text != "",
            "is_busy": self.com_session.Busy,
            "session_number": self.com_session.Info.SessionNumber,
            "transaction": self.com_session.Info.Transaction,
        }

    @property
    def id(self) -> str:
        return self.com_session.Id

    def findById(self, field_id: str):
        """Find an element by its ID within the active session."""
        return self.com_session.findById(field_id)

    def run_transaction(self, transaction_code: str):
        """
        Runs SAP transaction.
        There is no need to add "/n" or go back to the start screen.
        """
        self.com_session.StartTransaction(transaction_code)
        status = self.interactor.get_status_bar_message()
        if status.type == "E":
            raise SapTransactionError(f"{status.type} : {status.text}")

    def stop_transaction(self):
        """Stops SAP transaction."""
        self.com_session.EndTransaction()

    def set_active_window(self, index: int):
        """Sets active window for active SAP session."""
        return self.com_session.findById(f"wnd[{index}]")
