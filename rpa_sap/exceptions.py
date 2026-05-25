class SapRpaError(Exception):
    """Base exception for rpa-sap."""
    pass

class SapConnectionError(SapRpaError):
    """Raised when there is an error connecting to the SAP GUI."""
    pass

class SapSessionError(SapRpaError):
    """Raised when there is an error with the SAP session."""
    pass

class SapElementNotFoundError(SapRpaError):
    """Raised when a specific SAP element cannot be found."""
    pass

class SapTransactionError(SapRpaError):
    """Raised when an error occurs during SAP transaction execution."""
    pass

class SapProcessError(SapRpaError):
    """Raised when there is an error managing the SAP process (e.g. saplogon.exe)."""
    pass
