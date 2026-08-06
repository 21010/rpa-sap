import abc
import requests
import pandas as pd
from typing import Any, Dict, List, Optional
from rpa_sap.exceptions import SapODataError


class AuthStrategy(abc.ABC):
    """Abstract base class for OData authentication strategies."""

    @abc.abstractmethod
    def configure_session(self, session: requests.Session) -> None:
        """Configure the requests.Session with authentication details."""
        pass


class BasicAuthStrategy(AuthStrategy):
    """Basic Authentication strategy."""

    def __init__(self, username: str, password: str):
        self.username = username
        self.password = password

    def configure_session(self, session: requests.Session) -> None:
        session.auth = (self.username, self.password)


class OAuth2Strategy(AuthStrategy):
    """OAuth2 Authentication strategy using a Bearer token."""

    def __init__(self, token: str):
        self.token = token

    def configure_session(self, session: requests.Session) -> None:
        session.headers.update({"Authorization": f"Bearer {self.token}"})


class NoAuthStrategy(AuthStrategy):
    """No Authentication strategy (for testing or open endpoints)."""

    def configure_session(self, session: requests.Session) -> None:
        pass


class ODataClient:
    """Client for interacting with SAP OData services."""

    def __init__(self, base_url: str, auth_strategy: AuthStrategy):
        self.base_url = base_url.rstrip("/")
        self.auth_strategy = auth_strategy
        self.session = requests.Session()
        self.auth_strategy.configure_session(self.session)
        self.csrf_token: Optional[str] = None

    def fetch_csrf_token(self) -> str:
        """Fetches and caches the X-CSRF-Token from the SAP Gateway."""
        headers = {"X-CSRF-Token": "Fetch"}
        try:
            response = self.session.get(self.base_url, headers=headers)
            response.raise_for_status()
            token = response.headers.get("x-csrf-token")
            if not token:
                raise SapODataError(
                    "Failed to fetch CSRF token from SAP OData service."
                )
            self.csrf_token = token
            return token
        except requests.RequestException as e:
            raise SapODataError(f"Error fetching CSRF token: {e}")

    def _get_csrf_headers(self) -> Dict[str, str]:
        """Returns headers with the CSRF token. Fetches if not present."""
        if not self.csrf_token:
            self.fetch_csrf_token()
        if self.csrf_token is None:
            raise SapODataError("csrf_token is not available.")
        return {"X-CSRF-Token": self.csrf_token, "Content-Type": "application/json"}

    def get(
        self,
        entity_set: str,
        select: Optional[List[str]] = None,
        filter_query: Optional[str] = None,
        top: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """
        Executes a GET request on an OData entity set.

        Args:
            entity_set (str): The entity set name (e.g., 'UserSet').
            select (list): List of fields to select.
            filter_query (str): The $filter string.
            top (int): The $top value for pagination.

        Returns:
            list: The extracted results list.
        """
        url = f"{self.base_url}/{entity_set}"
        params = {"$format": "json"}
        if select:
            params["$select"] = ",".join(select)
        if filter_query:
            params["$filter"] = filter_query
        if top is not None:
            params["$top"] = str(top)

        try:
            response = self.session.get(url, params=params)
            response.raise_for_status()
            data = response.json()
            return data.get("d", {}).get("results", [])
        except requests.RequestException as e:
            raise SapODataError(
                f"GET request failed for entity set '{entity_set}': {e}"
            )

    def get_dataframe(self, entity_set: str, **kwargs) -> pd.DataFrame:
        """Executes a GET request and returns the results as a Pandas DataFrame."""
        results = self.get(entity_set, **kwargs)
        return pd.DataFrame(results)

    def post(self, entity_set: str, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Executes a POST request to create a new OData entity.
        Automatically handles CSRF token requirements.
        """
        url = f"{self.base_url}/{entity_set}"
        headers = self._get_csrf_headers()
        try:
            response = self.session.post(url, json=payload, headers=headers)
            response.raise_for_status()
            data = response.json()
            return data.get("d", {})
        except requests.RequestException as e:
            raise SapODataError(
                f"POST request failed for entity set '{entity_set}': {e}"
            )
