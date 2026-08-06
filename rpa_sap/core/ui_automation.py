import datetime
import time
from typing import Any, Union
from ..exceptions import SapElementNotFoundError


class BaseMixin:
    session: Any

    def _get_object(self, field_id: str) -> Any:
        raise NotImplementedError


class CoreResolutionMixin(BaseMixin):
    """Provides core object resolution and existence checking."""

    def _get_object(self, field_id: str) -> Any:
        """Resolves a field ID to its underlying COM object."""
        try:
            return self.session.findById(field_id)
        except Exception as ex:
            raise SapElementNotFoundError(f"Cannot find the field: {field_id}.") from ex

    def _is_object(self, field_id: str) -> bool:
        """Checks if a COM object can be resolved by its field ID."""
        try:
            self.session.findById(field_id)
            return True
        except Exception:
            return False

    def get_object_type(self, field_id: str) -> str:
        """Returns the SAP GUI type of the specified object."""
        return self._get_object(field_id).Type

    def check_if_object_exists(self, field_id: str) -> bool:
        """Returns True if the object exists in the current session."""
        return self._is_object(field_id)

    def wait_until_object_exists(
        self,
        field_id: str,
        timeout: Union[int, datetime.timedelta] = 30,
        ignore_timeout: bool = True,
    ) -> bool:
        """
        Polls for an object until it exists or the timeout is reached.

        Args:
            field_id (str): The ID of the element to wait for.
            timeout (int | datetime.timedelta): Max time to wait in seconds or timedelta. Defaults to 30.
            ignore_timeout (bool): If True, returns False on timeout. If False, raises SapElementNotFoundError.

        Returns:
            bool: True if object was found, False otherwise (if ignore_timeout is True).
        """
        _time = datetime.datetime.now()
        _time += (
            datetime.timedelta(seconds=timeout) if isinstance(timeout, int) else timeout
        )
        while (
            datetime.datetime.now() < _time
            and self.check_if_object_exists(field_id) is False
        ):
            time.sleep(1)

        if not ignore_timeout and not self.check_if_object_exists(field_id):
            raise SapElementNotFoundError(f"Sap object {field_id} couldn't be found.")

        return self.check_if_object_exists(field_id)


class KeyboardWindowMixin(BaseMixin):
    """Handles keyboard events and window-level interactions."""

    def send_v_key(self, key: int, window_index: int = 0) -> None:
        """Sends a virtual key to the specified window."""
        window = self._get_object(f"wnd[{window_index}]")
        window.SendVKey(key)

    def press_enter(self, window_index: int = 0) -> None:
        """Presses the Enter key (VKey 0)."""
        self.send_v_key(0, window_index)

    def press_F2(self, window_index: int = 0) -> None:
        """Presses the F2 key."""
        self.send_v_key(2, window_index)

    def press_F3(self, window_index: int = 0) -> None:
        """Presses the F3 key (Back/Cancel)."""
        self.send_v_key(3, window_index)

    def press_F8(self, window_index: int = 0) -> None:
        """Presses the F8 key (Execute)."""
        self.send_v_key(8, window_index)

    def get_status_bar_message(self, window_index: int = 0):
        """Retrieves the current message and type from the status bar."""
        from ..lib.common import StatusBar

        status_bar = self._get_object(f"wnd[{window_index}]/sbar")
        return StatusBar(status_bar.Text, status_bar.MessageType)


class GeneralElementMixin(BaseMixin):
    """Handles general element interactions like clicking and text entry."""

    def set_focus(self, field_id: str) -> None:
        """Sets UI focus to the specified element."""
        self._get_object(field_id).SetFocus()

    def get_text(self, field_id: str) -> str:
        """Gets the Text property of the specified element."""
        return self._get_object(field_id).Text

    def set_text(self, field_id: str, text: str) -> None:
        """Sets the Text property of the specified element."""
        self._get_object(field_id).Text = text

    def select(self, field_id: str) -> None:
        """Invokes the Select action on the specified element."""
        self._get_object(field_id).Select()

    def press_button(self, field_id: str) -> None:
        """Presses a standard GUI button."""
        self._get_object(field_id).press()

    def double_click(self, field_id: str) -> None:
        """Double clicks the specified element."""
        self._get_object(field_id).doubleClick()


class CheckboxComboboxMixin(BaseMixin):
    """Handles specific interactions with checkboxes and comboboxes."""

    def select_combobox_item(self, field_id: str, key_id: str) -> None:
        """Selects an item in a combobox by its Key."""
        self._get_object(field_id).Key = key_id

    def check_checkbox(self, field_id: str) -> None:
        """Checks a checkbox element."""
        self._get_object(field_id).Selected = True

    def uncheck_checkbox(self, field_id: str) -> None:
        """Unchecks a checkbox element."""
        self._get_object(field_id).Selected = False

    def set_checkbox_state(self, field_id: str, state: bool) -> None:
        """Sets a checkbox to the specified boolean state."""
        self._get_object(field_id).Selected = state

    def get_checkbox_state(self, field_id: str) -> bool:
        """Returns True if the checkbox is checked, False otherwise."""
        return self._get_object(field_id).Selected


class ContextMenuMixin(BaseMixin):
    """Handles context menu interactions."""

    def select_context_menu_item(self, field_id: str, item_id: str) -> None:
        """Selects an item from an open context menu."""
        self._get_object(field_id).SelectContextMenuItem(item_id)

    def press_context_menu_item(self, field_id: str, item_id: str) -> None:
        """Presses a button to open a context menu for a specific item."""
        self._get_object(field_id).PressContextButton(item_id)


class AdvancedPropertyMixin(BaseMixin):
    """Handles direct reading/writing of COM properties and methods."""

    def set_property(
        self, field_id: str, property_name: str, property_value: Any
    ) -> None:
        """Sets an arbitrary COM property on the element."""
        setattr(self._get_object(field_id), property_name, property_value)

    def get_property(self, field_id: str, property_name: str) -> Any:
        """Gets an arbitrary COM property from the element."""
        return getattr(self._get_object(field_id), property_name)

    def invoke_method(self, field_id: str, method_name: str, *args) -> Any:
        """Invokes an arbitrary COM method on the element."""
        return getattr(self._get_object(field_id), method_name)(*args)


class ElementInteractor(
    CoreResolutionMixin,
    KeyboardWindowMixin,
    GeneralElementMixin,
    CheckboxComboboxMixin,
    ContextMenuMixin,
    AdvancedPropertyMixin,
):
    """
    Handles direct interaction with SAP GUI elements.
    Acts as a Facade bringing together all specific interactors.
    """

    def __init__(self, session):
        """
        Initializes the ElementInteractor with a SapSession.

        Args:
            session: The active SapSession instance.
        """
        self.session = session
