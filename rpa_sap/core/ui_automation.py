import datetime
import time
import win32com.client
from ..exceptions import SapElementNotFoundError
from ..lib.common import StatusBar

class ElementInteractor:
    def __init__(self, session):
        """
        Initializes the ElementInteractor with a SapSession.
        """
        self.session = session

    def _get_object(self, field_id: str):
        try:
            return self.session.findById(field_id)
        except Exception as ex:
            raise SapElementNotFoundError(f"Cannot find the field: {field_id}.") from ex

    def _is_object(self, field_id: str) -> bool:
        try:
            self.session.findById(field_id)
            return True
        except Exception:
            return False

    def get_object_type(self, field_id: str) -> str:
        return self._get_object(field_id).Type

    def check_if_object_exists(self, field_id: str) -> bool:
        return self._is_object(field_id)

    def wait_until_object_exists(
        self,
        field_id: str,
        timeout: int | datetime.timedelta = 30,
        ignore_timeout: bool = True,
    ) -> bool:
        _time = datetime.datetime.now()
        _time += (
            datetime.timedelta(seconds=timeout) if isinstance(timeout, int) else timeout
        )
        while (
            datetime.datetime.now() < _time
            and self.check_if_object_exists(field_id) is False
        ):
            time.sleep(1)
        if ignore_timeout is False:
            raise SapElementNotFoundError(f"Sap object {field_id} couldn't be found.")

        return self.check_if_object_exists(field_id) is True

    def send_v_key(self, key: int, window_index: int = 0):
        window = self._get_object(f"wnd[{window_index}]")
        window.SendVKey(key)

    def press_enter(self, window_index: int = 0):
        self.send_v_key(0, window_index)

    def press_F2(self, window_index: int = 0):
        self.send_v_key(2, window_index)

    def press_F3(self, window_index: int = 0):
        self.send_v_key(3, window_index)

    def press_F8(self, window_index: int = 0):
        self.send_v_key(8, window_index)

    def set_focus(self, field_id: str):
        self._get_object(field_id).SetFocus()

    def get_status_bar_message(self, window_index: int = 0) -> StatusBar:
        status_bar = self._get_object(f"wnd[{window_index}]/sbar")
        return StatusBar(status_bar.Text, status_bar.MessageType)

    def get_text(self, field_id: str) -> str:
        return self._get_object(field_id).Text

    def set_text(self, field_id: str, text: str):
        self._get_object(field_id).Text = text

    def select(self, field_id: str):
        self._get_object(field_id).Select()

    def select_combobox_item(self, field_id: str, key_id: str):
        self._get_object(field_id).Key = key_id

    def check_checkbox(self, field_id: str):
        self._get_object(field_id).Selected = True

    def uncheck_checkbox(self, field_id: str):
        self._get_object(field_id).Selected = False

    def set_checkbox_state(self, field_id: str, state: bool):
        self._get_object(field_id).Selected = state

    def get_checkbox_state(self, field_id: str) -> bool:
        return self._get_object(field_id).Selected

    def select_context_menu_item(self, field_id: str, item_id: str):
        self._get_object(field_id).SelectContextMenuItem(item_id)

    def press_context_menu_item(self, field_id: str, item_id: str):
        self._get_object(field_id).PressContextButton(item_id)

    def press_button(self, field_id: str):
        self._get_object(field_id).press()

    def double_click(self, field_id: str):
        self._get_object(field_id).doubleClick()

    def set_property(self, field_id: str, property_name: str, property_value):
        setattr(self._get_object(field_id), property_name, property_value)

    def get_property(self, field_id: str, property_name: str):
        return getattr(self._get_object(field_id), property_name)

    def invoke_method(self, field_id: str, method_name: str, *args):
        return getattr(self._get_object(field_id), method_name)(*args)
