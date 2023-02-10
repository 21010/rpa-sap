""" SAPGUI tests """
from unittest import TestCase
import json
from rpa_sap import SapGui

SAPGUI: SapGui = SapGui()
with open('tests/credentials.json') as data:
    SECRETS = json.load(data)

class TestSapGui(TestCase):
    def test_open_and_close_new_session(self):
        test_value = SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        excpected_value = True

        SAPGUI.close_session()
        SAPGUI.close_sap_logon()
        
        self.assertEqual(test_value, excpected_value)

    def test_activate_session(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        SAPGUI.activate_session(0,0)
        self.assertEqual(SAPGUI.get_session_index(), 0)
        self.assertEqual(SAPGUI.get_connection_index(), 0)

        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        SAPGUI.activate_session()
        self.assertEqual(SAPGUI.get_session_index(), 0)
        self.assertEqual(SAPGUI.get_connection_index(), 1)

        SAPGUI.close_all_sessions()
        SAPGUI.close_sap_logon()

    def test_check_if_session_exists(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        self.assertEqual(SAPGUI.check_if_session_exists(), True)
        self.assertEqual(SAPGUI.check_if_session_exists(0), True)
        self.assertEqual(SAPGUI.check_if_session_exists(1), False)
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

        self.assertEqual(SAPGUI.check_if_session_exists(), False)
        self.assertEqual(SAPGUI.check_if_session_exists(0), False)
        
    def test_get_session_info(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        session_info = SAPGUI.get_session_info()
        self.assertEqual(session_info['connection index'], '/app/con[0]')
        print(session_info)
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_count_connections(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        self.assertEqual(SAPGUI.count_connections(), 1)
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        self.assertEqual(SAPGUI.count_connections(), 2)
        SAPGUI.close_all_sessions()
        SAPGUI.close_sap_logon()

    def test_count_sessions(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        self.assertEqual(SAPGUI.count_sessions(), 1)
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_is_session_busy(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        self.assertEqual(SAPGUI.is_session_busy(), False)
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_is_active(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        self.assertEqual(SAPGUI.is_session_active(), True)
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_wait_until_object_exists(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        SAPGUI.run_transaction('su3')
        SAPGUI.wait_until_object_exists('wnd[1]')
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

class TestGridView(TestCase):
    def test_double_click_gridview_cell(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        SAPGUI.run_transaction('sq01')
        SAPGUI.press_button('wnd[0]/tbar[1]/btn[19]')
        cell_address = SAPGUI.grid_view.get_cell_address_by_cell_value('wnd[1]/usr/cntlGRID1/shellcont/shell', 'SO99')
        SAPGUI.grid_view.double_click_cell('wnd[1]/usr/cntlGRID1/shellcont/shell', cell_address[0].row, cell_address[0].column)
        print(cell_address)
