from unittest import TestCase
from rpa_sap import SapGui

sap_gui = SapGui()

class TestSapGui(TestCase):
    def test_open_new_session(self):
        connection_string = '/H/oe_r3_copy_prod.euro.pilkington.net/S/3275'
        user_id = 'zg46002'
        password = 'misiaczek1'
        client = 900
        language = 'EN'
        test_value = sap_gui.open_new_session(connection_string, user_id, password, client, language)
        excpected_value = True

        sap_gui.close_session()
        sap_gui.close_sap_logon()
        
        self.assertEqual(test_value, excpected_value)

    def test_activate_session(self):
        sap_gui.activate_session(0,0)
        sap_gui.activate_session()