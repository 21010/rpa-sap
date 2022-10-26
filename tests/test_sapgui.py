""" SAPGUI tests """
from unittest import TestCase
import json
from rpa_sap import SapGui

SAPGUI: SapGui = SapGui()
SECRETS = json.load('credentials.json')

class TestSapGui(TestCase):
    def test_open_and_close_new_session(self):
        test_value = SAPGUI.open_new_session(
            connection_string=SECRETS['connection_string'],
            user_id=SECRETS['user_id'],
            password=SECRETS['password'],
            client=SECRETS['client'],
            language=SECRETS['language']
        )
        excpected_value = True

        SAPGUI.close_session()
        SAPGUI.close_sap_logon()
        
        self.assertEqual(test_value, excpected_value)

    def test_activate_session(self):
        SAPGUI.activate_session(0,0)
        SAPGUI.activate_session()