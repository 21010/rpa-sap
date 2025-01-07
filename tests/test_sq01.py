""" SAPGUI tests """
from unittest import TestCase
import json
from rpa_sap import SapGui

SAPGUI: SapGui = SapGui()
with open('tests/credentials.json') as data:
    SECRETS = json.load(data)

class TestSQ01(TestCase):
    def test_sq01(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])

        SAPGUI.sq01.start_query('PAL-EXPIMP-V0', 'BREXIT')
        
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_close_process(self):
        SAPGUI.close_process(process_name='excel.exe')
