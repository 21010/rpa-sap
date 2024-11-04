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
        
        session = SAPGUI.active_session

        session.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = "101"
        session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").text = "102"
        session.findById("wnd[0]/usr/ctxtSP$00005-LOW").text = "PL22"
        session.findById("wnd[0]/usr/ctxtSP$00008-LOW").text = "PL"
        session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").caretPosition = 3
        session.findById("wnd[0]").sendVKey(0)

        SAPGUI.sq01.execute_query(time_limit=10)

        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_close_process(self):
        SAPGUI.close_process(process_name='excel.exe')
