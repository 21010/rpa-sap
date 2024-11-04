""" SAPGUI tests """
from unittest import TestCase
import json
from rpa_sap import SapGui

SAPGUI: SapGui = SapGui()
with open('tests/credentials.json') as data:
    SECRETS = json.load(data)

class TestGridView(TestCase):
    def test_double_click_cell(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        SAPGUI.run_transaction('sq01')
        SAPGUI.press_button('wnd[0]/tbar[1]/btn[19]')
        cell_address = SAPGUI.grid_view.get_cell_address_by_cell_value('wnd[1]/usr/cntlGRID1/shellcont/shell', 'SO99')
        SAPGUI.grid_view.double_click_cell('wnd[1]/usr/cntlGRID1/shellcont/shell', cell_address[0].row, cell_address[0].column)
        print(cell_address)
    
    def test_press_toolbar_context_button_and_select_context_menu_item(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])
        SAPGUI.run_transaction('sq01')

        SAPGUI.grid_view.press_toolbar_context_button_and_select_context_menu_item('wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell', '&MB_VARIANT', '&MAINTAIN')
        
