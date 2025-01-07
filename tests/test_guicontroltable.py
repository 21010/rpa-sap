""" SAPGUI tests """
from unittest import TestCase
import json
from time import sleep
from rpa_sap import SapGui

SAPGUI: SapGui = SapGui()
with open('tests/credentials.json') as data:
    SECRETS = json.load(data)

class TestGuiControlTable(TestCase):
    def test_count_columns(self):
        self.__use_me32l__()

        columns = SAPGUI.gui_table_control.count_columns('wnd[0]/usr/tblSAPMM06ETC_0220')
        print(columns)
        
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_count_rows(self):
        self.__use_me32l__()

        rows = SAPGUI.gui_table_control.count_rows('wnd[0]/usr/tblSAPMM06ETC_0220')
        print(rows)
        
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_count_visible_rows(self):
        self.__use_me32l__()

        rows = SAPGUI.gui_table_control.count_visible_rows('wnd[0]/usr/tblSAPMM06ETC_0220')
        print(rows)
        
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_count_pages(self):
        try:
            self.__use_me32l__()

            pages = SAPGUI.gui_table_control.count_pages('wnd[0]/usr/tblSAPMM06ETC_0220')
            print(pages)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_get_table_header(self):
        self.__use_me32l__()

        columns = SAPGUI.gui_table_control.get_table_header('wnd[0]/usr/tblSAPMM06ETC_0220')
        print(columns)
        
        SAPGUI.close_session()
        SAPGUI.close_sap_logon()

    def test_get_column(self):
        try:
            self.__use_me32l__()

            value = SAPGUI.gui_table_control.get_column('wnd[0]/usr/tblSAPMM06ETC_0220', column_name='EKPO-TXZ01')
            print(value)

            value = SAPGUI.gui_table_control.get_column('wnd[0]/usr/tblSAPMM06ETC_0220', column_title='Material')
            print(value)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()  
    
    def test_get_row(self):
        try:
            self.__use_me32l__()

            value = SAPGUI.gui_table_control.get_row('wnd[0]/usr/tblSAPMM06ETC_0220', 3)
            print(value)
            value = SAPGUI.gui_table_control.get_row('wnd[0]/usr/tblSAPMM06ETC_0220', 35)
            print(value)
            value = SAPGUI.gui_table_control.get_row('wnd[0]/usr/tblSAPMM06ETC_0220', 103)
            print(value)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()  

    def test_get_rows(self):
        try:
            self.__use_me32l__()
            rows = SAPGUI.gui_table_control.get_rows('wnd[0]/usr/tblSAPMM06ETC_0220')
            print(rows)
        except Exception as ex:
            raise ex
        finally:  
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_get_cell(self):
        try:
            self.__use_me32l__()

            cell = SAPGUI.gui_table_control.get_cell(field_id='wnd[0]/usr/tblSAPMM06ETC_0220', absolute_row_index=100, column_title='Material')
            print(cell.id, cell.row_index, cell.column_name, cell.column_title, cell.type, cell. text)
            cell = SAPGUI.gui_table_control.get_cell(field_id='wnd[0]/usr/tblSAPMM06ETC_0220', absolute_row_index=25, column_name='EKPO-TXZ01')
            print(cell.id, cell.row_index, cell.column_name, cell.column_title, cell.type, cell. text)
            cell = SAPGUI.gui_table_control.get_cell(field_id='wnd[0]/usr/tblSAPMM06ETC_0220', absolute_row_index=45, column_title='Short Text')
            print(cell.id, cell.row_index, cell.column_name, cell.column_title, cell.type, cell. text)
            cells = SAPGUI.gui_table_control.get_cell(field_id='wnd[0]/usr/tblSAPMM06ETC_0220', value='ROL')
            print(cells)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()  
  
    def test_set_cell_value(self):
        try:
            self.__use_me32l__()

            SAPGUI.gui_table_control.set_cell_value('wnd[0]/usr/tblSAPMM06ETC_0220', value='100000', absolute_row_index=25, column_title='Material')
            SAPGUI.gui_table_control.set_cell_value('wnd[0]/usr/tblSAPMM06ETC_0220', value='100', absolute_row_index=25, column_title='Targ. Qty')

            cell = SAPGUI.gui_table_control.get_cell(field_id='wnd[0]/usr/tblSAPMM06ETC_0220', value='200008392')
            print(cell)
            SAPGUI.gui_table_control.set_cell_value('wnd[0]/usr/tblSAPMM06ETC_0220', value='200', absolute_row_index=cell[0].row_index, column_title='Targ. Qty')
            
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_press_cell(self):
        try:
            self.__use_me32l__()

            SAPGUI.gui_table_control.press_cell('wnd[0]/usr/tblSAPMM06ETC_0220', absolute_row_index=25, column_title='Texts')

        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_select_row(self):
        try:
            self.__use_me32l__()
            SAPGUI.gui_table_control.select_row('wnd[0]/usr/tblSAPMM06ETC_0220', 25)
            sleep(3)
            SAPGUI.gui_table_control.select_row('wnd[0]/usr/tblSAPMM06ETC_0220', 26)
            sleep(3)
            SAPGUI.gui_table_control.select_row('wnd[0]/usr/tblSAPMM06ETC_0220', 24)
            sleep(3)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_deselect_row(self):
        try:
            self.__use_me32l__()
            SAPGUI.gui_table_control.select_row('wnd[0]/usr/tblSAPMM06ETC_0220', 25)
            sleep(3)
            SAPGUI.gui_table_control.deselect_row('wnd[0]/usr/tblSAPMM06ETC_0220', 25)
            sleep(3)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_scroll_to_nth_row(self):
        try:
            self.__use_me32l__()

            SAPGUI.gui_table_control.scroll_to_nth_row('wnd[0]/usr/tblSAPMM06ETC_0220', 14)
            SAPGUI.gui_table_control.scroll_to_nth_row('wnd[0]/usr/tblSAPMM06ETC_0220', 123)
            SAPGUI.gui_table_control.scroll_to_nth_row('wnd[0]/usr/tblSAPMM06ETC_0220', 4)
            
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()  

    def test_count_pages(self):
        try:
            self.__use_me32l__()

            pages = SAPGUI.gui_table_control.count_pages('wnd[0]/usr/tblSAPMM06ETC_0220')
            print(pages)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()
    
    def test_get_page_size(self):
        try:
            self.__use_me32l__()

            pages = SAPGUI.gui_table_control.get_page_size('wnd[0]/usr/tblSAPMM06ETC_0220')
            print(pages)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_scroll_to_nth_page(self):
        try:
            self.__use_me32l__()

            SAPGUI.gui_table_control.scroll_to_nth_page('wnd[0]/usr/tblSAPMM06ETC_0220', 2)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_to_dataframe(self):
        try:
            self.__use_me32l__()

            value = SAPGUI.gui_table_control.to_DataFrame('wnd[0]/usr/tblSAPMM06ETC_0220')
            print(value.to_string())
            value.to_excel('tests/table.xlsx')
            value = SAPGUI.gui_table_control.to_DataFrame('wnd[0]/usr/tblSAPMM06ETC_0220', False)
            print(value.to_string())
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def test_to_array(self):
        try:
            self.__use_me32l__()
            array = SAPGUI.gui_table_control.to_array('wnd[0]/usr/tblSAPMM06ETC_0220')            
            print(array)
        except Exception as ex:
            raise ex
        finally:
            SAPGUI.close_session()
            SAPGUI.close_sap_logon()

    def __use_me32l__(self):
        SAPGUI.open_new_session(SECRETS['connection_string'], SECRETS['user_id'], SECRETS['password'], SECRETS['client'], SECRETS['language'])

        SAPGUI.run_transaction('me32l')
        SAPGUI.set_text(field_id='wnd[0]/usr/ctxtRM06E-EVRTN', text='5500077529')
        SAPGUI.press_enter()
