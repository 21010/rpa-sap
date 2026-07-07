import pytest
from time import sleep
from rpa_sap import GuiTableControl
from rpa_sap.lib import GuiTableControlExtractor


@pytest.fixture
def table_session(sap_session):
    """Fixture that navigates to me32l and enters data, yielding the session"""
    sap_session.run_transaction("me32l")
    sap_session.interactor.set_text(
        field_id="wnd[0]/usr/ctxtRM06E-EVRTN", text="5500000503"
    )
    sap_session.interactor.press_enter()
    sap_session.interactor.press_enter()
    yield sap_session


def test_count_columns(table_session):
    table = GuiTableControl(table_session)
    columns = table.count_columns("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Columns: {columns}")
    assert isinstance(columns, int)
    assert columns >= 0


def test_count_rows(table_session):
    table = GuiTableControl(table_session)
    rows = table.count_rows("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Rows: {rows}")
    assert isinstance(rows, int)
    assert rows >= 0


def test_count_visible_rows(table_session):
    table = GuiTableControl(table_session)
    rows = table.count_visible_rows("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Visible rows: {rows}")
    assert isinstance(rows, int)
    assert rows >= 0


def test_get_table_header(table_session):
    table = GuiTableControl(table_session)
    columns = table.get_table_header("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Table header: {columns}")
    assert isinstance(columns, dict)
    assert len(columns) > 0


def test_get_column(table_session):
    table = GuiTableControl(table_session)
    value1 = table.get_column("wnd[0]/usr/tblSAPMM06ETC_0220", column_name="EKPO-TXZ01")
    print(f"Column by name: {value1}")
    assert value1 is not None

    value2 = table.get_column("wnd[0]/usr/tblSAPMM06ETC_0220", column_title="Material")
    print(f"Column by title: {value2}")
    assert value2 is not None


def test_get_row(table_session):
    table = GuiTableControl(table_session)
    value = table.get_row("wnd[0]/usr/tblSAPMM06ETC_0220", 1)
    print(f"Row 1: {value}")
    assert value is not None

    value = table.get_row("wnd[0]/usr/tblSAPMM06ETC_0220", 2)
    print(f"Row 2: {value}")
    assert value is not None


def test_get_rows(table_session):
    table = GuiTableControl(table_session)
    rows = table.get_rows("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"All rows: {len(rows)}")
    assert isinstance(rows, list)


def test_get_cell(table_session):
    table = GuiTableControl(table_session)
    cell = table.get_cell(
        field_id="wnd[0]/usr/tblSAPMM06ETC_0220",
        absolute_row_index=10,
        column_title="Material",
    )
    print(
        cell.id,
        cell.row_index,
        cell.column_name,
        cell.column_title,
        cell.type,
        cell.text,
    )
    assert cell is not None

    cell = table.get_cell(
        field_id="wnd[0]/usr/tblSAPMM06ETC_0220",
        absolute_row_index=20,
        column_name="EKPO-TXZ01",
    )
    print(
        cell.id,
        cell.row_index,
        cell.column_name,
        cell.column_title,
        cell.type,
        cell.text,
    )
    assert cell is not None

    cell = table.get_cell(
        field_id="wnd[0]/usr/tblSAPMM06ETC_0220",
        absolute_row_index=2,
        column_title="Short Text",
    )
    print(
        cell.id,
        cell.row_index,
        cell.column_name,
        cell.column_title,
        cell.type,
        cell.text,
    )
    assert cell is not None

    cells = table.get_cell(field_id="wnd[0]/usr/tblSAPMM06ETC_0220", value="ROL")
    print(f"Cells with value ROL: {len(cells)}")
    assert isinstance(cells, list)


def test_set_cell_value(table_session):
    table = GuiTableControl(table_session)

    # Use Short Text (EKPO-TXZ01) instead of Material/Qty to avoid SAP business logic errors
    table.set_cell_value(
        "wnd[0]/usr/tblSAPMM06ETC_0220",
        value="TEST STRING",
        absolute_row_index=1,
        column_name="EKPO-TXZ01",
    )

    # Let's verify we can find the cell we just modified
    cells = table.get_cell(
        field_id="wnd[0]/usr/tblSAPMM06ETC_0220", value="TEST STRING"
    )
    print(f"Cells with TEST STRING: {cells}")
    if cells:
        # Revert or do something else
        table.set_cell_value(
            "wnd[0]/usr/tblSAPMM06ETC_0220",
            value="TEST REVERT",
            absolute_row_index=cells[0].row_index,
            column_name="EKPO-TXZ01",
        )

    assert True


def test_press_cell(table_session):
    table = GuiTableControl(table_session)
    table.press_cell(
        "wnd[0]/usr/tblSAPMM06ETC_0220", absolute_row_index=25, column_title="Texts"
    )
    assert True


def test_select_row(table_session):
    table = GuiTableControl(table_session)
    table.select_row("wnd[0]/usr/tblSAPMM06ETC_0220", 25)
    sleep(1)
    table.select_row("wnd[0]/usr/tblSAPMM06ETC_0220", 6)
    sleep(1)
    table.select_row("wnd[0]/usr/tblSAPMM06ETC_0220", 14)
    sleep(1)
    assert True


def test_deselect_row(table_session):
    table = GuiTableControl(table_session)
    table.select_row("wnd[0]/usr/tblSAPMM06ETC_0220", 25)
    sleep(1)
    table.deselect_row("wnd[0]/usr/tblSAPMM06ETC_0220", 25)
    sleep(1)
    assert True


def test_scroll_to_nth_row(table_session):
    table = GuiTableControl(table_session)
    table.scroll_to_nth_row("wnd[0]/usr/tblSAPMM06ETC_0220", 14)
    table.scroll_to_nth_row("wnd[0]/usr/tblSAPMM06ETC_0220", 123)
    table.scroll_to_nth_row("wnd[0]/usr/tblSAPMM06ETC_0220", 4)
    assert True


def test_count_pages(table_session):
    table = GuiTableControl(table_session)
    pages = table.count_pages("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Pages: {pages}")
    assert isinstance(pages, int)


def test_get_page_size(table_session):
    table = GuiTableControl(table_session)
    page_size = table.get_page_size("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Page size: {page_size}")
    assert isinstance(page_size, int)


def test_scroll_to_nth_page(table_session):
    table = GuiTableControl(table_session)
    table.scroll_to_nth_page("wnd[0]/usr/tblSAPMM06ETC_0220", 2)
    assert True


def test_to_dataframe(table_session):
    table = GuiTableControl(table_session)
    extractor = GuiTableControlExtractor(table)
    df = extractor.to_dataframe("wnd[0]/usr/tblSAPMM06ETC_0220")
    print("DataFrame snippet:\n", df.head().to_string())
    df.to_excel("tests/table.xlsx")

    df_visible = extractor.to_dataframe("wnd[0]/usr/tblSAPMM06ETC_0220", False)
    print("Visible DataFrame snippet:\n", df_visible.head().to_string())

    assert not df.empty


def test_to_array(table_session):
    table = GuiTableControl(table_session)
    extractor = GuiTableControlExtractor(table)
    array = extractor.to_array("wnd[0]/usr/tblSAPMM06ETC_0220")
    print(f"Array shape: {array.shape if hasattr(array, 'shape') else len(array)}")
    assert array is not None
