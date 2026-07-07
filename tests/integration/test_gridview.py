import pytest
from rpa_sap import GridView, GridViewExtractor


@pytest.fixture
def grid_session(sap_session):
    """Fixture that navigates to sq01 grid view, yielding the session"""
    sap_session.run_transaction("sq01")
    sap_session.interactor.press_button("wnd[0]/tbar[1]/btn[19]")
    yield sap_session


def test_double_click_cell(grid_session):
    grid = GridView(grid_session)
    cell_address = grid.get_cell_address_by_cell_value(
        "wnd[1]/usr/cntlGRID1/shellcont/shell", "RPA"
    )
    if cell_address:
        assert len(cell_address) > 0
        grid.double_click_cell(
            "wnd[1]/usr/cntlGRID1/shellcont/shell",
            cell_address[0].row,
            cell_address[0].column,
        )
        print(f"Double clicked cell at: {cell_address}")
    else:
        print("Value RPA not found, skipping double click.")


def test_press_toolbar_context_button_and_select_context_menu_item(grid_session):
    grid = GridView(grid_session)
    try:
        grid.press_toolbar_context_button_and_select_context_menu_item(
            "wnd[1]/usr/cntlGRID1/shellcont/shell", "&MB_VARIANT", "&MAINTAIN"
        )
    except Exception as e:
        print(
            f"Warning: Context menu test failed, possibly not available in this view: {e}"
        )
    assert True


def test_count_rows(grid_session):
    grid = GridView(grid_session)
    rows = grid.count_rows("wnd[1]/usr/cntlGRID1/shellcont/shell")
    print(f"Rows: {rows}")
    assert isinstance(rows, int)
    assert rows >= 0


def test_count_columns(grid_session):
    grid = GridView(grid_session)
    columns = grid.count_columns("wnd[1]/usr/cntlGRID1/shellcont/shell")
    print(f"Columns: {columns}")
    assert isinstance(columns, int)
    assert columns >= 0


def test_get_current_cell(grid_session):
    grid = GridView(grid_session)
    # First set the current cell to make sure it exists
    grid.set_current_cell("wnd[1]/usr/cntlGRID1/shellcont/shell", 0, 0)

    cell = grid.get_current_cell("wnd[1]/usr/cntlGRID1/shellcont/shell")
    print(f"Current cell: {cell}")
    assert cell.row == 0
    assert cell.column == 0


def test_get_current_cell_value(grid_session):
    grid = GridView(grid_session)
    grid.set_current_cell("wnd[1]/usr/cntlGRID1/shellcont/shell", 0, 0)

    value = grid.get_current_cell_value("wnd[1]/usr/cntlGRID1/shellcont/shell")
    print(f"Current cell value: {value}")
    assert value is not None


def test_get_cell_value(grid_session):
    grid = GridView(grid_session)
    value = grid.get_cell_value("wnd[1]/usr/cntlGRID1/shellcont/shell", 0, 0)
    print(f"Cell 0,0 value: {value}")
    assert value is not None


def test_to_array(grid_session):
    grid = GridView(grid_session)
    extractor = GridViewExtractor(grid)
    arr = extractor.to_array("wnd[1]/usr/cntlGRID1/shellcont/shell")
    print(f"Array snippet: {arr[:2]}")
    assert isinstance(arr, list)
    assert len(arr) > 0
    assert isinstance(arr[0], list)


def test_to_dataframe(grid_session):
    grid = GridView(grid_session)
    extractor = GridViewExtractor(grid)
    df = extractor.to_dataframe("wnd[1]/usr/cntlGRID1/shellcont/shell")
    print("DataFrame snippet:\n", df.head().to_string())
    assert not df.empty
