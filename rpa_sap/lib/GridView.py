from time import sleep
from collections import namedtuple
import win32com.client
from pandas import DataFrame

Cell_Address = namedtuple('Cell_Address', ['row', 'column'])

class GridView:
    def __init__(self, sap_gui):
        self.sap_gui = sap_gui

    def get_object(self, field_id: str):
        return self.sap_gui.object(field_id)

    def count_rows(self, grid_view_id: str) -> int:
        """
        Count row of GridView object.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number of rows
        """
        grid_view = self.get_object(grid_view_id)
        return grid_view.RowCount

    def count_columns(self, grid_view_id: str) -> int:
        """
        Count columns of GridView object.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number of columns
        """
        grid_view = self.get_object(grid_view_id)
        return grid_view.ColumnCount

    def get_current_cell_value(self, grid_view_id: str) -> object:
        """
        Return the value of current GridView cell.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            object: Value of GridView cell
        """
        grid_view = self.get_object(grid_view_id)
        return grid_view.GetCellValue(grid_view.CurrentCellRow, grid_view.CurrentCellColumn)

    def get_current_cell(self, grid_view_id: str):
        """
        Return row index and column index of current GridView cell.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            GridViewCell['row', 'column']: object with row and column attributes.
        """
        grid_view = self.get_object(grid_view_id)
        GridViewCell = namedtuple('GridViewCell', ['row', 'column'])
        return GridViewCell(grid_view.CurrentCellRow, self.__get_column_index__(grid_view, grid_view.CurrentCellColumn))

    def set_current_cell(self, grid_view_id: str, row_index: int, column_index: int):
        """
        Set current cell of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            row_index (int): Row index.
            column_index (int): Column index.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.SetCurrentCell(row_index, self.__get_column_name__(grid_view, column_index))

    def get_current_column_name(self, grid_view_id: str) -> str:
        """
        Return the name of current column of the GridView object.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            str: column name
        """
        grid_view = self.get_object(grid_view_id)
        return grid_view.CurrentCellColumn

    def set_current_column_name(self, grid_view_id: str, column_name: str):
        """
        Set current column of the GridView by column name

        Args:
            grid_view_id (str): GridView field id.
            column_name (str): Column name.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.CurrentCellColumn = column_name

    def get_current_column_index(self, grid_view_id: str) -> int:
        """
        Return index of current GridView column.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number value.
        """
        grid_view = self.get_object(grid_view_id)
        for column_index in range(0, grid_view.ColumnOrder.Count):
            if grid_view.ColumnOrder[column_index] == grid_view.CurrentCellColumn:
                return column_index

    def set_current_column_index(self, grid_view_id: str, column_index: int):
        """
        Set the index of current column of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            column_index (int): Column Index
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.CurrentCellColumn = self.__get_column_name__(grid_view, column_index)

    def get_current_row_index(self, grid_view_id: str) -> int:
        """
        Return the index of current GridView row.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            int: number value
        """
        grid_view = self.get_object(grid_view_id)
        return grid_view.CurrentCellRow

    def set_current_row_index(self, grid_view_id: str, row_index: int):
        """
        Set the index of current row of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            row_index (int): Row index.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.CurrentCellRow = row_index

    def get_selected_rows(self, grid_view_id: str) -> list:
        """
        Return indexes of selected GridView rows.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            list: list of selected row indexes
        """
        grid_view = self.get_object(grid_view_id)
        selected_rows: str = str(grid_view.SelectedRows)
        if selected_rows == "":
            return None
        rows_list: list = []
        for row in selected_rows.split(','):
            if '-' in row:
                index_range: list[str] = row.split('-')
                for index in range(index_range[0], index_range[1]):
                    rows_list.append(index)
            rows_list.append(int(row))
        return rows_list

    def set_selected_rows(self, grid_view_id: str, row_indexes: list[int] | str):
        """
        Set selected rows of GridView object.

        Args:
            grid_view_id (str): GridView field id
            row_indexes (list[int] | str): can be a str, ex. "1", or "1,2" or "1-3" if you want to select a range, or the list of int ex. [1,2,3]
        """
        selected_rows: str
        if isinstance(row_indexes, str):
            selected_rows = row_indexes
        if isinstance(row_indexes, list[int]):
            selected_rows = ','.join([str(item) for item in row_indexes])

        grid_view = self.get_object(grid_view_id)
        grid_view.SelectedRows(selected_rows)

    def clear_selection(self, grid_view_id: str):
        """
        Clear row selection of the GridView object.

        Args:
            grid_view_id (str): GridView field id.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.ClearSelection()

    def double_click_cell(self, grid_view_id: str, row_index: int = None, column_index: int = None):
        """
        Double click the cell of GridView object.

        Args:
            grid_view_id (str): GridView field id
            row_index (int, optional): _description_. Defaults to None.
            column_index (int, optional): _description_. Defaults to None.
        """
        grid_view = self.get_object(grid_view_id)
        if row_index is not None or column_index is not None:
            column_name: str = self.__get_column_name__(grid_view, self.get_current_column_index(grid_view_id)) if column_index is None else self.__get_column_name__(grid_view, column_index)
            row_index = self.get_current_row_index(grid_view_id) if row_index is None else row_index
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index
            grid_view.selectedRows = row_index
        grid_view.DoubleClickCurrentCell()

    def convert_column_index_to_name(self, grid_view_id: str, column_name: str) -> int:
        grid_view = self.get_object(grid_view_id)
        column_index: int
        for column_index in range(0, grid_view.ColumnCount):
            if column_name == grid_view.ColumnOrder[column_index]:
                return column_index

    def get_cell_address_by_cell_value(self, grid_view_id: str, cell_value: str) -> list[Cell_Address]:
        """ Return the list of Cell_Address[row, column] objects

        Args:
            grid_view_id (str): Field id
            cell_value (str): searched value

        Returns:
            list[Cell_Address]: Cell_Address object with parameters: row and column
        
        Usage:
            cell_address = sap.grid_view.get_cell_address_by_cell_value('wnd[0]/shell', 'test)\n\r
            cell_address[0].row    # contains the index of a row for the first matched cell\n\r
            cell_address[0].column # contains the index of a column for the first matched cell\n\r
        """
        grid_view = self.get_object(grid_view_id)
        indexes = self.__get_cell_address_by_value__(grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(f'The GridView row not found for the value: {cell_value}')
        return indexes

    def get_cell_state(self, grid_view_id: str, row_index: int = None, column_index: int = None) -> str:
        grid_view = self.get_object(grid_view_id)
        r_index = row_index if row_index is not None else self.get_current_row_index
        c_index = column_index if column_index is not None else self.get_current_column_index
        return grid_view.GetCellState(r_index, self.__get_column_name__(grid_view, c_index))

    def get_cell_value(self, grid_view_id: str, row_index: int = None, column_index: int = None) -> object:
        grid_view = self.get_object(grid_view_id)
        r_index = row_index if row_index is not None else self.get_current_row_index
        c_index = column_index if column_index is not None else self.get_current_column_index
        return grid_view.GetCellValue(r_index, self.__get_column_name__(grid_view, c_index))

    def press_toolbar_button(self, grid_view_id: str, button_id: str):
        grid_view = self.get_object(grid_view_id)
        grid_view.pressToolbarButton(button_id)

    def press_toolbar_context_button(self, grid_view_id: str, button_id: str):
        grid_view = self.get_object(grid_view_id)
        grid_view.pressToolbarContextButton(button_id)

    def press_toolbar_context_button_and_select_context_menu_item(self, grid_view_id: str, button_id: str, function_code: str):
        grid_view = self.get_object(grid_view_id)
        grid_view.pressToolbarContextButton(button_id)
        sleep(1)
        grid_view.selectContextMenuItem(function_code)
        grid_view.ActiveWindow.setFocus()

    def select_all_cells(self, grid_view_id: str):
        grid_view = self.get_object(grid_view_id)
        grid_view.SelectAll()

    def select_column(self, grid_view_id: str, column_index: int):
        grid_view = self.get_object(grid_view_id)
        grid_view.SelectColumn(self.__get_column_name__(grid_view, column_index))

    def select_context_menu_item(self, grid_view_id: str, function_code: str):
        grid_view = self.get_object(grid_view_id)
        grid_view.selectContextMenuItem(function_code)

    def select_rows_by_cell_value(self, grid_view_id: str, cell_value: object):
        grid_view = self.get_object(grid_view_id)
        indexes = self.__get_cell_address_by_value__(grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception('The GridView row not found for the value: %s' % cell_value)

        for row_index, column_index in indexes:
            column_name = self.__get_column_name__(grid_view, column_index)
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index

        grid_view.selectedRows = ','.join([str(r) for r, c in indexes])

    def set_current_cell_by_cell_value(self, grid_view_id: str, cell_value: object):
        grid_view = self.get_object(grid_view_id)
        indexes = self.__get_cell_address_by_value__(grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(f'The GridView row not found for the value: {cell_value}')

        for row_index, column_index in indexes:
            column_name = self.__get_column_name__(grid_view, column_index)
            grid_view.SetCurrentCell(row_index, column_name)

    def to_array(self, grid_view_id: str) -> list:
        grid_view = self.get_object(grid_view_id)
        return [self.__get_headers__(grid_view), *self.__get_body__(grid_view)]

    def to_dict(self, grid_view_id: str) -> dict:
        grid_view = self.get_object(grid_view_id)
        return {'columns': self.__get_headers__(grid_view), 'data': self.__get_body__(grid_view)}

    def to_dataframe(self, grid_view_id: str) -> DataFrame:
        grid_view = self.get_object(grid_view_id)
        return DataFrame(data=self.__get_body__(grid_view), columns=self.__get_headers__(grid_view))

    def to_csv(self, grid_view_id: str, path_or_buf: str):
        grid_view = self.get_object(grid_view_id)
        self.to_dataframe(grid_view).to_csv(
            path_or_buf=path_or_buf, index=False)

    def to_xlsx(self, grid_view_id: str, file_path: str):
        grid_view = self.get_object(grid_view_id)
        self.to_dataframe(grid_view).to_excel(file_path, index=False)


    # Magic methods - Grid View
    def __get_column_index__(self, grid_view: win32com.client.dynamic.CDispatch, column_name: str):
        for column_index in range(0, grid_view.ColumnOrder.Count):
            return column_index if column_name == grid_view.ColumnOrder[column_index] else None

    def __get_column_name__(self, grid_view: win32com.client.dynamic.CDispatch, column_index: int) -> str:
        return grid_view.ColumnOrder[column_index]

    def __get_cell_address_by_value__(self, grid_view: win32com.client.dynamic.CDispatch, cell_value: object):
        results = []
        for row_index in range(0, grid_view.RowCount):
            for column_index in range(0, grid_view.ColumnOrder.Count):
                if cell_value == grid_view.GetCellValue(row_index, grid_view.ColumnOrder(column_index)):
                    results.append(Cell_Address(row_index, column_index))
        return results

    def __get_headers__(self, grid_view: win32com.client.dynamic.CDispatch) -> list:
        return [grid_view.GetColumnTitles(column_name)[0] for column_name in grid_view.ColumnOrder]

    def __get_body__(self, grid_view: win32com.client.dynamic.CDispatch) -> list:
        body = []
        for row_index in range(0, grid_view.RowCount):
            row = []
            for column_index in range(0, grid_view.ColumnCount):
                row.append(grid_view.GetCellValue(
                    row_index, self.__get_column_name__(grid_view, column_index)))
            body.append(row)
        return body
