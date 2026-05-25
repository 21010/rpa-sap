from typing import Optional, List, Dict
from collections import namedtuple
import win32com.client

from pandas import DataFrame

from .. import SapGui

Cell_Address = namedtuple("Cell_Address", ["row", "column"])


class GridView:
    def __init__(self, sap_gui: SapGui):
        self.sap_gui = sap_gui

    def get_object(self, field_id: str):
        return self.sap_gui.get_object(field_id)

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
        return grid_view.GetCellValue(
            grid_view.CurrentCellRow, grid_view.CurrentCellColumn
        )

    def get_current_cell(self, grid_view_id: str):
        """
        Return row index and column index of current GridView cell.

        Args:
            grid_view_id (str): GridView field id.

        Returns:
            GridViewCell['row', 'column']: object with row and column attributes.
        """
        grid_view = self.get_object(grid_view_id)
        GridViewCell = namedtuple("GridViewCell", ["row", "column"])
        return GridViewCell(
            grid_view.CurrentCellRow,
            self.__get_column_index(grid_view, grid_view.CurrentCellColumn),
        )

    def set_current_cell(self, grid_view_id: str, row_index: int, column_index: int):
        """
        Set current cell of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            row_index (int): Row index.
            column_index (int): Column index.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.SetCurrentCell(
            row_index, self.__get_column_name(grid_view, column_index)
        )

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
        return -1

    def set_current_column_index(self, grid_view_id: str, column_index: int):
        """
        Set the index of current column of GridView object.

        Args:
            grid_view_id (str): GridView field id.
            column_index (int): Column Index
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.CurrentCellColumn = self.__get_column_name(grid_view, column_index)

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
            return []
        rows_list: list = []
        for row in selected_rows.split(","):
            if "-" in row:
                index_range: list[str] = row.split("-")
                for index in range(int(index_range[0]), int(index_range[1])):
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
        if isinstance(row_indexes, str):
            selected_rows = row_indexes
        elif isinstance(row_indexes, list):
            selected_rows = ",".join([str(item) for item in row_indexes])
        else:
            raise TypeError("row_indexes must be a list of integers or a string.")

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

    def click_cell(
        self,
        grid_view_id: str,
        row_index: Optional[int] = None,
        column_index: Optional[int] = None,
    ):
        """
        Click the cell of GridView object.

        Args:
            grid_view_id (str): GridView field id
            row_index (int, optional): Row index. Defaults to None.
            column_index (int, optional): Column index. Defaults to None.
        """
        grid_view = self.get_object(grid_view_id)
        if row_index is not None or column_index is not None:
            column_name: str = (
                self.__get_column_name(
                    grid_view, self.get_current_column_index(grid_view_id)
                )
                if column_index is None
                else self.__get_column_name(grid_view, column_index)
            )
            row_index = (
                self.get_current_row_index(grid_view_id)
                if row_index is None
                else row_index
            )
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index
            grid_view.selectedRows = row_index
        grid_view.ClickCurrentCell()

    def double_click_cell(
        self,
        grid_view_id: str,
        row_index: Optional[int] = None,
        column_index: Optional[int] = None,
    ):
        """
        Double click the cell of GridView object.

        Args:
            grid_view_id (str): GridView field id
            row_index (int, optional): Row index. Defaults to None.
            column_index (int, optional): Column index. Defaults to None.
        """
        grid_view = self.get_object(grid_view_id)
        if row_index is not None or column_index is not None:
            column_name: str = (
                self.__get_column_name(
                    grid_view, self.get_current_column_index(grid_view_id)
                )
                if column_index is None
                else self.__get_column_name(grid_view, column_index)
            )
            row_index = (
                self.get_current_row_index(grid_view_id)
                if row_index is None
                else row_index
            )
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index
            grid_view.selectedRows = row_index
        grid_view.DoubleClickCurrentCell()

    def convert_column_name_to_index(self, grid_view_id: str, column_name: str) -> int:
        """
        Returns column index by given column name.

        Args:
            grid_view_id (str): GridView field id
            column_name (str): Column name

        Returns:
            int: Column index
        """
        grid_view = self.get_object(grid_view_id)

        for column_index in range(0, grid_view.ColumnCount):
            if column_name == grid_view.ColumnOrder[column_index]:
                return column_index
        return -1

    def get_cell_address_by_cell_value(
        self,
        grid_view_id: str,
        cell_value: str,
    ) -> List[Cell_Address]:
        """Return the list of Cell_Address[row, column] objects

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
        indexes = self.__get_cell_address_by_value(grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(f"The GridView row not found for the value: {cell_value}")
        return indexes

    def get_cell_state(
        self,
        grid_view_id: str,
        row_index: Optional[int] = None,
        column_index: Optional[int] = None,
    ) -> str:
        """
        Returns the state of a specific cell in the GridView.

        Args:
            grid_view_id (str): The field ID of the GridView control.
            row_index (Optional[int]): The row index of the cell. If None, the current row index is used.
            column_index (Optional[int]): The column index of the cell. If None, the current column index is used.

        Returns:
            str: returns the state of the cell (e.g., checkbox checked/unchecked, editable, selected).
        """
        grid_view = self.get_object(grid_view_id)
        r_index = row_index if row_index is not None else self.get_current_row_index
        c_index = (
            column_index if column_index is not None else self.get_current_column_index
        )
        return grid_view.GetCellState(
            r_index, self.__get_column_name(grid_view, c_index)
        )

    def get_cell_value(
        self,
        grid_view_id: str,
        row_index: Optional[int] = None,
        column_index: Optional[int] = None,
    ) -> object:
        """Return the value of the GridView cell

        Args:
            grid_view_id (str): Field id
            row_index (int, optional): Row index. Defaults to None.
            column_index (int, optional): Column index. Defaults to None.

        Returns:
            object: value (Any)
        """
        grid_view = self.get_object(grid_view_id)
        r_index = row_index if row_index is not None else self.get_current_row_index
        c_index = (
            column_index if column_index is not None else self.get_current_column_index
        )
        return grid_view.GetCellValue(
            r_index, self.__get_column_name(grid_view, c_index)
        )

    def press_toolbar_button(self, grid_view_id: str, button_id: str):
        """
        Presses a button on the GridView toolbar.

        Args:
            grid_view_id (str): GridView field id.
            button_id (str): The ID of the button to press.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.pressToolbarButton(button_id)

    def press_toolbar_context_button(self, grid_view_id: str, button_id: str):
        """
        Presses a context button on the GridView toolbar.

        Args:
            grid_view_id (str): GridView field id.
            button_id (str): The ID of the context button to press.
        """
        grid_view = self.get_object(grid_view_id)
        grid_view.pressToolbarContextButton(button_id)

    def press_toolbar_context_button_and_select_context_menu_item(
        self,
        grid_view_id: str,
        button_id: str,
        function_code: str,
    ):
        """
        Presses a context button on the GridView toolbar and then selects a menu item from the opened context menu.

        Args:
            grid_view_id (str): GridView field id.
            button_id (str): The ID of the context button to press.
            function_code (str): The function code of the menu item to select.
        """

        grid_view = self.get_object(grid_view_id)
        grid_view.pressToolbarContextButton(button_id)
        grid_view.selectContextMenuItem(function_code)

    def select_all_cells(self, grid_view_id: str):
        """
        Selects all cells in the GridView.

        Args:
            grid_view_id (str): GridView field id.
        """

        grid_view = self.get_object(grid_view_id)
        grid_view.SelectAll()

    def select_column(self, grid_view_id: str, column_index: int):
        """
        Selects a specific column in the GridView.

        Args:
            grid_view_id (str): GridView field id.
            column_index (int): The index of the column to select.
        """

        grid_view = self.get_object(grid_view_id)
        grid_view.SelectColumn(self.__get_column_name(grid_view, column_index))

    def select_context_menu_item(self, grid_view_id: str, function_code: str):
        """
        Selects a context menu item from the GridView.

        Args:
            grid_view_id (str): GridView field id.
            function_code (str): The function code of the menu item to select.
        """

        grid_view = self.get_object(grid_view_id)
        grid_view.ContextMenu()
        grid_view.selectContextMenuItem(function_code)

    def select_rows_by_cell_value(self, grid_view_id: str, cell_value: object):
        """
        Selects rows in the GridView based on a given cell value.

        Args:
            grid_view_id (str): GridView field id.
            cell_value (object): The value to search for in cells.
        """

        grid_view = self.get_object(grid_view_id)
        indexes = self.__get_cell_address_by_value(grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(f"The GridView row not found for the value: {cell_value}")

        for row_index, column_index in indexes:
            column_name = self.__get_column_name(grid_view, column_index)
            grid_view.SetCurrentCell(row_index, column_name)
            grid_view.currentCellRow = row_index

        grid_view.selectedRows = ",".join([str(r) for r, c in indexes])

    def set_current_cell_by_cell_value(self, grid_view_id: str, cell_value: object):
        """
        Sets the current cell in the GridView by searching for a specific cell value.
        If multiple cells contain the value, the last one found will be set as current.

        Args:
            grid_view_id (str): GridView field id.
            cell_value (object): The value to search for in cells.

        Raises:
            Exception: If no cell containing the specified value is found.
        """

        grid_view = self.get_object(grid_view_id)
        indexes = self.__get_cell_address_by_value(grid_view, cell_value)
        if len(indexes) == 0:
            raise Exception(f"The GridView row not found for the value: {cell_value}")

        for row_index, column_index in indexes:
            column_name = self.__get_column_name(grid_view, column_index)
            grid_view.SetCurrentCell(row_index, column_name)

    def to_array(self, grid_view_id: str) -> List[List]:
        """
        Extracts all data from the GridView and returns it as a list of lists.
        The first sub-list contains the column headers, and subsequent sub-lists
        represent the rows of data.

        Args:
            grid_view_id (str): The ID of the GridView control.

        Returns:
            list: A list of lists, where the first inner list is headers and the rest are rows.
        """

        grid_view = self.get_object(grid_view_id)
        return [self.__get_headers(grid_view), *self.__get_body(grid_view)]

    def to_dict(self, grid_view_id: str) -> Dict:
        """
        Extracts all data from the GridView and returns it as a dictionary.
        The dictionary contains two keys: "columns" for the column headers
        and "data" for the rows of data.

        Args:
            grid_view_id (str): The ID of the GridView control.

        Returns:
            dict: A dictionary with "columns" (list of headers) and "data" (list of lists for rows).
        """

        grid_view = self.get_object(grid_view_id)
        return {
            "columns": self.__get_headers(grid_view),
            "data": self.__get_body(grid_view),
        }

    def to_DataFrame(self, grid_view_id: str) -> DataFrame:
        """
        Extracts all data from the GridView and returns it as a pandas DataFrame.
        The DataFrame columns are set to the GridView headers.

        Args:
            grid_view_id (str): The ID of the GridView control.

        Returns:
            DataFrame: A pandas DataFrame containing the GridView data.
        """
        grid_view = self.get_object(grid_view_id)
        return DataFrame(
            data=self.__get_body(grid_view), columns=self.__get_headers(grid_view)
        )

    def to_csv(self, grid_view_id: str, path_or_buf: str):
        """
        Exports the GridView data to a CSV file.

        Args:
            grid_view_id (str): The ID of the GridView control.
            path_or_buf (str): The file path or buffer to write the CSV data to.
        """

        grid_view = self.get_object(grid_view_id)
        self.to_DataFrame(grid_view).to_csv(path_or_buf=path_or_buf, index=False)

    def to_xlsx(self, grid_view_id: str, file_path: str):
        """
        Exports the GridView data to an XLSX file.

        Args:
            grid_view_id (str): The ID of the GridView control.
            file_path (str): The full path to the output XLSX file.
        """
        grid_view = self.get_object(grid_view_id)
        self.to_DataFrame(grid_view).to_excel(file_path, index=False)

    # Magic methods - Grid View

    def __get_column_index(
        self,
        grid_view: win32com.client.dynamic.CDispatch,
        column_name: str,
    ):
        """
        Returns column index by given column name.

        Args:
            grid_view (win32com.client.dynamic.CDispatch): GridView object.
            column_name (str): Column name.

        Returns:
            int: Column index
        """

        for column_index in range(0, grid_view.ColumnOrder.Count):
            return (
                column_index
                if column_name == grid_view.ColumnOrder[column_index]
                else None
            )

    def __get_column_name(
        self,
        grid_view: win32com.client.dynamic.CDispatch,
        column_index: int,
    ) -> str:
        """
        Returns column name by given column index.

        Args:
            grid_view (win32com.client.dynamic.CDispatch): GridView object.
            column_index (int): Column index.

        Returns:
            str: Column name
        """

        return grid_view.ColumnOrder[column_index]

    def __get_cell_address_by_value(
        self,
        grid_view: win32com.client.dynamic.CDispatch,
        cell_value: object,
    ) -> List[Cell_Address]:
        """
        Returns a list of Cell_Address[row, column] objects for cells containing the specified value.

        Args:
            grid_view (win32com.client.dynamic.CDispatch): GridView object.
            cell_value (object): The value to search for in cells.

        Returns:
            list[Cell_Address]: A list of Cell_Address objects with row and column attributes.
        """

        results = []
        for row_index in range(0, grid_view.RowCount):
            for column_index in range(0, grid_view.ColumnOrder.Count):
                if cell_value == grid_view.GetCellValue(
                    row_index, grid_view.ColumnOrder(column_index)
                ):
                    results.append(Cell_Address(row_index, column_index))
        return results

    def __get_headers(self, grid_view: win32com.client.dynamic.CDispatch) -> List:
        """
        Extracts the column headers from the GridView object.

        Args:
            grid_view (win32com.client.dynamic.CDispatch): GridView object.

        Returns:
            List: A list of column header names.
        """

        return [
            grid_view.GetColumnTitles(column_name)[0]
            for column_name in grid_view.ColumnOrder
        ]

    def __get_body(self, grid_view: win32com.client.dynamic.CDispatch) -> List:
        """
        Extracts the data rows from the GridView object.

        Args:
            grid_view (win32com.client.dynamic.CDispatch): GridView object.

        Returns:
            List: A list of lists, where each inner list represents a row of data.
        """

        body = []
        for row_index in range(0, grid_view.RowCount):
            row = []
            for column_index in range(0, grid_view.ColumnCount):
                row.append(
                    grid_view.GetCellValue(
                        row_index, self.__get_column_name(grid_view, column_index)
                    )
                )
            body.append(row)
        return body
