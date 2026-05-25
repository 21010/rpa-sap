from collections import namedtuple
from typing import overload, Optional, Union, Dict
from pandas import DataFrame
import numpy as np
import re
import math
import warnings
from ..core.session import SapSession

Column = namedtuple("Column", ["name", "title", "cells"])
Row = namedtuple("Row", ["index", "cells"])
Cell = namedtuple(
    "Cell",
    ["id", "row_index", "column_name", "column_title", "column_index", "type", "text"],
)


class GuiTableControl:
    def __init__(self, sap_session: SapSession):
        self.sap_session = sap_session

    def get_object(self, field_id: str):
        return self.sap_session.interactor._get_object(field_id)

    def count_columns(self, field_id: str) -> int:
        gui_table_control = self.get_object(field_id)
        return gui_table_control.Columns.Count

    def count_rows(self, field_id: str) -> int:
        gui_table_control = self.get_object(field_id)
        return gui_table_control.RowCount

    def count_visible_rows(self, field_id: str) -> int:
        gui_table_control = self.get_object(field_id)
        return gui_table_control.VisibleRowCount

    def get_table_header(self, field_id: str) -> Dict[int, Dict[str, str]]:
        """
        Returns a dictionary where the key is the Column Index (int)
        and the value is a dict containing 'name' and 'title'.
        This prevents overwriting columns with identical names.
        """
        gui_table_control = self.get_object(field_id)
        columns = {}

        # We iterate by index to ensure order and uniqueness
        for i in range(gui_table_control.Columns.Count):
            col_obj = gui_table_control.Columns.Item(i)
            columns[i] = {"name": col_obj.Name, "title": col_obj.Title}
        return columns

    # get column
    @overload
    def get_column(self, field_id: str, column_index: int) -> Column: ...
    @overload
    def get_column(self, field_id: str, column_name: str) -> Column: ...
    @overload
    def get_column(self, field_id: str, column_title: str) -> Column: ...

    def get_column(
        self,
        field_id: str,
        column_index: Optional[int] = None,
        column_name: Optional[str] = None,
        column_title: Optional[str] = None,
    ) -> Column:
        """
        Retrieves a specific column from a SAP GUI Table Control.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.
            column_index (Optional[int]): The index of the column. Defaults to None.
            column_name (Optional[str]): The name of the column. Defaults to None.
            column_title (Optional[str]): The title of the column. Defaults to None.

        Returns:
            Column: A named tuple containing column information and cell data.

        Note: Either column_index, column_name or column_title must be specified, but not both.
        """
        columns = self.get_table_header(field_id)
        target_col_index = -1

        # resolve column_name if column_title is provided
        if isinstance(column_index, int):
            target_col_index = column_index
        elif isinstance(column_title, str):
            for idx, data in columns.items():
                if data["title"] == column_title:
                    column_name = data["name"]
                    target_col_index = idx
                    break
        elif isinstance(column_name, str):
            for idx, data in columns.items():
                if data["name"] == column_name:
                    column_title = data["title"]
                    target_col_index = idx
                    break
        else:
            raise Exception(
                "Either column_index, column_name or column_title must be specified."
            )

        if target_col_index == -1:
            raise Exception(f"Column not found: {column_name or column_title}")

        # Get number of total rows and visible rows
        page_size = self.get_page_size(field_id)
        # initiate a list for elements of the table
        cells = []
        # calculate number of pages
        pages = self.count_pages(field_id)
        # iterate pages and read all cells
        for page in range(pages):
            # scroll to the first page
            if page == 0:
                self.get_object(field_id).VerticalScrollbar.Position = 0
            # iterate all visible cells
            for child in self.get_object(field_id).Children:
                column_index, row_index = self.__extract_coordinates__(child.Id)

                if column_index == target_col_index:
                    cells.append(
                        Cell(
                            id=self.__extract_field_id__(child.Id),
                            row_index=row_index + (page_size * page),
                            column_index=column_index,
                            column_name=column_name,
                            column_title=column_title,
                            type=child.Type,
                            text=child.Text,
                        )
                    )

            # scroll to the next page
            if page < pages - 1:
                self.get_object(field_id).VerticalScrollbar.Position = (
                    page + 1
                ) * page_size

        # return resuls
        return Column(name=column_name, title=column_title, cells=cells)

    # get rows
    def get_row(self, field_id: str, absolute_row_index: int) -> Row:
        """
        Retrieves a specific row from a SAP GUI Table Control.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.
            absolute_row_index (int): The absolute row index of the row.

        Returns:
            list[Cell]: A list of Cell objects representing each cell in the row.
        """
        # get table header
        columns = self.get_table_header(field_id)

        # scroll to the row
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        # get the row object
        row = self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
        # iterage through all cells and append to the list
        cells = []
        for cell in range(row.Count):
            column_index, _ = self.__extract_coordinates__(row[cell].Id)
            col_data = columns.get(column_index, {"name": row[cell].Name, "title": ""})
            cells.append(
                Cell(
                    id=self.__extract_field_id__(row[cell].Id),
                    row_index=absolute_row_index,
                    column_index=column_index,
                    column_name=col_data["name"],
                    column_title=col_data["title"],
                    type=row[cell].Type,
                    text=row[cell].Text,
                )
            )
        # return the results
        return Row(index=absolute_row_index, cells=cells)

    def get_rows(self, field_id: str, entire_table: bool = True) -> list[Row]:
        """
        Retrieves all rows from a SAP GUI Table Control.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.
            entire_table (bool): Whether to retrieve the entire table or only visible rows. Defaults to True.

        Returns:
            list[Row]: A list of Row objects representing each row in the table.
        """
        # initiate list
        rows = []

        # extract table
        table = (
            self.__extract_table__(field_id)
            if entire_table
            else self.__extract_visible_rows__(field_id)
        )
        # initiathe list of rows with unique indexes
        indexes = set()
        for cell in table:
            if cell["absolute_row_index"] not in indexes:
                indexes.add(cell["absolute_row_index"])
                rows.append(Row(index=cell["absolute_row_index"], cells=[]))

        # sort cells by row index
        rows.sort(key=lambda x: x.index)

        row_map = {row.index: row for row in rows}

        # iterate through all cells and append cell to appropiae row
        for cell in table:
            row_map[cell["absolute_row_index"]].cells.append(
                Cell(
                    id=cell["field_id"],
                    row_index=cell["absolute_row_index"],
                    column_index=cell["column_index"],
                    column_name=cell["name"],
                    column_title=cell["title"],
                    text=cell["text"],
                    type=cell["type"],
                )
            )

        # sort cells by column index and row index
        for row in rows:
            row.cells.sort(key=lambda x: x.column_index)

        return rows

    # get cell
    @overload
    def get_cell(
        self, field_id: str, absolute_row_index: int, column_name: str
    ) -> Cell: ...
    @overload
    def get_cell(
        self, field_id: str, absolute_row_index: int, column_title: str
    ) -> Cell: ...
    @overload
    def get_cell(self, field_id: str, value: str) -> list[Cell]: ...

    def get_cell(
        self,
        field_id: str,
        value: Optional[str] = None,
        absolute_row_index: Optional[int] = None,
        column_name: Optional[str] = None,
        column_title: Optional[str] = None,
    ) -> Union[list[Cell], Cell]:
        """
        Retrieves cell(s) from a SAP GUI Table Control based on various criteria.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.
            value (Optional[str]): Value to search for in cells. Defaults to None.
            absolute_row_index (Optional[int]): Absolute row index. Defaults to None.
            column_name (Optional[str]): Name of the column. Defaults to None.
            column_title (Optional[str]): Title of the column. Defaults to None.

        Returns:
            Union[list[Cell], Cell]: List of Cell objects if searching by value, single Cell object otherwise.

        Note: Use either 'value' to search across all cells, or provide 'absolute_row_index' along with either 'column_name' or 'column_title'.

        """
        columns = self.get_table_header(field_id)

        # SCENARIO 1: SEARCH BY VALUE
        if isinstance(value, str):
            pages = self.count_pages(field_id)
            page_size = self.get_page_size(field_id)
            cells = []

            for page in range(pages):
                if page == 0:
                    self.get_object(field_id).VerticalScrollbar.Position = 0

                for child in self.get_object(field_id).Children:
                    if child.Text.lower() == str(value).lower():
                        column_index, row_index = self.__extract_coordinates__(child.Id)
                        col_data = columns.get(
                            column_index, {"name": child.Name, "title": ""}
                        )

                        cells.append(
                            Cell(
                                id=self.__extract_field_id__(child.Id),
                                row_index=row_index + (page_size * page),
                                column_index=column_index,
                                column_name=col_data["name"],
                                column_title=col_data["title"],
                                text=child.Text,
                                type=child.Type,
                            )
                        )

                if page < pages - 1:
                    self.get_object(field_id).VerticalScrollbar.Position = (
                        page + 1
                    ) * page_size

            return cells

        # SCENARIO 2: SEARCH BY ROW INDEX AND COLUMN
        elif isinstance(absolute_row_index, int):
            if not (column_name or column_title):
                raise Exception("Either column_name or column_title must be specified.")

            # Find target column index
            target_col_index = -1
            if isinstance(column_title, str):
                for idx, data in columns.items():
                    if data["title"] == column_title:
                        column_name = data["name"]
                        target_col_index = idx
                        break
            elif isinstance(column_name, str):
                for idx, data in columns.items():
                    if data["name"] == column_name:
                        column_title = data["title"]
                        target_col_index = idx
                        break

            self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
            row = self.get_object(field_id).GetAbsoluteRow(absolute_row_index)

            for cell in range(row.Count):
                column_index, _ = self.__extract_coordinates__(row[cell].Id)
                # Compare by index if we found one, otherwise fall back to name check
                match = (
                    (column_index == target_col_index)
                    if target_col_index != -1
                    else (row[cell].Name == column_name)
                )

                if match:
                    return Cell(
                        id=self.__extract_field_id__(row[cell].Id),
                        row_index=absolute_row_index,
                        column_index=column_index,
                        column_name=row[cell].Name,
                        column_title=column_title,
                        type=row[cell].Type,
                        text=row[cell].Text,
                    )

        raise Exception("Invalid parameters.")

    # set cell value
    @overload
    def set_cell_value(
        self, field_id: str, value: str, absolute_row_index: int, column_name: str
    ): ...
    @overload
    def set_cell_value(
        self, field_id: str, value: str, absolute_row_index: int, column_title: str
    ): ...

    def set_cell_value(
        self,
        field_id: str,
        value: str,
        absolute_row_index: int,
        column_name: Optional[str] = None,
        column_title: Optional[str] = None,
    ):
        # Raise exception if column_name or column_title is not provided or both values are provided at the same time.
        if not (column_name or column_title):
            raise Exception("Either column_name or column_title must be specified.")

        # get table headers (columns)
        columns = self.get_table_header(field_id)
        target_col_index = -1

        if isinstance(column_title, str):
            for idx, data in columns.items():
                if data["title"] == column_title:
                    column_name = data["name"]
                    target_col_index = idx
                    break
        elif isinstance(column_name, str):
            for idx, data in columns.items():
                if data["name"] == column_name:
                    target_col_index = idx
                    break

        # scroll to the row
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        # get the row object
        row = self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
        # iterate through cells in row and return desired row as Cell object.
        for cell in range(row.Count):
            col_idx, _ = self.__extract_coordinates__(row[cell].Id)

            match = False
            if target_col_index != -1:
                if col_idx == target_col_index:
                    match = True
            elif row[cell].Name == column_name:
                match = True

            if match:
                if row[cell].Changeable:
                    row[cell].Text = str(value)
                else:
                    warnings.warn(f"Column {row[cell].Name} is read-only.")
                return

    # press cell object
    @overload
    def press_cell(self, field_id: str, absolute_row_index: int, column_name: str): ...
    @overload
    def press_cell(self, field_id: str, absolute_row_index: int, column_title: str): ...

    def press_cell(
        self,
        field_id: str,
        absolute_row_index: int,
        column_name: Optional[str] = None,
        column_title: Optional[str] = None,
    ):
        # Raise exception if column_name or column_title is not provided or both values are provided at the same time.
        if not (column_name or column_title):
            raise Exception("Either column_name or column_title must be specified.")

        # get table headers (columns)
        columns = self.get_table_header(field_id)

        columns = self.get_table_header(field_id)
        target_col_index = -1

        if isinstance(column_title, str):
            for idx, data in columns.items():
                if data["title"] == column_title:
                    column_name = data["name"]
                    target_col_index = idx
                    break
        elif isinstance(column_name, str):
            for idx, data in columns.items():
                if data["name"] == column_name:
                    target_col_index = idx
                    break

        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        row = self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
        for cell in range(row.Count):
            col_idx, _ = self.__extract_coordinates__(row[cell].Id)
            match = (
                (col_idx == target_col_index)
                if target_col_index != -1
                else (row[cell].Name == column_name)
            )

            if match:
                row[cell].Press()
                return

    # row selecting/deselecting
    def select_row(self, field_id: str, absolute_row_index: int):
        """
        Selects a specific row in a SAP GUI Table Control.

        Args:
            field_id (str): Field ID of the SAP GUI Table Control.
            absolute_row_index (int): Absolute row index to select.

        Returns:
            None

        Note: Scrolls to the specified row, then selects the target row.
        """
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        self.get_object(field_id).GetAbsoluteRow(absolute_row_index).Selected = True

    def deselect_row(self, field_id: str, absolute_row_index: int):
        """
        Deselects a specific row in a SAP GUI Table Control.

        Args:
            field_id (str): Field ID of the SAP GUI Table Control.
            absolute_row_index (int): Absolute row index to deselect.

        Returns:
            None

        Note: Scrolls to the specified row and removes its selection.
        """
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        self.get_object(field_id).GetAbsoluteRow(absolute_row_index).Selected = False

    # scrolling and paggination
    def scroll_to_nth_row(self, field_id: str, absolute_row_index: int):
        """
        Scrolls to a specific row in a SAP GUI Table Control.

        Args:
            field_id (str): Field ID of the SAP GUI Table Control.
            absolute_row_index (int): Absolute row index to scroll to.

        Returns:
            None

        Note: Adjusts the vertical scrollbar position to make the specified row visible.
        """
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index

    def count_pages(self, field_id: str) -> int:
        # calculate number of pages
        range = self.get_object(field_id).VerticalScrollbar.Range
        page_size = self.get_object(field_id).VerticalScrollbar.PageSize

        if range == 0:
            return 1

        return int(math.ceil(range / page_size))

    def get_page_size(self, field_id: str) -> int:
        return self.get_object(field_id).VerticalScrollbar.PageSize

    def scroll_to_nth_page(self, field_id: str, page: int):
        pages = self.count_pages(field_id)
        if page > pages or page < 1:
            raise Exception(f"Page index out of range. Range: 1:{pages}")

        page_size = self.get_page_size(field_id)
        self.get_object(field_id).VerticalScrollbar.Position = page_size * (page - 1)

    # data extraction
    def to_DataFrame(self, field_id: str, entire_table: bool = True) -> DataFrame:
        """
        Converts the SAP Table Control data into a pandas DataFrame.
        Handles duplicate column names/titles by mapping data based on column index.
        """
        # 1. Get headers indexed by position to ensure order and uniqueness
        columns_header = self.get_table_header(field_id)

        # 2. Sort indices to ensure we build the list in the correct visual order (0, 1, 2...)
        sorted_indices = sorted(columns_header.keys())

        # 3. Extract column titles for the DataFrame header
        # Note: Duplicate titles are allowed in this list.
        columns = [columns_header[i]["title"] for i in sorted_indices]

        # 4. Build the data rows
        data = []
        for row in self.get_rows(field_id, entire_table):
            # Create a row filled with None to handle potential sparse data or gaps
            row_data = [None] * len(columns)

            for cell in row.cells:
                # Map cell text to its specific column index
                # This prevents overwriting data if two columns share a name
                if cell.column_index < len(row_data):
                    row_data[cell.column_index] = cell.text

            data.append(row_data)

        # 5. Create DataFrame
        return DataFrame(data, columns=columns)

    def to_array(self, field_id: str) -> np.ndarray:
        table = self.__extract_table__(field_id)
        # Using indexes instead of names ensures uniqueness
        unique_col_indexes = sorted(list(set([x["column_index"] for x in table])))
        rows = sorted(list(set([x["absolute_row_index"] for x in table])))

        # Map absolute row index to array row index (0, 1, 2...)
        row_map = {idx: i for i, idx in enumerate(rows)}
        col_map = {idx: i for i, idx in enumerate(unique_col_indexes)}

        data = np.empty((len(rows), len(unique_col_indexes)), dtype=object)

        for cell in table:
            r = row_map[cell["absolute_row_index"]]
            c = col_map[cell["column_index"]]
            data[r, c] = cell["text"]

        return data

    def __extract_visible_rows__(self, field_id: str) -> list:
        columns = self.get_table_header(field_id)
        table = []
        gui_table_control = self.get_object(field_id)

        for child in gui_table_control.Children:
            col, row = self.__extract_coordinates__(child.Id)

            # Lookup metadata by INDEX
            col_data = columns.get(col, {"name": child.Name, "title": ""})

            table.append(
                {
                    "absolute_row_index": None,
                    "visible_row_index": row,
                    "column_index": col,
                    "title": col_data["title"],
                    "text": child.Text,
                    "name": col_data["name"],  # Or child.Name
                    "field_id": self.__extract_field_id__(child.Id),
                    "type": child.Type,
                }
            )
        return table

    def __extract_table__(self, field_id: str) -> list:
        columns = self.get_table_header(field_id)
        page_size = self.get_page_size(field_id)
        table = []
        pages = self.count_pages(field_id)

        for page in range(pages):
            if page == 0:
                self.get_object(field_id).VerticalScrollbar.Position = 0

            for child in self.get_object(field_id).Children:
                col, row = self.__extract_coordinates__(child.Id)

                # CRITICAL FIX: Use 'col' index to get metadata, not child.Name
                col_data = columns.get(col, {"name": child.Name, "title": ""})

                table.append(
                    {
                        "absolute_row_index": row + (page_size * page),
                        "visible_row_index": row,
                        "column_index": col,
                        "title": col_data["title"],
                        "text": child.Text,
                        "name": col_data["name"],
                        "field_id": self.__extract_field_id__(child.Id),
                        "type": child.Type,
                    }
                )

            if page < pages - 1:
                self.get_object(field_id).VerticalScrollbar.Position = (
                    page + 1
                ) * page_size

        return table

    def __extract_coordinates__(
        self, gui_table_control_cell_field_id: str
    ) -> Union[tuple[int, int], None]:
        match = re.search(r"\[(-?\d+),(-?\d+)\]$", gui_table_control_cell_field_id)
        if match:
            try:
                x = int(match.group(1))
                y = int(match.group(2))
                return (x, y)
            except ValueError:
                raise ValueError("Coordinates must be integers.")
        else:
            return None

    def __extract_field_id__(self, full_field_id: str) -> str:
        return full_field_id[full_field_id.index("wnd") :]
