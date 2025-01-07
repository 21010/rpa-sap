from collections import namedtuple
from typing import overload, Optional, Union, List
import win32com.client
from pandas import DataFrame
import numpy as np
import time
import re
import math
import warnings
from .. import SapGui

Column = namedtuple('Column', ['name', 'title', 'cells'])
Row = namedtuple('Row', ['index', 'cells'])
Cell = namedtuple('Cell', ['id', 'row_index', 'column_name', 'column_title', 'column_index', 'type', 'text'])

class GuiTableControl:
    def __init__(self, sap_gui: SapGui):
        self.sap_gui = sap_gui

    def get_object(self, field_id: str):
        return self.sap_gui.get_object(field_id)

    def count_columns(self, field_id: str) -> int:
        gui_table_control = self.get_object(field_id)
        return gui_table_control.Columns.Count
    
    def count_rows(self, field_id: str) -> int:
        gui_table_control = self.get_object(field_id)
        return gui_table_control.RowCount

    def count_visible_rows(self, field_id: str) -> int:
        gui_table_control = self.get_object(field_id)
        return gui_table_control.VisibleRowCount

    def get_table_header(self, field_id: str) -> dict:
        gui_table_control = self.get_object(field_id)
        columns = dict()
        for column in gui_table_control.Columns:
            columns[column.Name] = column.Title
        return columns

    # get column
    @overload
    def get_column(self, field_id: str, column_name: str) -> Column:
        ...
    @overload
    def get_column(self, field_id: str, column_title: str) -> Column:
        ...

    def get_column(self, field_id: str, column_name: Optional[str] = None, column_title: Optional[str] = None) -> Column:
        """
        Retrieves a specific column from a SAP GUI Table Control.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.
            column_name (Optional[str]): The name of the column. Defaults to None.
            column_title (Optional[str]): The title of the column. Defaults to None.

        Returns:
            Column: A named tuple containing column information and cell data.

        Note: Either column_name or column_title must be specified, but not both.
        """
        columns = self.get_table_header(field_id)

        # resolve column_name if column_title is provided
        if isinstance(column_title, str):
            for name, title in columns.items():
                if title == column_title:
                    column_name = name   
        # resolve column_title if column_name is provided
        elif isinstance(column_name, str):
            column_title = columns[column_name]
        else:
            raise Exception("Either column_name or column_title must be specified, but not both.")

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
                if child.Name == column_name:
                    # get visible row and column indexes
                    column_index, row_index = self.__extract_coordinates__(child.Id)
                    # append values to the list
                    cells.append(Cell(
                        id=self.__extract_field_id__(child.Id),
                        row_index= row_index + (page_size * page),
                        column_index=column_index,
                        column_name=column_name,
                        column_title=column_title,
                        type=child.Type,
                        text=child.Text
                    ))

            # scroll to the next page
            if page < pages - 1:
                self.get_object(field_id).VerticalScrollbar.Position = (page + 1) * page_size
        
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
        row =  self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
        # iterage through all cells and append to the list
        cells = []
        for cell in range(row.Count):
            column_index, _ = self.__extract_coordinates__(row[cell].Id)
            cells.append(Cell(
                id=self.__extract_field_id__(row[cell].Id),
                row_index=absolute_row_index,
                column_index=column_index,
                column_name=row[cell].Name,
                column_title=columns[row[cell].Name],
                type=row[cell].Type,
                text=row[cell].Text
            ))
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
        table = self.__extract_table__(field_id) if entire_table else self.__extract_visible_rows__(field_id)
        # initiathe list of rows with unique indexes
        indexes = set()
        for cell in table:
            if cell['absolute_row_index'] not in indexes:
                indexes.add(cell['absolute_row_index'])
                rows.append(Row(
                    index=cell['absolute_row_index'],
                    cells=[]
                ))

        # sort cells by row index
        rows.sort(key=lambda x: x.index)

        # iterate through all cells and append cell to appropiae row
        for cell in table:
            rows[cell['absolute_row_index']].cells.append(Cell(
                id=cell['field_id'],
                row_index=cell['absolute_row_index'],
                column_index=cell['column_index'],
                column_name=cell['name'],
                column_title=cell['title'],
                text=cell['text'],
                type=cell['type']
            ))

        # sort cells by column index and row index
        for row in rows:
            row.cells.sort(key=lambda x: x.column_index)

        return rows

    # get cell
    @overload
    def get_cell(self, field_id: str, absolute_row_index: int, column_name: str) -> Cell:
        ...
    @overload
    def get_cell(self, field_id: str, absolute_row_index: int, column_title: str) -> Cell:
        ...
    @overload
    def get_cell(self, field_id: str, value: str) -> list[Cell]:
        ...
        
    def get_cell(self, field_id: str, value: Optional[str] = None, absolute_row_index: Optional[int] = None, 
                 column_name: Optional[str] = None, column_title: Optional[str] = None) -> Union[list[Cell], Cell]:
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
        
        if(isinstance(value, str)):
            pages = self.count_pages(field_id)
            page_size = self.get_page_size(field_id)
        
            cells = []
            
            for page in range(pages):
                if page == 0:
                    self.get_object(field_id).VerticalScrollbar.Position = 0
                
                for child in self.get_object(field_id).Children:
                    if child.Text.lower() == str(value).lower():
                        column_index, row_index = self.__extract_coordinates__(child.Id)
                        cells.append(Cell(
                            id=self.__extract_field_id__(child.Id),
                            row_index=row_index + (page_size * page),
                            column_index=column_index,
                            column_name=child.Name,
                            column_title=columns[child.Name],
                            text=child.Text,
                            type=child.Type
                        ))
                
                if page < pages - 1:  
                    self.get_object(field_id).VerticalScrollbar.Position = (page + 1) * page_size
            
            return cells
        
        elif(isinstance(absolute_row_index, int) and (isinstance(column_name, str) or isinstance(column_title, str))):
            # Raise exception if column_name or column_title is not provided or both values are provided at the same time.
            if column_name is None and column_title is None:
                raise Exception(f'Either column_name or column_title must be specified.')
            if column_name is not None and column_title is not None:
                raise Exception(f'Provide column_name or column_title, not both at a time.')
        
            # resolve column_name if column_title is provided
            if isinstance(column_title, str):
                for name, title in columns.items():
                    if title == column_title:
                        column_name = name  
        
            # resolve column_title if column_name is provided
            if isinstance(column_name, str):
                column_title = columns[column_name]

            # scroll to the row 
            self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
            # get the row object
            row =  self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
            # iterate through cells in row and return desired row as Cell object.
            for cell in range(row.Count):
                column_index, _ = self.__extract_coordinates__(row[cell].Id)
                if row[cell].Name == column_name:
                    return Cell(
                        id=self.__extract_field_id__(row[cell].Id),
                        row_index=absolute_row_index,
                        column_index=column_index,
                        column_name=row[cell].Name,
                        column_title=column_title,
                        type=row[cell].Type,
                        text=row[cell].Text
                    )
        
        raise Exception("Invalid combination of parameters. Specify either a value, or both absolute_row_index and one of column_name/column_title.")

    # set cell value
    @overload
    def set_cell_value(self, field_id: str, value: str, absolute_row_index: int, column_name: str):
        ...
    @overload
    def set_cell_value(self, field_id: str, value: str, absolute_row_index: int, column_title: str):
        ...
    
    def set_cell_value(self, field_id: str, value: str, absolute_row_index: int, column_name: Optional[str] = None, column_title: Optional[str] = None):
        # Raise exception if column_name or column_title is not provided or both values are provided at the same time.
        if column_name is None and column_title is None:
            raise Exception(f'Either column_name or column_title must be specified.')
        if column_name is not None and column_title is not None:
            raise Exception(f'Provide column_name or column_title, not both at a time.')

        # get table headers (columns)
        columns = self.get_table_header(field_id)

        # resolve column_name if column_title is provided
        if isinstance(column_title, str):
            for name, title in columns.items():
                if title == column_title:
                    column_name = name  
    
        # resolve column_title if column_name is provided
        if isinstance(column_name, str):
            column_title = columns[column_name]

        # scroll to the row 
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        # get the row object
        row =  self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
        # iterate through cells in row and return desired row as Cell object.
        for cell in range(row.Count):
            if row[cell].Name == column_name:
                if row[cell].Changeable:
                    row[cell].Text = str(value)
                else:
                    warnings.warn(f'Column {row[cell].Name} is read-only. Cannot modify the value.')

    # press cell object
    @overload
    def press_cell(self, field_id: str, absolute_row_index: int, column_name: str):
        ...
    @overload
    def press_cell(self, field_id: str, absolute_row_index: int, column_title: str):
        ...
    
    def press_cell(self, field_id: str, absolute_row_index: int, column_name: Optional[str] = None, column_title: Optional[str] = None):
        # Raise exception if column_name or column_title is not provided or both values are provided at the same time.
        if column_name is None and column_title is None:
            raise Exception(f'Either column_name or column_title must be specified.')
        if column_name is not None and column_title is not None:
            raise Exception(f'Provide column_name or column_title, not both at a time.')

        # get table headers (columns)
        columns = self.get_table_header(field_id)

        # resolve column_name if column_title is provided
        if isinstance(column_title, str):
            for name, title in columns.items():
                if title == column_title:
                    column_name = name  
    
        # resolve column_title if column_name is provided
        if isinstance(column_name, str):
            column_title = columns[column_name]

        # scroll to the row 
        self.get_object(field_id).VerticalScrollbar.Position = absolute_row_index
        # get the row object
        row =  self.get_object(field_id).GetAbsoluteRow(absolute_row_index)
        # iterate through cells in row and return desired row as Cell object.
        for cell in range(row.Count):
            if row[cell].Name == column_name:
                row[cell].Press()

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
        return int(math.ceil(range / page_size))

    def get_page_size(self, field_id: str) -> int:
        return self.get_object(field_id).VerticalScrollbar.PageSize

    def scroll_to_nth_page(self, field_id: str, page: int):
        pages = self.count_pages(field_id)
        if page > pages or page < 1:
            raise Exception(f'Page index out of range. Range: 1:{pages}')
        
        page_size = self.get_page_size(field_id)
        self.get_object(field_id).VerticalScrollbar.Position = page_size * (page-1)
        
    # data extraction
    def to_DataFrame(self, field_id: str, entire_table: bool = True) -> DataFrame:
        """
        Converts a SAP GUI Table Control to a pandas DataFrame.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.

        Returns:
            pd.DataFrame: A pandas DataFrame representation of the table data.
        """
        # extract table
        table = self.__extract_table__(field_id) if entire_table else self.__extract_visible_rows__(field_id)
        # convert outcome to DataFrame and return the data
        data = DataFrame(table)
        results = data.pivot(index='absolute_row_index' if entire_table else 'visible_row_index', columns='title', values='text')
        return results

    def to_array(self, field_id: str) -> np.ndarray:
        """
        Converts a SAP GUI Table Control to a 2D numpy array.

        Args:
            field_id (str): The field ID of the SAP GUI Table Control.

        Returns:
            np.ndarray: A 2D numpy array representation of the table data.
        """
        # extract table
        table = self.__extract_table__(field_id)
        # get unique column names
        columns = list(set([x['title'] for x in table]))
        # get unique row indexes
        rows = list(set([x['absolute_row_index'] for x in table]))
        # create an empty array
        data = np.empty((len(rows), len(columns)), dtype=object)
        # iterate through the table and fill the array
        for row in rows:
            for column in columns:
                for cell in table:
                    if cell['absolute_row_index'] == row and cell['title'] == column:
                        data[row, columns.index(column)] = cell['text']
        # return the results
        return data

    
    def __extract_visible_rows__(self, field_id: str) -> list:
        # Get table columns (names means field names, titles means displayed column headers)
        columns = self.get_table_header(field_id)
        
        # initiate a list for elements of the table
        table = []
        
        gui_table_control = self.get_object(field_id)
        # iterate all visible cells
        for child in gui_table_control.Children:
            # get visible row and column indexes
            col, row = self.__extract_coordinates__(child.Id)
            # append values to the table
            table.append(
                {
                    'absolute_row_index': None,
                    'visible_row_index': row,
                    'column_index': col,
                    'title': columns[child.Name],
                    'text': child.Text,
                    'name': child.Name,
                    'field_id': self.__extract_field_id__(child.Id),
                    'type': child.Type
                }
            )
        
        # return resuls
        return table

    def __extract_table__(self, field_id: str) -> list:
        # Get table columns (names means field names, titles means displayed column headers)
        columns = self.get_table_header(field_id)
        
        # Get number of visible rows (page size)
        page_size = self.get_page_size(field_id)
        
        # initiate a list for elements of the table
        table = []
        # calculate number of pages
        pages = self.count_pages(field_id)
        # iterate pages and read all cells
        for page in range(pages):
            # scroll to the first page
            if page == 0:
                self.get_object(field_id).VerticalScrollbar.Position = 0
            # iterate all visible cells
            for child in self.get_object(field_id).Children:
                # get visible row and column indexes
                col, row = self.__extract_coordinates__(child.Id)
                # append values to the table
                table.append(
                    {
                        'absolute_row_index': row + (page_size * page),
                        'visible_row_index': row,
                        'column_index': col,
                        'title': columns[child.Name],
                        'text': child.Text,
                        'name': child.Name,
                        'field_id': self.__extract_field_id__(child.Id),
                        'type': child.Type
                    }
                )
            
            # scroll to the next page
            if page < pages - 1:
                self.get_object(field_id).VerticalScrollbar.Position = (page + 1) * page_size
        
        # return resuls
        return table

    def __extract_coordinates__(self, gui_table_control_cell_field_id: str) -> tuple[int, int]:
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
        return full_field_id[full_field_id.index('wnd'):]
