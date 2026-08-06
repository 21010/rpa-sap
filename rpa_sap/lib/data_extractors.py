from typing import List, Dict, Any
import numpy as np
from pandas import DataFrame
from abc import ABC, abstractmethod

# Assuming these are available, but to avoid circular imports, type hints might just use strings or generic types if needed.
# Since we just pass the object, duck typing works, or we can import them if there are no circular dependencies.
from .GridView import GridView
from .GuiTableControl import GuiTableControl


class BaseExtractor(ABC):
    """Abstract base class defining the standard interface for SAP data extractors."""

    @abstractmethod
    def to_dataframe(self, field_id: str, **kwargs) -> DataFrame:
        pass

    @abstractmethod
    def to_array(self, field_id: str, **kwargs) -> Any:
        pass

    def to_csv(self, field_id: str, path_or_buf: str, **kwargs) -> None:
        self.to_dataframe(field_id, **kwargs).to_csv(
            path_or_buf=path_or_buf, index=False
        )

    def to_xlsx(self, field_id: str, file_path: str, **kwargs) -> None:
        self.to_dataframe(field_id, **kwargs).to_excel(file_path, index=False)


class GridViewExtractor(BaseExtractor):
    """
    Handles data extraction and export from SAP GridView controls.
    """

    def __init__(self, grid_view: GridView):
        self.grid_view = grid_view

    def to_array(self, field_id: str, **kwargs: Any) -> List[List]:
        grid_obj = self.grid_view.get_object(field_id)
        return [self._get_headers(grid_obj), *self._get_body(grid_obj)]

    def to_dict(self, field_id: str, **kwargs: Any) -> Dict:
        grid_obj = self.grid_view.get_object(field_id)
        return {
            "columns": self._get_headers(grid_obj),
            "data": self._get_body(grid_obj),
        }

    def to_dataframe(self, field_id: str, **kwargs: Any) -> DataFrame:
        grid_obj = self.grid_view.get_object(field_id)
        return DataFrame(
            data=self._get_body(grid_obj), columns=self._get_headers(grid_obj)
        )

    def _get_headers(self, grid_obj) -> List:
        return [
            grid_obj.GetColumnTitles(column_name)[0]
            for column_name in grid_obj.ColumnOrder
        ]

    def _get_body(self, grid_obj) -> List:
        body = []
        for row_index in range(0, grid_obj.RowCount):
            row = []
            for column_index in range(0, grid_obj.ColumnCount):
                column_name = grid_obj.ColumnOrder[column_index]
                row.append(grid_obj.GetCellValue(row_index, column_name))
            body.append(row)
        return body


class GuiTableControlExtractor(BaseExtractor):
    """
    Handles data extraction and export from SAP GuiTableControl controls.
    """

    def __init__(self, table_control: GuiTableControl):
        self.table_control = table_control

    def to_dataframe(
        self, field_id: str, entire_table: bool = True, **kwargs: Any
    ) -> DataFrame:
        columns_header = self.table_control.get_table_header(field_id)
        sorted_indices = sorted(columns_header.keys())
        columns = [columns_header[i]["title"] for i in sorted_indices]

        data = []
        for row in self.table_control.get_rows(field_id, entire_table):
            row_data = [None] * len(columns)
            for cell in row.cells:
                if cell.column_index < len(row_data):
                    row_data[cell.column_index] = cell.text
            data.append(row_data)

        return DataFrame(data, columns=columns)

    def to_array(self, field_id: str, **kwargs: Any) -> np.ndarray:
        table = self.table_control.__extract_table__(field_id)
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
