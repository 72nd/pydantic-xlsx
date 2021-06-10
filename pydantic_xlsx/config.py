"""
Provides additional configuration capabilities for the `model.XlsxModel`.
"""

from typing import Optional

from openpyxl.styles import Alignment, Font
from openpyxl import utils as pyxl_utils
from pydantic import BaseConfig


class XlsxConfig(BaseConfig):
    """
    Extends pydantic's config class with some Excel specific stuff.
    """

    header_font: Optional[Font] = Font(
        name="Arial",
        bold=True,
    )
    """
    Font of the header row (first row). Defaults to:
    `Font(name="Arial", bold=True)`.
    """
    font: Optional[Font] = None
    """Font for the non header rows. Defaults to `None`."""
    header_alignment: Optional[Alignment] = None
    """
    Optional alignment for the header cells. Defaults to `None`."""
    freeze_cell: Optional[str] = "A2"
    """
    Define a cell coordinate where the cells should be freeze. The same value
    is used to calculate the rows and columns which should be repeated on each
    printed page. Defaults to `A2` (header row (i.e. first row) stick to the
    top).

    Per default this value is also used to calculate
    `XlsxConfig.print_title_columns` and `XlsxConfig.print_title_rows` (rows
    and columns which should repeated on each printed page)
    """
    ignore_additional_columns: bool = False
    """
    When true, additional row exceeding the models definition will be ignored
    when importing data from a existing xlsx file. The presence of such rows
    will otherwise lead to an validation error of pydantic. Use this option
    with caution.
    """
    disable_width_calculation: bool = False
    """
    The library uses a primitive algorithm to set the width for each column
    based on number of chars of the longest cell content.
    """
    print_horizontal_centered: bool = True
    """
    Whether to horizontally center the content when printing the document.
    """
    print_vertical_centered: bool = True
    """
    Whether to vertically center the content when printing the document.
    """
    print_title_columns: Optional[str] = None
    """
    Defines the range of columns which should be repeated on each page in
    the print out. Per default this is calculated based on the value of
    `XlsxConfig.freeze_cell`. To disable the feature set this property to an
    empty string (`""`). You can define individual column ranges using the
    format `START_COLUMN:END_COLUMN`. E.g. `A:B` will repeat the first two
    columns.
    """
    print_title_rows: Optional[str] = None
    """
    Defines the range of rows which should be repeated on each page in the
    print out. Per default this is calculated based on the value of
    `XlsxConfig.freeze_cell`. To disable the feature set this property to
    an empty string (`""`). You can define individual row ranges using the
    format `START_ROW:END_ROW`. E.g. `1:2` will repeat the first two rows.
    """

    @classmethod
    def _print_title_columns(cls) -> Optional[str]:
        if cls.freeze_cell is None:
            return None
        freeze_column =\
            pyxl_utils.cell.coordinate_from_string(cls.freeze_cell)[0]
        max_column_num =\
            pyxl_utils.cell.column_index_from_string(freeze_column) - 1
        if max_column_num == 0:
            return None
        max_column = pyxl_utils.cell.get_column_letter(max_column_num)
        return f"A:{max_column}"

    @classmethod
    def _print_title_rows(cls) -> Optional[str]:
        if cls.print_title_rows is not None:
            return cls.print_title_rows
        if cls.freeze_cell is None:
            return None
        freeze_row =\
            pyxl_utils.cell.coordinate_from_string(cls.freeze_cell)[1] - 1
        if freeze_row == 0:
            return None
        return f"1:{freeze_row}"
