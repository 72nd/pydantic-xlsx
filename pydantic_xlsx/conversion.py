"""
This module handles the representation of some additional types as a cell
value in a xlsx file. It also provides the needed import functionality. This
functionality should only be used for data types which need altering additional
cell properties like number format. When only the value of a cell is altered
there should be another way of implementing the needed functionality by using
classes extending the primitive types which can be represented in a Excel cell.
"""

from abc import ABCMeta, abstractmethod
from typing import Generic, Optional, TypeVar

from openpyxl.cell.cell import Cell


T = TypeVar('T')
"""A conversion object is always implemented for a specific type."""


class Conversion(Generic[T], metaclass=ABCMeta):
    """
    Defines a type/class which a specific representation in the Xlsx file. Use
    the `ConversionFactory` to obtain the type class for your type.
    """

    @abstractmethod
    def field_value(cell: Cell) -> T:
        """
        Converts the content of the Excel cell into the type defined by the
        class. This value can then be used to populate the model field.
        """
        pass

    @abstractmethod
    def populate_cell(cell: Cell, value: T):
        """
        Populates a given Xlsx cell with the with the given value. The
        representation of this value is defined by the Type subclass.
        """
        pass


class ConversionFactory:
    """
    Creates the correct Conversion implementation based on a given pydantic
    model field.
    """

    @classmethod
    def from_field(
        cls,
        field: T,
    ) -> Optional[Conversion]:
        """
        Determines based on a given pydantic Field (-type). If there is no
        implementation None will be returned.
        """
