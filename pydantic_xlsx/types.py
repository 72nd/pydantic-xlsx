"""
This module contains additional data types which are important when working
with Excel data sources like a primitive Money type.
"""

from abc import ABCMeta, abstractmethod
from typing import Optional


class Money(float, metaclass=ABCMeta):
    """
    Handles amounts of money by subclassing float. In general it's a very bad
    idea to store amounts of money as floats as all kind of funny things can
    happen. But as openpyxl parses currency cell values as int or float we have
    to work with it anyway.

    Depending on the user input in the Excel file. The input float can have any
    number of decimal places. The value is rounded according to the number of
    minor units of the given currency and then converted to an integer.

    To define a money field in your model you first have to define the currency
    you like to use. For this subclass the Money class and set some class
    variables:

    ```
    class Euro(Money):
        minor_unit = 2
        code = "€"
        code_before_amount = False
        delimiter = ","
        thousands_separator = "."
    ```
    """

    amount: int
    """The amount of money."""

    @property
    @classmethod
    @abstractmethod
    def minor_unit(cls) -> int:
        """
        Expresses the decimal relationship between the currency and it's minor
        unit. 1 means a ratio of 10:1, 2 equals to 100:1 and so on. For example
        the European currency "Euro" has a minor unit 2 as one Euro is made up
        of 100 cents.
        """
        pass

    @property
    @classmethod
    @abstractmethod
    def code(cls) -> str:
        """Freely chosen code to represent your currency."""
        pass

    @property
    @classmethod
    @abstractmethod
    def code_before_amount(cls) -> bool:
        """The position of the currency code."""
        pass

    @property
    @classmethod
    @abstractmethod
    def delimiter(cls) -> bool:
        """
        Delimiter used to distinguish between the currency and it's minor unit.
        """
        pass

    @property
    @classmethod
    @abstractmethod
    def thousands_separator(cls) -> str:
        """
        Separator used to group thousands.
        """
        pass

    def __new__(cls, value: float):
        return float.__new__(cls, value)

    def __init__(self, value: float) -> None:
        normalized = "{1:.{0}f}".format(self.minor_unit, value)
        self.amount = int(normalized.replace(".", ""))

    @classmethod
    def validate(cls, value: float):
        return cls(value)

    @classmethod
    def __get_validators__(cls):
        yield cls.validate

    @classmethod
    def number_format(cls) -> str:
        """Returns the Excel number format code for the currency."""
        # Defines how to display the thousands and minors of a number.
        # Ex.: `#,##0,00`
        decimal_seperation = "#{}{}0.{}".format(
            cls.delimiter,
            "#" * cls.minor_unit,
            "0" * cls.minor_unit,
        )
        if cls.code_before_amount:
            amount = "{} {}".format(cls.code, decimal_seperation)
        else:
            amount = "{} {}".format(decimal_seperation, cls.code)
        rsl = "{};[RED]-{}".format(amount, amount)
        return rsl

    def __str__(self) -> str:
        minor_amount = self.amount % (10 ** self.minor_unit)
        minor = "{1:0{0}d}".format(self.minor_unit, minor_amount)
        integer_amount = self.amount - minor_amount
        integer_thousands = "{:,}".format(integer_amount)
        integer = integer_thousands.replace(",", self.thousands_separator)

        number = "{}{}{}".format(integer, self.delimiter, minor)
        if self.code_before_amount:
            return "{} {}".format(self.code, number)
        return "{} {}".format(number, self.code)

    def __repr__(self) -> str:
        return self.__str__()


class Url:
    """
    A URL to a website with an optional title. Will be converted into a Excel
    Hyperlink.
    """
    title: Optional[str]
    url: str

    def __init__(self, title: Optional[str], url: str):
        self.title = title
        self.url = url
