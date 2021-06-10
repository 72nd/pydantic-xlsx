"""
This module defines some frequently used currencies for convince.
"""

from .types import Money


class Euro(Money):
    """Definition for Euro."""
    minor_unit = 2
    code = "€"
    code_before_amount = False
    delimiter = ","
    thousands_separator = "."


class SwissFranc(Money):
    """Defines Swiss franc."""
    minor_unit = 2
    code = "CHF"
    code_before_amount = False
    delimiter = "."
    thousands_separator = " "


class UnitedStatesDollar(Money):
    """Defines United States dollar."""
    minor_unit = 2
    code = "$"
    code_before_amount = False
    delimiter = "."
    thousands_separator = ","

