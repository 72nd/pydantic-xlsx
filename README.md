# pydantic-xlsx

 <p align="center">
  <img width="140" src="misc/logo.png">
</p>

This Python library tries to ease parsing and dumping data from and to Excel's xlsx (also known as [Office Open XML Workbook](https://en.wikipedia.org/wiki/Office_Open_XML)) with the help of [pydantic](https://pydantic-docs.helpmanual.io/) models. It uses [openpyxl](https://openpyxl.readthedocs.io/) to interact with the Excel files. As with pydantic you define the structure of your in- or output xlsx file as a pydantic model. 

As with pydantic you define the structure of your data as Models with the help of Python's typing system.


## Motivation

First of all: If there is another way to accomplish your goal without using spreadsheet software or data formats _do it._ Spreadsheets have many drawbacks in contrasts to »real« databases and programming. Consider using [Jupyter](https://jupyter.org/) if you need some sort of interaction with your data.

This package was written for circumstances where you're forced to work with spreadsheet files. The goal of pydantic-xlsx is to make the interaction with such data sources/environments as comfortable as possible while enforcing as much data validation as possible. Another more or less legit use case for this library is the ability to get a quick overview over your data for debugging.


## Overview

To get a first glimpse consider the following example:

```python
from enum import Enum
from typing import List

from pydantic_xlsx import XlsxField, XlsxModel
from pydantic_xlsx.money import Euro


class Function(str, Enum):
	boss = "Boss"
	worker = "Worker"


class Employee(XlsxModel):
	name: str = XlsxField(alias="Full Name")
	age: int
	wage: Euro
	function: Function


class Company(XlsxModel):
	staff: List[Employee]


my_company = Company(staff=[
	Employee(name="Morio Rossi", age=42, wage=4200, function=Function.boss)
	Employee(name="John Doe", age=23, wage=2300, function=Function.worker)
])

my_company.to_file("my-company.xlsx")
```


## Mapping/Composition

Unlike data formates supported by pydantic (like [JSON](https://en.wikipedia.org/wiki/JSON) or [YAML](https://en.wikipedia.org/wiki/YAML)) spreadsheets do not have an straight forward way of arbitrary nesting of data structures. This quickly become a problem if your model describes some lists of lists or alike. Undressed this will lead to a wide range of undefined behavior when translating a pydantic model to a spreadsheets. To circumvent this pydantic-xlsx only allows a defined range of nested data structures to be a valid input. This section gives an overview about this types and their mapping into spreadsheets (this process is internally known as »composition«).


### Single Model

This handles all models which do _not contain any models as property type._ The resulting spreadsheet will contain one sheet with a single row of data.

```python 
class Employee(XlsxModel):
	name: str = XlsxField(alias="Full Name")
	age: int

employee = Employee(name="Morio Rossi", age=42)
employee.to_file("employee.xlsx")
```

Will result in the following file:

TODO: Image


### Root collection


### Collection
 


## Types



## Field options


## Document options

The library tries to output spreadsheets with some reasonable styles and tweaks. By defining the inner `Config` class in your model, you can control this behavior and the appearance of the output document. For more information you can consult the documentation on the `XlsxConfig` class (TODO: insert link).


### Header font (`header_font`)

The library expects the first row of every sheet to contain the names of field. Use this option to alter the appearance of this row by defining your own [openpyxl Font](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fonts.html) (learn more about styling with openpyxl [here](https://openpyxl.readthedocs.io/en/stable/styles.html)). The field defaults to `openpyxl.styles.Fonts(name="Arial", bold=true)`.


### Font (`font`)

Optional [openpyxl Font](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fonts.html) (learn more about styling with openpyxl [here](https://openpyxl.readthedocs.io/en/stable/styles.html)) for all of your spreadsheet except the title row. Defaults to `None`.


### Freeze Cells (`freeze_cell`)

Freezing cells makes them visible while scrolling trough your document (learn more in the [Excel documentation](https://support.microsoft.com/en-us/office/freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f)). This is especially to pin the header row. This is also what pydantic-xlsx is doing by default (freeze cell at `A2`) 

TODO: The rest
