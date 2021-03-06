# pydantic-xlsx

 <p align="center">
  <img width="100" src="./misc/logo.png">
</p>

This Python library tries to ease parsing and dumping data from and to Excel's xlsx (also known as [Office Open XML Workbook](https://en.wikipedia.org/wiki/Office_Open_XML)) with the help of [pydantic](https://pydantic-docs.helpmanual.io/) models. It uses [openpyxl](https://openpyxl.readthedocs.io/) to interact with the Excel files. As with pydantic you define the structure of your in- or output xlsx file as a pydantic model. 

As with pydantic you define the structure of your data as Models with the help of Python's typing system.

You can find the API documentation [here](https://72nd.github.io/pydantic-xlsx/pydantic_xlsx.html).

## State of the project

Alpha state. This package started as a module within another application. I'm currently extracting pydantic-xlsx from this project. So expect some rough edges and missing documentation.


## Motivation and Overview

First of all: If there is another way to accomplish your goal without using spreadsheet software or data formats _do it._ Spreadsheets have many drawbacks in contrasts to »real« databases and programming. Consider using [Jupyter](https://jupyter.org/) if you need some sort of interaction with your data.

This package was written for circumstances where you're forced to work with spreadsheet files. The goal of pydantic-xlsx is to make the interaction with such data sources/environments as comfortable as possible while enforcing as much data validation as possible. Another more or less legit use case for this library is the ability to get a quick overview over your data for debugging.


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
	name: str = XlsxField("", alias="Full Name")
	age: int
	wage: Euro
	function: Function

    class Config:
        use_enum_values = True
        allow_population_by_field_name = True


class Company(XlsxModel):
	staff: List[Employee]

	class Config:
		freeze_cell = "A2"


my_company = Company(staff=[
	Employee(name="Morio Rossi", age=42, wage=4200, function=Function.boss),
	Employee(name="John Doe", age=23, wage=2300, function=Function.worker)
])

my_company.to_file("my-company.xlsx")
```

Results in the following file:

![Resulting Xlsx File](misc/example.png)

You can parse the file using the `from_file` method.

```python
loaded_company = Company.from_file("my-company.xlsx")
print(loaded_company)
```

_A word on the Config sub-class:_ Inside the Employee model the `Config` sub-class sets two fairly common pydantic options when working with Excel files. `use_enum_values` represents enums as their values and not as a classpath without this option »Boss« would be represented as `function.boss` and »Worker« as `function.worker`. Using the enum value makes the spreadsheet more user-friendly. By setting `allow_population_by_field_name` to `True` you can define alternative column names by setting the `alias` property of a field.


## Features

- In-/Exporting structured data from/to Excel while benefiting from Pydantic's mature features like data validation.
- The correct Excel number-format is applied according to the field's data type. It's also possible to customize the formatting for a specific field.
- Define document wide fonts as well as alter the font for columns.
- Enums columns provides the spreadsheet user with a drop-down menu showing all allowed values.
- The data is referenced as a [Worksheet Table](https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c) in the Xlsx document. Providing more information on the structure of the data and fast sorting per column.
- Internal money field type which generates the correct Excel number-format. Some currencies (like Euro or US Dollar) are already defined others can be easily created.
- The format for printing can be controlled withing pydantic_xlsx.


## Mapping/Composition

Unlike data formates supported by pydantic (like [JSON](https://en.wikipedia.org/wiki/JSON) or [YAML](https://en.wikipedia.org/wiki/YAML)) spreadsheets do not have an straight forward way of arbitrary nesting data structures. This quickly become a problem if your model describes some lists of lists or alike. Undressed this will lead to a wide range of undefined behavior when translating a pydantic model to a spreadsheets. To circumvent this pydantic-xlsx only allows a defined range of nested data structures to be a valid input. This section gives an overview about this types and their mapping into spreadsheets (this process is internally known as »composition«).


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

![Single Model mapping](misc/mapping-01.svg)

As you see the Excel sheet is named after your Model Class (`Employee`) which contains a single row of data. Single Model Mapping can only represents a single entry of data thus is not very helpful in most cases.


### Root collection

![Root collection mapping](misc/mapping-02.svg)


### Collection
 
![Collection mapping](misc/mapping-03.svg)


## Types

_Todo._


## Field options

You can alter the appearance and behavior of columns by using `XlsxField`. The available options are defined in the [`XlsxFieldInfo` Class](https://72nd.github.io/pydantic-xlsx/pydantic_xlsx/fields.html#XlsxFieldInfo).


### Font (`font`)

Alter the font of a specific field. The fonts are defined using the [openpyxl Font](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fonts.html) object (see above for an example).


### Number Format (`number_format`)

Optional Excel number format code to alter the display of numbers, dates and so on. Pleas refer to the [Excel documentation](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68) to learn more about the syntax.



## Document options

The library tries to output spreadsheets with some reasonable styles and tweaks. By defining the inner `Config` class in your model, you can control this behavior and the appearance of the output document. For more information you can consult the documentation on the [`XlsxConfig` class](https://72nd.github.io/pydantic-xlsx/pydantic_xlsx/config.html#XlsxConfig).


### Header font (`header_font`)

The library expects the first row of every sheet to contain the names of field. Use this option to alter the appearance of this row by defining your own [openpyxl Font](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fonts.html) (learn more about styling with openpyxl [here](https://openpyxl.readthedocs.io/en/stable/styles.html)). The field defaults to `openpyxl.styles.Font(name="Arial", bold=True)`.


### Font (`font`)

Optional [openpyxl Font](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fonts.html) (learn more about styling with openpyxl [here](https://openpyxl.readthedocs.io/en/stable/styles.html)) for all of your spreadsheet except the title row. Defaults to `None`.


### Freeze Cells (`freeze_cell`)

Freezing cells makes them visible while scrolling trough your document (learn more in the [Excel documentation](https://support.microsoft.com/en-us/office/freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f)). This is especially to pin the header row. This is also what pydantic-xlsx is doing by default (freeze cell at `A2`) 

_Todo: The rest_


## Known pitfalls

### Massive amount of empty cells when loading a Spreadsheet with data validation (Generics)

_Cause:_ pydantic-xlsx uses a method of openpyxl to determine the dimension of the data area (aka the part of the spreadsheet actually containing some data). A cell is treated as non-empty (thus expanding the size of the imported area) from the point some properties where set for this cell. Defining a valid data range is on of them. If a user accidentally define a valid data range for the whole column you end up with pydanic-xlsx trying to import and validate thousands of seemingly empty rows.

_Solution._ This is exactly why you shouldn't use spreadsheets in the first place. The only solution is to manually delete all formatting etc from _all_ worksheets. Or just copy the relevant data into a new spreadsheet (including the header).
