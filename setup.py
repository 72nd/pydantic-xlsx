import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pydantic-xlsx",
    version="0.1.0",
    author="72nd",
    author_email="msg@frg72.com",
    description="Parsing and dumping from and to Excel's xlsx files using pydantic Models.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/72nd/pydantic-xlsx",
    license="MIT",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    packages=["pydantic_xlsx"],
    install_requires=[
        "openpyxl==3.0.7",
        "pydantic==1.8.2",
    ]
)
