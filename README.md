# Python Code to Word Document Converter

This Python class enables conversion of Python code into a Word document. It allows you to input Python code as a string and save it as a Word document with a specified file name.

## Attributes

- `code` (str): The Python code to be converted.
- `file_name` (str): The name of the Word document file.
- `bold` (bool): Default is False. Set it to True if bold lettering is required.
- `font_name` (str): Default is "Consolas".
- `font_size` (int): Default is 9.
- `table_style` (str): Default is "Light Shading Accent 1".
- `validate` (bool): Default is True. If set to True, it checks the syntax of the input code. Throws an exception if any syntax errors are found.
- `color_codes` (dict): Default is None. Indicates usage of default color codes; provide color codes for specific tokens in dictionary format for custom color formats.

## Methods

- `__init__(self, code: str, file_name: str, bold: bool = False, font_name: str = "Consolas", font_size: int = 9, table_style: str = "Light Shading Accent 1", validate: bool = True, color_codes: dict = None)`: Initializes the CodeToDocx with the provided arguments.
- `generate_docx(self)`: Converts the provided Python code to a Word document and saves it with the specified file name.

## Example Usage

```python
converter = CodeToDocx(
    code=r"print('Hello, World!')",
    file_name="python_code.docx"
)

converter.generate_docx()