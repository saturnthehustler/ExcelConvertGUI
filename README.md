# ExcelConvertGUI
 This Python script uses Tkinter to create a GUI for Excel file conversion. Users select multiple files, validate their format, choose a target location, and then the script converts the data into a single Excel file. It provides feedback for errors and successful conversions, enhancing user experience.

## Features
- Select multiple Excel files for conversion.
- Validate selected files to ensure they are valid Excel files.
- Choose the target location and filename for the converted Excel file.
- Convert and merge selected Excel files into a single file with each file's data on a separate sheet.
- Provides error handling and informative messages for user interaction.

## Installation
1. Ensure you have Python installed on your system. You can download it from [python.org](https://www.python.org/downloads/).
2. Clone or download this repository.
3. Install the required dependencies using pip:
    ```
    pip install pandas
    pip install pandas openpyxl
    ```
4. Run the script using Python:
    ```
    python excel_converter.py
    ```

## Usage
1. Launch the script by running `excel_converter.py`.
2. Click on the "Convert Excel" button.
3. Select one or more Excel files you want to convert.
4. Choose the location and filename for the converted Excel file.
5. Wait for the conversion process to complete.
6. Once the conversion is done, you will receive a success message.

## Contributing
Contributions are welcome! If you encounter any issues or have suggestions for improvements, please open an issue or create a pull request on GitHub.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
