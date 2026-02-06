# Elevate Helper

The Elevate Helper is a software application that assists with the management and processing of Elevate files, a specialized file format used in the building industry. This tool provides a user-friendly interface to perform various tasks, such as creating copies of Elevate files, modifying building types, and generating reports.

## Installation

To use the Elevate Helper, follow these steps:

1. Ensure you have Python 3.12 or higher installed on your system.
2. Install the required dependencies by running the following command in your terminal or command prompt:

   ```
   pip install -r requirements.txt
   ```

3. Run the `app.py` file to launch the Elevate Helper application.

## Usage

The Elevate Helper provides the following functionality:

1. **Building Type Selection**: Choose the building type (Office, Residence, or Hotel) for the Elevate file you want to process.
2. **File Path**: Specify the path to the Elevate file you want to work with.
3. **Run**: Execute the Elevate file processing based on the selected building type and morning/full-day options.
4. **Print Report**: Generate a report based on the processed Elevate file.

## API

The Elevate Helper utilizes the following Python libraries:

- `tkinter`: For creating the graphical user interface.
- `body`: Handles the main functionality of the application, such as creating file copies, modifying building types, and generating reports.
- `app_icon`: Provides the application icon.

## Contributing

If you would like to contribute to the Elevate Helper project, please follow these guidelines:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Implement your changes and ensure they work as expected.
4. Submit a pull request with a detailed description of your changes.

## License

The Elevate Helper is released under the [MIT License](LICENSE).

## Testing

To ensure the Elevate Helper functions correctly, the following test cases have been implemented:

1. **Building Type Selection**: Verify that the application correctly handles the selection of different building types (Office, Residence, and Hotel).
2. **File Path**: Ensure the application can correctly process the provided Elevate file path.
3. **Run**: Test the execution of the Elevate file processing for each building type, including the morning-only and full-day options.
4. **Print Report**: Validate the generation of the report based on the processed Elevate file.

These test cases can be executed manually or integrated into an automated testing framework.