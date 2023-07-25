# GIT README

This repository contains several files related to the functionality of a custom action plugin in WebCon BPS. Each file serves a specific purpose and contributes to the overall functionality of the plugin.

## Excel to JSON Plugin: Functionality Breakdown

Our Excel to JSON Plugin is a quite useful tool designed to streamline and optimize the process of transforming data from Excel spreadsheets into the universally-compatible JSON format. The plugin is implemented with a variety of functionalities to ensure a seamless Extraction, Transformation, and Loading (ETL) process. Below is a comprehensive, technical overview of its features:

### 1.File Selection: This module provides an interface to choose the target Excel file for data extraction by providing the Attachment ID. It allows the user to specify the sheet and row number to begin from. If not explicitly provided, the process starts from the first cell of the first sheet, indexing from zero.

### 2.Unique Data Identification: Leveraging hash-based or sorting techniques, this feature identifies unique data elements within a high volume of data, thereby optimizing the storage and preventing redundancy in the database.

### 3.Data Validation: This function serves as a data integrity checker, designed to maintain the consistency and accuracy of the data. It detects any duplicate identifiers and flags any values that deviate from the predefined rules. The validation rules can be customized according to user requirements.

### 4.Data Transformation and Output: Post-extraction and validation, the data is transformed from Excel format into the JSON format using a serialization process. JSON, being a lightweight data-interchange format, is easy for humans to read and write and easy for machines to parse and generate. Post-transformation, the plugin provides various output formats, such as a comprehensive JSON containing all data rows, a JSON consisting of unique rows, one with duplicate rows, and another with the errors encountered during the process.

### 5.Field Mapping: This functionality links Excel columns with specific properties in the JSON output. The mapping follows a user-defined schema that correlates a particular column in Excel with a respective property in the JSON structure. This ensures accurate data placement and structured JSON output.

### 6.Processing Settings: The plugin employs a multi-threaded processing strategy for handling large Excel files. This concurrent approach allows for simultaneous processing of various segments of the file, effectively optimizing the processing speed and resource usage.

### 7.Error Management: This module is designed to handle critical errors that might occur during the Excel data processing. Users can configure the plugin to either halt the entire operation upon encountering an error or continue the processing while logging the error for post-processing review. This flexible error management ensures robust operation and minimizes data loss or corruption.

## Files Description

### 1. CustomAction1.cs

This file contains the main logic and implementation of the custom action plugin. It includes the following functionalities:

- Retrieving the current document / attachment from the WebCon instance
- Validating the mapping configuration and generating error messages if necessary
- Checking if the required attachment is present and has the correct file extension
- Reading the data from the Excel file and processing it based on the configuration settings
- Extracting data from the Excel file, validating it, and storing it in the appropriate output fields in the document
- Logging debug information and errors when necessary

This file also includes a helper method to validate the mapping configuration and a method to get the mapping data from dictionaries.

### 2. CustomAction1Config.cs

This file contains the configuration model for the custom action plugin. It includes properties that represent the configuration settings for the plugin. These settings include:

- Attachment ID: Specifies the ID of the attachment to be used for processing. If no ID is provided, the first attachment is used.
- Excel Sheet Index: Specifies the index of the sheet in the Excel file where the data is located.
- Starting Row Index: Specifies the starting index of the row where the data is located.
- JSON with Unique Keys: Specifies the JSON array that includes unique keys for verifying if some rows already exist in the database.
- Check for Duplicate Keys in File: Indicates whether to check for duplicate keys within the Excel file.
- Output JSON - All Rows: Specifies the output field for the JSON including all rows extracted from the file.
- Output JSON - Unique Rows: Specifies the output field for the JSON including unique values from the Excel file.
- Output JSON - Duplicate Rows: Specifies the output field for the JSON including duplicate values from the Excel file.
- Output JSON - Errors: Specifies the output field for the JSON including all errors found in the file.
- Output JSON - Invalid Column Values: Specifies the output field for the JSON including values from columns that are present in the "Exclude dictionary".
- Output JSON - Unrecognized Column Values: Specifies the output field for the JSON including values from columns that are not present in the "Include dictionary".
- Excel Column to Database Field Mapping: Specifies the mapping of Excel column identifiers to WebCon fields or other names.
- Logs Output Field: Specifies the field where the generated logs are output.
- Activate Debug Logs for Multithreading: Indicates whether to enable logging within the multithreading function for processing the document.
- Number of Processing Threads: Specifies the number of threads to be used for processing the document.
- Suppress Threading Critical Errors: Indicates whether to suppress critical errors that occur within the multithreaded processing function.
- Output Field for Threading Critical Errors: Specifies the field where critical errors are output when suppressing threading critical errors.
- Custom Translation: Specifies the custom translation for error messages.

This file also includes a nested class for mapping columns to fields and a property for populating the custom translation dropdown list.

### 3. CustomTranslations.cs

This file contains a class that defines custom translations for error messages used in the custom action plugin.
It includes a dictionary that maps translation types to translated error messages for multiple languages.

### 4. ExcelExtensions.cs

This file contains an extension method for the NPOI library that helps retrieve the string value of a cell in Excel.
It includes a method called `StringValue` that returns the string value of a cell regardless of its data type.

### 5. ReadExcel.cs

This file contains a class that handles the logic for reading and processing data from an Excel file.
It includes methods for extracting data, handling error messages, and generating output in JSON format.
It also includes methods for validating numeric formats, removing characters from cells, splitting cell values, and checking for blank rows.
The class supports multithreading for faster data processing.

## Usage

To use this custom action plugin in WebCon BPS, follow these steps:

1. Compile and Publish the project and generate the plugin archive (.zip) file.
2. Import the archive file into the WebCon Workflow Designer Studio.
3. Add the custom action to your workflow.
4. Configure the custom action by providing the necessary settings and mappings.

For a deeper understanding and more comprehensive details of each functionality, we encourage you to explore the dedicated description of each parameter within the plugin. This will provide you with valuable insights into the operation and customization options, enabling you to maximize the utility of the Excel to JSON Plugin for your specific use-cases.

   ![image](https://github.com/RalucaENC/GetJsonAndValidateDataFromXlsx/assets/140409176/dc169551-7cc7-4ad0-858a-905947c564a2)

