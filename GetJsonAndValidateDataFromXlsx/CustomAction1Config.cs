using System.Collections.Generic;
using WebCon.WorkFlow.SDK.Common;
using WebCon.WorkFlow.SDK.ConfigAttributes;

// Run in Package Manager Console
// Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
// [A]

namespace GetJsonAndValidateDataFromXlsx
{
    public class CustomAction1Config : PluginConfiguration
    {
        [ConfigEditableText(DisplayName = "Attachment ID", NullText = "First Attachment", DescriptionAsHTML = true, Description = "Uses the first attachment if Attachment ID is not provided or is 0. The action will not be executed if no attachments are found.")]
        public int? AttId { get; set; }

        [ConfigEditableInteger(DisplayName = "Excel Sheet Index", DefaultValue = 0, Description = "Integer value. By default it's 0. Refers to the index of the sheet where data is located.")]
        public int XlsxSheetIndex { get; set; }

        [ConfigEditableInteger(DisplayName = "Starting Row Index", DefaultValue = 1, Description = "Integer value, by default it's 1. The starting index of the row where data is located.")]
        public int XlsxStartRow { get; set; }

        [ConfigEditableText(DisplayName = "JSON with Unique Keys", Description = "Used when it's necessary to verify if some rows already exist in the database. This field will contain a JSON Array that includes unique keys. The unique key can be a combination of columns and can be configured in the grid below. The JArray format is: [{\"VALUE\": \"text that represents the key\", \"WFD_ID\": \"some id\"}, etc..]")]
        public string JSON_Unique_KEY { get; set; }

        [ConfigEditableBool(DisplayName = "Check for Duplicate Keys in File", Description = "In case the file contains duplicate keys, they can be found in the error output JSON.")]
        public bool CheckExcelKeyDuplicates { get; set; } = false;

        [ConfigEditableFormFieldID(DisplayName = "Output JSON - All Rows", Description = "JSON including all rows extracted from the file.")]
        public int? CompleteJsonOut { get; set; }

        [ConfigEditableFormFieldID(DisplayName = "Output JSON - Unique Rows", Description = "Outputs unique values from the Excel file. This field is populated only when 'JSON with Unique Keys' is not empty and not '[]'.")]
        public int? UniqueRowsJsonOut { get; set; }

        [ConfigEditableFormFieldID(DisplayName = "Output JSON - Duplicate Rows", Description = "Outputs duplicate values from the Excel file. This field is populated only when 'JSON with Unique Keys' is not empty and not '[]'.")]
        public int? DuplicateRowsJsonOut { get; set; }

        [ConfigEditableFormFieldID(DisplayName = "Output JSON - Errors", Description = "JSON including all errors found in the file.")]
        public int? ErrorJsonOut { get; set; }

        [ConfigEditableFormFieldID(DisplayName = "Output JSON - Invalid Column Values", Description = "Outputs values from columns that are present in the 'Exclude dictionary'. This field is populated only when the 'Exclude dictionary' is set for certain columns, indicating that these values are prohibited.")]
        public int? ExcludeColumnsJsonOut { get; set; }

        [ConfigEditableFormFieldID(DisplayName = "Output JSON - Unrecognized Column Values", Description = "Outputs values from columns that are not present in the 'Include dictionary'. This field is populated only when the 'Include dictionary' is set for certain columns, indicating that only these values are acceptable.")]
        public int? IncludeColumnsJsonOut { get; set; }

        [ConfigEditableGrid(DisplayName = "Excel Column to Database Field Mapping", Description = "Maps Excel column identifiers to WebCon fields or other names.")]
        public List<MapColumnToField> ColumnsToFieldsDb { get; set; }

        [ConfigEditableFormFieldID(DisplayName = "Logs Output Field", Description = "Field where the generated logs are output.")]
        public int? LogOutput { get; set; }

        [ConfigEditableBool(DisplayName = "Activate Debug Logs for Multithreading?", Description = "Enables the logging within the multithreading function that processes the document. Default is set to 'False'.", DefaultValue = false)]
        public bool debugLogs { get; set; } = false;

        [ConfigEditableInteger(DisplayName = "Number of Processing Threads", Description = "Defines the number of threads to be used for processing the document. The minimum value is 1, and the default value is 4.", MinValue = 1, DefaultValue = 4)]
        public int? NoThreads { get; set; }

        [ConfigEditableBool(DisplayName = "Suppress Threading Critical Errors", Description = "When set to 'True', the SDK will not stop execution due to critical errors occurring within the multithreaded processing function, allowing the process to continue. Default is set to 'False'.", DefaultValue = false)]
        public bool SupressCerrors { get; set; } = false;

        [ConfigEditableFormFieldID(DisplayName = "Output Field for Threading Critical Errors", Description = "Field where the critical errors are output when 'Suppress Threading Critical Errors' is enabled. If 'Suppress Threading Critical Errors' is not set, these errors will not be visible.")]
        public int? CriticalErrorsOutput { get; set; }

        public class MapColumnToField
        {
            [ConfigEditableGridColumBoolean(DisplayName = "Part of Unique Key?", Description = "Last action: Select the columns that constitute the unique key. This is applicable only when 'JSON with Unique Keys' is provided.")]
            public bool isUniqueKey { get; set; } = false;

            [ConfigEditableGridColumn(DisplayName = "Column Identifier", IsRequired = true, Description = "Specifies the column identifier (e.g., A, B, ...) to be mapped.")]
            public string ColumnIndex { get; set; }

            [ConfigEditableGridColumn(DisplayName = "Output JSON Parameter Name", Description = "This represents the name of the parameter in the output JSON corresponding to the column in the Excel file. For example, if 'WFD_AttText1' is mapped to column 'A' in Excel, the value of column 'A' will be stored under the parameter 'WFD_AttText1' in the output JSON. Avoid using spaces in the name. Note that 'WFD_ID' is reserved.")]
            public string FieldDbColumn { get; set; } = "";

            [ConfigEditableGridColumBoolean(DisplayName = "Is Column Numeric?", Description = "First action: Checks if the column values are numeric. If validation fails, details will be available in 'Errors JSON'.")]
            public bool NumericValue { get; set; } = false;

            [ConfigEditableGridColumBoolean(DisplayName = "Exclude # and ;", Description = "Second action: Remove '#' and ';' characters from the original string.")]
            public bool RemoveSemicolonHash { get; set; } = false;

            [ConfigEditableGridColumn(DisplayName = "Split Characters", Description = "Third action: Define the character used to split data, for example: '#', the first value from the splitting will be taken.")]
            public string SplitBy { get; set; } = "";

            [ConfigEditableGridColumn(DisplayName = "Character Length Limit", Description = "Fourth action: Specify the maximum number of characters allowed in a cell.")]
            public string TextLenLimit { get; set; } = "";

            [ConfigEditableGridColumBoolean(DisplayName = "Is Column Required?", Description = "Fifth action: Ensure this column does not contain empty values when checked.")]
            public bool Required { get; set; } = false;

            [ConfigEditableGridColumn(DisplayName = "Exclude Dictionary", Description = "Sixth action: Uused when column values must not belong to a specific dictionary. Errors will be listed in 'Errors JSON'/'Output JSON - Invalid Column Values'. Example JSON format: '[{\"VALUE\":\"ExampleValue1\"}, {\"VALUE\":\"ExampleValue2\"}, etc..]'.")]
            public string JSON_Dictionar_NonApartenenta { get; set; } = "[]";

            [ConfigEditableGridColumn(DisplayName = "Include Dictionary", Description = "Sevenh action: Used when the column values match with a specific dictionary. If a value isn't found in the provided JArray, errors will be registered in 'Errors JSON'/'Output JSON - Unrecognized Column Values'. Example JSON format: '[{\"VALUE\":\"Admis\", \"ID\": 1}, {\"VALUE\":\"Respins\"}, etc..]'. Refer to 'BPS Format' for usage of the 'ID' parameter.")]
            public string JSON_Dictionar_Apartenenta { get; set; } = "[]";

            [ConfigEditableGridColumBoolean(DisplayName = "Enable BPS Format?", Description = "Eight action: Enable this for results in BPS format. Only applicable when 'Include Dictionary' is provided and the 'ID' parameter within the dictionary's JArray is not empty. The 'ID' and 'VALUE' parameters will be used for constructing the BPS format.")]
            public bool getBpsFormat { get; set; } = false;
        }

        [ConfigEditableDropDownList(DisplayName = "Custom Translation", DataSourcePropertyName = "TranslatesList", Description = "Only the most important error messages are translated")]
        public string Translation { get; set; } = "";

        #region Dropdown data
        public List<DropDownListItem> TranslatesList
        {
            get
            {
                var list = new List<DropDownListItem>();

                foreach (var item in CustomTranslations.TranslatesDictionary.Keys)
                {
                    list.Add(new DropDownListItem { Text = item, Value = item });
                }

                return list;
            }
        }
        #endregion
    }
}