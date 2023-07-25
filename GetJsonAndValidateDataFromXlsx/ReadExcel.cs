using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Threading;

// Run in Package Manager Console
// Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
// [A]

namespace GetJsonAndValidateDataFromXlsx
{
    // Internal class ReadExcel
    internal class ReadExcel
    {
        // Private fields of the class
        private ISheet sheet;
        private Dictionary<string, Dictionary<string, string>> mapAp;
        private Dictionary<string, Dictionary<string, string>> mapNonAp;
        private CustomAction1Config Configuration;
        private ConcurrentBag<string> errors;
        private ConcurrentBag<string> debugLogs;
        private ConcurrentBag<string> criticalErrors;
        private ConcurrentBag<Dictionary<string, string>> result;
        private ConcurrentBag<Dictionary<string, string>> unique;
        private ConcurrentBag<Dictionary<string, string>> duplicates;
        private ConcurrentBag<Dictionary<string, string>> elementsNotFound;
        private ConcurrentBag<Dictionary<string, string>> elementsDuplicate;
        private ConcurrentDictionary<string, int> fileUnique;
        private int threadConut = Environment.ProcessorCount;
        private int maxLen;

        // Constructor for the ReadExcel class
        public ReadExcel(
          ISheet s,
          ref Dictionary<string, Dictionary<string, string>> mapAp,
          ref Dictionary<string, Dictionary<string, string>> mapNonAp,
          CustomAction1Config Configuration)
        {
            // Initializing class fields with the provided parameters
            this.sheet = s;
            this.mapAp = mapAp;
            this.mapNonAp = mapNonAp;
            this.Configuration = Configuration;
            this.maxLen = s.LastRowNum;
        }

        // Public method to extract data from Excel using the default number of threads
        public void ExtractData()
        {
            this.ExtractData(Environment.ProcessorCount);
        }

        // Public method to extract data from Excel using a specified number of threads
        public void ExtractData(int threads)
        {   
            // Setting the number of threads for data extraction
            this.threadConut = threads;
            // Initializing several concurrent collections to store extracted data and logs
            this.errors = new ConcurrentBag<string>();
            this.debugLogs = new ConcurrentBag<string>();
            this.criticalErrors = new ConcurrentBag<string>();
            this.result = new ConcurrentBag<Dictionary<string, string>>();
            this.elementsNotFound = new ConcurrentBag<Dictionary<string, string>>();
            this.elementsDuplicate = new ConcurrentBag<Dictionary<string, string>>();
            this.unique = new ConcurrentBag<Dictionary<string, string>>();
            this.duplicates = new ConcurrentBag<Dictionary<string, string>>();
            this.fileUnique = new ConcurrentDictionary<string, int>();

            // Logging the maximum length of the sheet
            this.debugLogs.Add(string.Format("Main thread, max length is: {0}\n", this.maxLen));

            // Finding the index of the first blank row (if any) to set the maximum length accordingly
            int blankIndex = this.GetBlankIndex(0, this.maxLen);
            if (blankIndex > 0)
                this.maxLen = blankIndex + 1;
            this.debugLogs.Add(string.Format("Main thread, max length set to: {0}\n", this.maxLen));

            // Creating a list of threads for parallel data extraction
            List<Thread> threadList = new List<Thread>();
            List<int> intList = new List<int>();
            for (int index = 0; index < this.threadConut; ++index)
            {
                threadList.Add(new Thread(new ParameterizedThreadStart(this.ParcurgereExcel)));
                intList.Add(index);
                threadList[index].Start(intList[index]);
            }

            // Waiting for all threads to complete
            for (int index = 0; index < this.threadConut; ++index)
                threadList[index].Join();
        }

        // Public method to get the extracted data as a JSON string
        public string GetFullJson() => JsonConvert.SerializeObject(this.result.ToArray());

        // Public method to get error messages as a JSON string
        public string GetErrors() => JsonConvert.SerializeObject(this.errors.ToArray());

        // Public method to get critical error messages as a JSON string
        public string GetCriticalErrors() => JsonConvert.SerializeObject(this.criticalErrors.ToArray());

        // Public method to get columns with duplicates as a JSON string
        public string GetColumnsDuplicates() => JsonConvert.SerializeObject(this.elementsDuplicate.ToArray());

        // Public method to get columns not found as a JSON string
        public string GetColumnsNotFound() => JsonConvert.SerializeObject(this.elementsNotFound.ToArray());

        // Public method to get rows with duplicates as a JSON string
        public string GetRowsDuplicates() => JsonConvert.SerializeObject(this.duplicates.ToArray());

        // Public method to get unique rows as a JSON string
        public string GetRowsUnique() => JsonConvert.SerializeObject(this.unique.ToArray());

        // Public method to get debug logs as a concatenated string
        public string GetDebugLogs()
        {
            string debugLogs = "";
            foreach (string debugLog in this.debugLogs)
                debugLogs += debugLog.ToString();
            return debugLogs;
        }

        // Private method to find the index of the first blank row between two given indices (low and high)
        private int GetBlankIndex(int low, int high)
        {
            if (high < low)
                return -1;
            int rowIndex = (high + low) / 2;
            if (this.RowIsBlank(rowIndex))
            {
                if (rowIndex <= 0)
                    return -1;
                return !this.RowIsBlank(rowIndex - 1) ? rowIndex - 1 : this.GetBlankIndex(low, rowIndex - 1);
            }
            if (rowIndex >= this.maxLen)
                return -1;
            return this.RowIsBlank(rowIndex + 1) ? rowIndex : this.GetBlankIndex(rowIndex + 1, high);
        }

        // Private method to check if a row is blank based on the configuration
        private bool RowIsBlank(int rowIndex)
        {
            IRow row = this.sheet.GetRow(rowIndex);
            if (row == null)
                return true;
            foreach (CustomAction1Config.MapColumnToField mapColumnToField in this.Configuration.ColumnsToFieldsDb)
            {
                CellReference cellReference = new CellReference(string.Format("{0}{1}", mapColumnToField.ColumnIndex, 1));
                ICell cell = row.GetCell((int)cellReference.Col);
                this.debugLogs.Add(string.Format("index: {0}, cell: {1}, value: [{2}]\n", rowIndex, mapColumnToField.ColumnIndex, cell.StringValue()));
                if (cell.StringValue() != "")
                    return false;
            }
            return true;
        }

        // Private method for each thread to process the Excel data
        private void ParcurgereExcel(object t_index)
        {
            var TranslatesDictionary = CustomTranslations.TranslatesDictionary;
            // Cast the parameter to an integer representing the thread index
            int num1 = (int)t_index;
            // Calculate the start and end rows to be processed by this thread
            int num2 = (int)((double)num1 * ((double)this.maxLen / (double)this.threadConut)) + 1;
            int num3 = (int)((double)(num1 + 1) * ((double)this.maxLen / (double)this.threadConut));
            if (num1 == 0)
                --num2;
            if (num2 < this.Configuration.XlsxStartRow)
                num2 = this.Configuration.XlsxStartRow;
            this.debugLogs.Add(string.Format("Thread #{0}, begin: {1} / end: {2}\n", num1, num2, num3));

            // Loop through the rows assigned to this thread
            for (int index = num2; index <= num3; ++index)
            {
                try
                {
                    // Get the current row from the sheet
                    IRow row = this.sheet.GetRow(index);
                    if (row != null)
                    {
                        // Create a dictionary to hold the data for the current row
                        Dictionary<string, string> dictionary = new Dictionary<string, string>();
                        dictionary.Add("_INDEX_", index.ToString());
                        string key1 = "";

                        // Loop through the columns and process each one based on the defined rules
                        foreach (CustomAction1Config.MapColumnToField mapColumnToField in this.Configuration.ColumnsToFieldsDb)
                        {
                            CellReference cellReference = new CellReference(string.Format("{0}{1}", mapColumnToField.ColumnIndex, 1));
                            ICell cell = row.GetCell((int)cellReference.Col);
                            string key2 = "";

                            // Handling numeric format and value retrieval
                            try
                            {
                                if (mapColumnToField.NumericValue)
                                {
                                    key2 = this.HandleNumericFormat(cell);
                                    if (key2 == "")
                                    {
                                        var errString = "Linia {0}, Coloana {1} -> Format numeric invalid";
                                        if (Configuration.Translation != "")
                                        {
                                            if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.NumericFormatValidation))
                                                errString = TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.NumericFormatValidation];
                                        }
                                        this.errors.Add(string.Format(errString, (index + 1), mapColumnToField.ColumnIndex));
                                    }
                                }
                                else
                                    key2 = cell.StringValue();
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Numeric format validation: Line: {0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Removing # and ; from the value, if required
                            try
                            {
                                if (mapColumnToField.RemoveSemicolonHash)
                                {
                                    key2 = key2.Replace("#", "").Replace(";", "");
                                }
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Eliminationof # and ; at: Line: {0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Splitting the value based on a specified character and taking the first part
                            try
                            {
                                if (mapColumnToField.SplitBy != "")
                                {
                                    if (key2.IndexOf(mapColumnToField.SplitBy) != -1)
                                        key2 = key2.Substring(0, key2.IndexOf(mapColumnToField.SplitBy)).Trim();
                                }
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> SplitBy at: Line: 0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Verifying the length of the value if a limit is specified
                            try
                            {
                                int limit;
                                if (mapColumnToField.TextLenLimit != "" && int.TryParse(mapColumnToField.TextLenLimit, out limit))
                                {
                                    if (key2.Length > limit && limit != 0)
                                    {
                                        var errString = "Linia {0}, Coloana {1} -> Valoare depaseste limita maxima admisa ({2})";
                                        if (Configuration.Translation != "")
                                        {
                                            if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.MaxLengthValidation))
                                                errString = TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.MaxLengthValidation];
                                        }
                                        this.errors.Add(string.Format(errString, (index + 1), mapColumnToField.ColumnIndex, limit));
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Max length validation at: Line: {0}, Column {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Verifying if the value is required and present
                            try
                            {
                                if (key2 == "" && mapColumnToField.Required)
                                {
                                    var errString = "Linia {0}, Coloana {1} -> Valoare lipsa din celula obligatorie";
                                    if (Configuration.Translation != "")
                                    {
                                        if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.RequiredColumnValidation))
                                            errString = TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.RequiredColumnValidation];
                                    }
                                    this.errors.Add(string.Format(errString, (index + 1), mapColumnToField.ColumnIndex));

                                }
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Required column validation at: Line: {0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Verifying if the value is excluded based on a dictionary
                            try
                            {
                                if (this.mapNonAp.ContainsKey(mapColumnToField.FieldDbColumn))
                                {
                                    if (this.mapNonAp[mapColumnToField.FieldDbColumn].ContainsKey(key2))
                                    {
                                        var errString = "Linia {0}, Coloana {1}, {2} -> Valoarea nu este permisa conform configuratiei";
                                        if (Configuration.Translation != "")
                                        {
                                            if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.ExcludedValuesValidation))
                                                errString = TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.ExcludedValuesValidation];
                                        }
                                        this.errors.Add(string.Format(errString, (index + 1), mapColumnToField.ColumnIndex, key2));

                                        this.elementsDuplicate.Add(new Dictionary<string, string>()
                                        {
                                          {
                                            mapColumnToField.ColumnIndex,
                                            key2
                                          }
                                        });
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Exclude values validation at: Line {0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Verifying if the value is included based on a dictionary
                            try
                            {
                                if (this.mapAp.ContainsKey(mapColumnToField.FieldDbColumn))
                                {
                                    if (this.mapAp[mapColumnToField.FieldDbColumn].ContainsKey(key2))
                                    {
                                        string str = this.mapAp[mapColumnToField.FieldDbColumn][key2];
                                        if (mapColumnToField.getBpsFormat && str != "")
                                            key2 = str + "#" + key2;
                                    }
                                    else
                                    {
                                        var errString = "Linia {0}, Coloana {1}, {2} -> Valoarea nu a fost identificata in sistem";
                                        if (Configuration.Translation != "")
                                        {
                                            if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.IncludeDictionaryValidation))
                                                errString = TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.IncludeDictionaryValidation];
                                        }
                                        this.errors.Add(string.Format(errString, (index + 1), mapColumnToField.ColumnIndex, key2));

                                        this.elementsNotFound.Add(new Dictionary<string, string>()
                                        {
                                          {
                                            mapColumnToField.ColumnIndex,
                                            key2
                                          }
                                        });
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Include dictionary validation at: Line: {0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }

                            // Adding the cell value to the current row dictionary
                            try
                            {
                                if (!dictionary.ContainsKey(mapColumnToField.FieldDbColumn))
                                    dictionary.Add(mapColumnToField.FieldDbColumn, key2);
                                if (mapColumnToField.isUniqueKey)
                                    key1 += key2;
                            }
                            catch (Exception ex)
                            {
                                this.criticalErrors.Add(string.Format("Exception -> Adding cell at current line at: Line: {0}, Column: {1} {2}", (index + 1), mapColumnToField.ColumnIndex, ex));
                            }
                        }

                        // Checking for unique keys and adding the row dictionary to the appropriate collections
                        if (this.mapAp["UNIQUE"].Count > 0)
                        {
                            if (this.mapAp["UNIQUE"].ContainsKey(key1))
                            {
                                if (this.mapAp["UNIQUE"][key1] != "")
                                    dictionary["WFD_ID"] = this.mapAp["UNIQUE"][key1];
                                this.duplicates.Add(dictionary);
                            }
                            else
                            {
                                this.unique.Add(dictionary);
                            }

                            // Checking for duplicate keys within the file
                            if (this.Configuration.CheckExcelKeyDuplicates)
                            {
                                var existing_info = fileUnique.GetOrAdd(key1, k => (index + 1));
                                if (existing_info != index + 1)
                                {
                                    var errString = "Linia {0} -> Cheia unica ({1}) este duplicata in fisierul incarcat, cheie gasita anterior la linia: {2}";
                                    if (Configuration.Translation != "")
                                    {
                                        if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.UniqueKeyValidation))
                                            errString = TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.UniqueKeyValidation];
                                    }
                                    this.errors.Add(string.Format(errString, (index + 1), key1, existing_info));
                                }
                            }
                        }

                        // Adding the row dictionary to the result collection
                        this.result.Add(dictionary);
                    }
                }
                catch (Exception ex)
                {
                    // Handling critical errors and logging them
                    int fileLineNumber = new StackTrace(ex, true).GetFrame(0).GetFileLineNumber();
                    this.criticalErrors.Add(string.Format("Exception -> Outer loop at:  Line: {0} {1}, index: {2}", (index + 1), ex, fileLineNumber));
                }
            }
        }

        // Private method to handle numeric format in a cell and convert it to a standardized format
        private string HandleNumericFormat(ICell cell)
        {
            string s = cell.StringValue();
            s.Replace(" ", "");
            if (s == "")
                s = "0";
            Decimal result;
            string string_Result;
            if (Decimal.TryParse(s, out result))
            {
                string_Result = result.ToString(CultureInfo.InvariantCulture).Replace(",", ".");
            }
            else
            {
                string_Result = "";
            }
            return string_Result;
        }
    }
}
