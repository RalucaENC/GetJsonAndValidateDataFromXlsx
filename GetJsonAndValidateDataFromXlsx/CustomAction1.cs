using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using WebCon.WorkFlow.SDK.ActionPlugins;
using WebCon.WorkFlow.SDK.ActionPlugins.Model;
using WebCon.WorkFlow.SDK.Documents.Model;
using WebCon.WorkFlow.SDK.Documents.Model.Attachments;


// Run in Package Manager Console
// Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
// [A]

namespace GetJsonAndValidateDataFromXlsx
{
    public class CustomAction1 : CustomAction<CustomAction1Config>
    {
        private long start;
        private long end;
        private string log = "";

        public override void Run(RunCustomActionParams args)
        {
            CurrentDocumentData currentDocument = args.Context.CurrentDocument;
            var TranslatesDictionary = CustomTranslations.TranslatesDictionary;

            start = Stopwatch.GetTimestamp();
            string str = this.validateMap(Configuration);

            if (!(str == "") && !(str == "[]"))
            {
                args.HasErrors = true;
                args.Message += str;
            }
            else if (args.Context.CurrentDocument.Attachments.Count == 0)
            {
                args.HasErrors = true;
                if (Configuration.Translation != "")
                {
                    if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.AttachmentNotFound))
                        args.Message += TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.AttachmentNotFound];
                }
                else
                    args.Message += "Nu a fost gasit nici un atasament!";

            }
            else
            {

                AttachmentData attachmentData;
                if (Configuration.AttId.HasValue && Configuration.AttId.Value != 0)
                    attachmentData = currentDocument.Attachments.GetByID(Configuration.AttId.Value);
                else
                    attachmentData = (currentDocument.Attachments)[0];

                if (!attachmentData.FileExtension.ToLower().StartsWith(".xls"))
                {
                    args.HasErrors = true;
                    if (Configuration.Translation != "")
                    {
                        if (TranslatesDictionary[Configuration.Translation].ContainsKey(CustomTranslations.TranslationType.WrongAttachmentExtension))
                            args.Message += TranslatesDictionary[Configuration.Translation][CustomTranslations.TranslationType.WrongAttachmentExtension];
                    }
                    else
                        args.Message += "Extensia fisierului trebuie sa fie .xls sau .xlsx!\n";
                }
                else
                {
                    end = Stopwatch.GetTimestamp();
                    log += string.Format("File validation time: {0}\n", (end - start) / 10000.0);

                    start = Stopwatch.GetTimestamp();
                    ISheet sheetAt = WorkbookFactory.Create(new MemoryStream(attachmentData.Content)).GetSheetAt(Configuration.XlsxSheetIndex);
                    Dictionary<string, Dictionary<string, string>> mapAp = new Dictionary<string, Dictionary<string, string>>();
                    Dictionary<string, Dictionary<string, string>> mapNonAp = new Dictionary<string, Dictionary<string, string>>();
                    end = Stopwatch.GetTimestamp();
                    log += string.Format("Create workbook time: {0}\n", (end - start) / 10000.0);

                    start = Stopwatch.GetTimestamp();
                    getMapData(ref mapAp, ref mapNonAp, Configuration);
                    end = Stopwatch.GetTimestamp();
                    log += string.Format("Inlude/Exclude/Unique dictionary processing time: {0}\n", (end - start) / 10000.0);

                    start = Stopwatch.GetTimestamp();
                    ReadExcel readExcel1 = new ReadExcel(sheetAt, ref mapAp, ref mapNonAp, Configuration);
                    if (Configuration.NoThreads.HasValue && Configuration.NoThreads.Value != 0)
                    {
                        ReadExcel readExcel2 = readExcel1;
                        int threads = Configuration.NoThreads.Value;
                        readExcel2.ExtractData(threads);
                    }
                    else
                        readExcel1.ExtractData();

                    end = Stopwatch.GetTimestamp();
                    log += string.Format("Extracting data time: {0}\n", (end - start) / 10000.0);


                    start = Stopwatch.GetTimestamp();
                    if (!(readExcel1.GetCriticalErrors() == "") && !(readExcel1.GetCriticalErrors() == "[]"))
                    {
                        if (Configuration.SupressCerrors)
                        {
                            if (Configuration.CriticalErrorsOutput.HasValue)
                            {
                                string criticalErrors = readExcel1.GetCriticalErrors();
                                args.Context.CurrentDocument.SetFieldValue(Configuration.CriticalErrorsOutput.Value, criticalErrors);
                            }
                        }
                        else
                        {
                            args.HasErrors = true;
                            RunCustomActionParams customActionParams = args;
                            args.Message += "Critical Error: " + readExcel1.GetCriticalErrors();
                            return;
                        }
                    }

                    if (Configuration.CompleteJsonOut.HasValue)
                    {
                        string fullJson = readExcel1.GetFullJson();
                        currentDocument.SetFieldValue(Configuration.CompleteJsonOut.Value, fullJson);
                    }

                    if (Configuration.ErrorJsonOut.HasValue)
                    {
                        string errors = readExcel1.GetErrors();
                        currentDocument.SetFieldValue(Configuration.ErrorJsonOut.Value, errors);
                    }

                    if (Configuration.UniqueRowsJsonOut.HasValue)
                    {
                        string rowsUnique = readExcel1.GetRowsUnique();
                        currentDocument.SetFieldValue(Configuration.UniqueRowsJsonOut.Value, rowsUnique);
                    }

                    if (Configuration.DuplicateRowsJsonOut.HasValue)
                    {
                        string rowsDuplicates = readExcel1.GetRowsDuplicates();
                        currentDocument.SetFieldValue(Configuration.DuplicateRowsJsonOut.Value, rowsDuplicates);
                    }

                    if (Configuration.ExcludeColumnsJsonOut.HasValue)
                    {
                        string columnsDuplicates = readExcel1.GetColumnsDuplicates();
                        currentDocument.SetFieldValue(Configuration.ExcludeColumnsJsonOut.Value, columnsDuplicates);
                    }

                    if (Configuration.IncludeColumnsJsonOut.HasValue)
                    {
                        string columnsNotFound = readExcel1.GetColumnsNotFound();
                        currentDocument.SetFieldValue(Configuration.IncludeColumnsJsonOut.Value, columnsNotFound);
                    }

                    end = Stopwatch.GetTimestamp();
                    log += string.Format("Set field output values: {0}\n", (end - start) / 10000.0);

                    if (Configuration.debugLogs)
                        log = log + "\n" + readExcel1.GetDebugLogs();

                    if (Configuration.LogOutput.HasValue)
                    {
                        currentDocument.SetFieldValue(Configuration.LogOutput.Value, log);
                    }
                    
                }
            }
        }

        public string validateMap(CustomAction1Config Confioguration)
        {
            List<string> stringList = new List<string>();
            Dictionary<string, bool> dictionary = new Dictionary<string, bool>();
            List<CustomAction1Config.MapColumnToField> columnsToFieldsDb = Configuration.ColumnsToFieldsDb;
            for (int index = 0; index < columnsToFieldsDb.Count; ++index)
            {
                if (dictionary.ContainsKey(columnsToFieldsDb[index].FieldDbColumn))
                {
                    stringList.Add("Duplicate mapping, column " + columnsToFieldsDb[index].ColumnIndex + ", value: " + columnsToFieldsDb[index].FieldDbColumn);
                    dictionary.Add(columnsToFieldsDb[index].FieldDbColumn, true);
                }
            }
            return JsonConvert.SerializeObject(stringList);
        }

        private static string getMapData(
          ref Dictionary<string, Dictionary<string, string>> mapAp,
          ref Dictionary<string, Dictionary<string, string>> mapNonAp,
          CustomAction1Config Configuration)
        {
            string mapData = "";
            foreach (CustomAction1Config.MapColumnToField mapColumnToField in Configuration.ColumnsToFieldsDb)
            {
                if (mapColumnToField.JSON_Dictionar_Apartenenta != "[]" && mapColumnToField.JSON_Dictionar_Apartenenta != "")
                {
                    try
                    {
                        JArray jarray = JArray.Parse(mapColumnToField.JSON_Dictionar_Apartenenta);
                        Dictionary<string, string> dictionary = new Dictionary<string, string>();
                        foreach (JObject jobject in jarray)
                        {
                            JToken jtoken1 = jobject["VALUE"];
                            if (jtoken1 != null)
                            {
                                JToken jtoken2 = jobject["ID"];
                                if (jtoken2 != null)
                                    dictionary[jtoken1.ToString()] = jtoken2.ToString();
                                else
                                    dictionary[jtoken1.ToString()] = "";
                            }
                        }
                        mapAp.Add(mapColumnToField.FieldDbColumn, dictionary);
                    }
                    catch (Exception ex)
                    {
                        mapData = mapData + "JSON parsing error: include dictionary" + mapColumnToField.ColumnIndex + " " + mapColumnToField.FieldDbColumn + "\n";
                        mapData += ex.ToString();
                        mapData += "\n";
                    }
                }
                if (mapColumnToField.JSON_Dictionar_NonApartenenta != "[]" && mapColumnToField.JSON_Dictionar_NonApartenenta != "")
                {
                    try
                    {
                        JArray jarray = JArray.Parse(mapColumnToField.JSON_Dictionar_NonApartenenta);
                        Dictionary<string, string> dictionary = new Dictionary<string, string>();
                        foreach (JObject jobject in jarray)
                        {
                            JToken jtoken3 = jobject["VALUE"];
                            if (jtoken3 != null)
                            {
                                JToken jtoken4 = jobject["ID"];
                                if (jtoken4 != null)
                                    dictionary[jtoken3.ToString()] = jtoken4.ToString();
                                else
                                    dictionary[jtoken3.ToString()] = "";
                            }
                        }
                        mapNonAp.Add(mapColumnToField.FieldDbColumn, dictionary);
                    }
                    catch (Exception ex)
                    {
                        mapData = mapData + "JSON parsing error: exclude dictionary" + mapColumnToField.ColumnIndex + " " + mapColumnToField.FieldDbColumn + "\n";
                        mapData += ex.ToString();
                        mapData += "\n";
                    }
                }
            }
            try
            {
                if (Configuration.JSON_Unique_KEY != "[]" && Configuration.JSON_Unique_KEY != "")
                {
                    JArray jarray = JArray.Parse(Configuration.JSON_Unique_KEY);
                    Dictionary<string, string> dictionary = new Dictionary<string, string>();
                    foreach (JObject jobject in jarray)
                    {
                        if (jobject.ContainsKey("VALUE"))
                        {
                            if (jobject.ContainsKey("WFD_ID"))
                                dictionary[jobject["VALUE"].ToString()] = jobject["WFD_ID"].ToString();
                            else
                                dictionary[jobject["VALUE"].ToString()] = "";
                        }
                    }
                    mapAp.Add("UNIQUE", dictionary);
                }
                else
                    mapAp.Add("UNIQUE", new Dictionary<string, string>());
            }
            catch (Exception ex)
            {
                mapData = mapData + "JSON parsing error: Unique KEY dictionary\n" + ex.ToString() + "\n";
            }
            return mapData;
        }
    }
}
