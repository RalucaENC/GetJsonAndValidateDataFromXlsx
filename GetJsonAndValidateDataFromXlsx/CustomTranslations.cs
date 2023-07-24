using System;
using System.Collections.Generic;
using System.Text;

namespace GetJsonAndValidateDataFromXlsx
{
    internal class CustomTranslations
    {
        public enum TranslationType
        {
            AttachmentNotFound, // free text for attachment not found error
            WrongAttachmentExtension, // free text for wrong attachment error
            NumericFormatValidation, // string must containt {0} for line number and {1} for column letter
            MaxLengthValidation, // string must containt {0} for line number, {1} for column letter and {2} for max length limit
            RequiredColumnValidation, // string must containt {0} for line number and {1} for column letter
            ExcludedValuesValidation, // string must containt {0} for line number, {1} for column letter and {2} for the actual value
            IncludeDictionaryValidation, // string must containt {0} for line number, {1} for column letter and {2} for the actual value
            UniqueKeyValidation // string must containt {0} for line number, {1} for unique key value and {2} for the line number
        }

        //do not use empty strings for language acronym, reserved value
        public static Dictionary<string, Dictionary<TranslationType, string>> TranslatesDictionary = new Dictionary<string, Dictionary<TranslationType, string>>
        {
            // Continue with data for English
            {
                "en", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Attachment not found!" },
                    { TranslationType.WrongAttachmentExtension, "Wrong file extension! The plugin works with .xlsx and .xls files only!" },
                    { TranslationType.NumericFormatValidation, "Line {0}, Column {1} -> Invalid numeric format" },
                    { TranslationType.MaxLengthValidation, "Line {0}, Column {1} -> Value exceeds the maximum allowed length ({2})" },
                    { TranslationType.RequiredColumnValidation, "Line {0}, Column {1} -> Missing value in required cell" },
                    { TranslationType.ExcludedValuesValidation, "Line {0}, Column {1}, {2} -> The value is not allowed according to the configuration" },
                    { TranslationType.IncludeDictionaryValidation, "Line {0}, Column {1}, {2} -> The value was not found in the system" },
                    { TranslationType.UniqueKeyValidation, "Line {0} -> The unique key ({1}) is duplicated in the uploaded file, previously found at line: {2}" }
                }
            },

            // Continue with data for French
            {
                "fr", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Pièce jointe introuvable!" },
                    { TranslationType.WrongAttachmentExtension, "Mauvaise extension de fichier! Le plugin fonctionne uniquement avec les fichiers .xlsx et .xls!" },
                    { TranslationType.NumericFormatValidation, "Ligne {0}, Colonne {1} -> Format numérique invalide" },
                    { TranslationType.MaxLengthValidation, "Ligne {0}, Colonne {1} -> La valeur dépasse la longueur maximale autorisée ({2})" },
                    { TranslationType.RequiredColumnValidation, "Ligne {0}, Colonne {1} -> Valeur manquante dans la cellule requise" },
                    { TranslationType.ExcludedValuesValidation, "Ligne {0}, Colonne {1}, {2} -> La valeur n'est pas autorisée selon la configuration" },
                    { TranslationType.IncludeDictionaryValidation, "Ligne {0}, Colonne {1}, {2} -> La valeur n'a pas été trouvée dans le système" },
                    { TranslationType.UniqueKeyValidation, "Ligne {0} -> La clé unique ({1}) est dupliquée dans le fichier téléchargé, précédemment trouvée à la ligne: {2}" }
                }
            },

            // Continue with data for Spanish
            {
                "es", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "¡Archivo adjunto no encontrado!" },
                    { TranslationType.WrongAttachmentExtension, "¡Extensión de archivo incorrecta! ¡El complemento solo funciona con archivos .xlsx y .xls!" },
                    { TranslationType.NumericFormatValidation, "Línea {0}, Columna {1} -> Formato numérico inválido" },
                    { TranslationType.MaxLengthValidation, "Línea {0}, Columna {1} -> El valor excede la longitud máxima permitida ({2})" },
                    { TranslationType.RequiredColumnValidation, "Línea {0}, Columna {1} -> Falta valor en celda requerida" },
                    { TranslationType.ExcludedValuesValidation, "Línea {0}, Columna {1}, {2} -> El valor no está permitido según la configuración" },
                    { TranslationType.IncludeDictionaryValidation, "Línea {0}, Columna {1}, {2} -> El valor no se encontró en el sistema" },
                    { TranslationType.UniqueKeyValidation, "Línea {0} -> La clave única ({1}) está duplicada en el archivo cargado, se encontró previamente en la línea: {2}" }
                }
            },

            // Continue with data for German
            {
                "de", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Anhang nicht gefunden!" },
                    { TranslationType.WrongAttachmentExtension, "Falsche Dateierweiterung! Das Plugin arbeitet nur mit .xlsx und .xls Dateien!" },
                    { TranslationType.NumericFormatValidation, "Zeile {0}, Spalte {1} -> Ungültiges numerisches Format" },
                    { TranslationType.MaxLengthValidation, "Zeile {0}, Spalte {1} -> Wert überschreitet die maximale erlaubte Länge ({2})" },
                    { TranslationType.RequiredColumnValidation, "Zeile {0}, Spalte {1} -> Fehlender Wert in erforderlicher Zelle" },
                    { TranslationType.ExcludedValuesValidation, "Zeile {0}, Spalte {1}, {2} -> Der Wert ist laut Konfiguration nicht zulässig" },
                    { TranslationType.IncludeDictionaryValidation, "Zeile {0}, Spalte {1}, {2} -> Der Wert wurde im System nicht gefunden" },
                    { TranslationType.UniqueKeyValidation, "Zeile {0} -> Der eindeutige Schlüssel ({1}) ist in der hochgeladenen Datei doppelt vorhanden, zuvor gefunden in Zeile: {2}" }
                }
            },

            // Continue with data for Simplified Chinese
            {
                "zh", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "找不到附件！" },
                    { TranslationType.WrongAttachmentExtension, "文件扩展名错误！插件只支持.xlsx和.xls文件！" },
                    { TranslationType.NumericFormatValidation, "行{0}，列{1} -> 数字格式无效" },
                    { TranslationType.MaxLengthValidation, "行{0}，列{1} -> 值超过最大允许长度 ({2})" },
                    { TranslationType.RequiredColumnValidation, "行{0}，列{1} -> 必填单元格中缺少值" },
                    { TranslationType.ExcludedValuesValidation, "行{0}，列{1}，{2} -> 根据配置，该值不被允许" },
                    { TranslationType.IncludeDictionaryValidation, "行{0}，列{1}，{2} -> 系统中找不到该值" },
                    { TranslationType.UniqueKeyValidation, "行{0} -> 上传的文件中重复的唯一键 ({1})，先前在行 {2} 中找到" }
                }
            },

            // Continue with data for Japanese
            {
                "ja", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "添付ファイルが見つかりません！" },
                    { TranslationType.WrongAttachmentExtension, "ファイル拡張子が間違っています！プラグインは.xlsxおよび.xlsファイルのみと互換性があります！" },
                    { TranslationType.NumericFormatValidation, "行{0}、列{1} -> 数値形式が無効です" },
                    { TranslationType.MaxLengthValidation, "行{0}、列{1} -> 値が最大許容長 ({2}) を超えています" },
                    { TranslationType.RequiredColumnValidation, "行{0}、列{1} -> 必要なセルに値がありません" },
                    { TranslationType.ExcludedValuesValidation, "行{0}、列{1}、{2} -> 値は設定により許可されていません" },
                    { TranslationType.IncludeDictionaryValidation, "行{0}、列{1}、{2} -> 値がシステムで見つかりませんでした" },
                    { TranslationType.UniqueKeyValidation, "行{0} -> 一意のキー ({1}) がアップロードしたファイル内で重複しています、以前は行 {2} で見つかりました" }
                }
            },

            // Continue with data for Russian
            {
                "ru", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Вложение не найдено!" },
                    { TranslationType.WrongAttachmentExtension, "Неправильное расширение файла! Плагин работает только с файлами .xlsx и .xls!" },
                    { TranslationType.NumericFormatValidation, "Строка {0}, Колонка {1} -> Неверный числовой формат" },
                    { TranslationType.MaxLengthValidation, "Строка {0}, Колонка {1} -> Значение превышает максимально допустимую длину ({2})" },
                    { TranslationType.RequiredColumnValidation, "Строка {0}, Колонка {1} -> Пропущено значение в обязательной ячейке" },
                    { TranslationType.ExcludedValuesValidation, "Строка {0}, Колонка {1}, {2} -> Значение не допускается согласно настройкам" },
                    { TranslationType.IncludeDictionaryValidation, "Строка {0}, Колонка {1}, {2} -> Значение не найдено в системе" },
                    { TranslationType.UniqueKeyValidation, "Строка {0} -> Уникальный ключ ({1}) повторяется в загружаемом файле, ранее найдено в строке: {2}" }
                }
            },

            // Continue with data for Portuguese
            {
                "pt", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Anexo não encontrado!" },
                    { TranslationType.WrongAttachmentExtension, "Extensão de arquivo errada! O plugin funciona apenas com arquivos .xlsx e .xls!" },
                    { TranslationType.NumericFormatValidation, "Linha {0}, Coluna {1} -> Formato numérico inválido" },
                    { TranslationType.MaxLengthValidation, "Linha {0}, Coluna {1} -> O valor excede o comprimento máximo permitido ({2})" },
                    { TranslationType.RequiredColumnValidation, "Linha {0}, Coluna {1} -> Valor ausente na célula necessária" },
                    { TranslationType.ExcludedValuesValidation, "Linha {0}, Coluna {1}, {2} -> O valor não é permitido de acordo com a configuração" },
                    { TranslationType.IncludeDictionaryValidation, "Linha {0}, Coluna {1}, {2} -> O valor não foi encontrado no sistema" },
                    { TranslationType.UniqueKeyValidation, "Linha {0} -> A chave única ({1}) está duplicada no arquivo carregado, encontrada anteriormente na linha: {2}" }
                }
            },

            // Continue with data for Italian
            {
                "it", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Allegato non trovato!" },
                    { TranslationType.WrongAttachmentExtension, "Estensione del file errata! Il plugin funziona solo con file .xlsx e .xls!" },
                    { TranslationType.NumericFormatValidation, "Riga {0}, Colonna {1} -> Formato numerico non valido" },
                    { TranslationType.MaxLengthValidation, "Riga {0}, Colonna {1} -> Il valore supera la lunghezza massima consentita ({2})" },
                    { TranslationType.RequiredColumnValidation, "Riga {0}, Colonna {1} -> Valore mancante nella cella richiesta" },
                    { TranslationType.ExcludedValuesValidation, "Riga {0}, Colonna {1}, {2} -> Il valore non è consentito in base alla configurazione" },
                    { TranslationType.IncludeDictionaryValidation, "Riga {0}, Colonna {1}, {2} -> Il valore non è stato trovato nel sistema" },
                    { TranslationType.UniqueKeyValidation, "Riga {0} -> La chiave univoca ({1}) è duplicata nel file caricato, precedentemente trovata alla riga: {2}" }
                }
            },

            // Continue with data for Dutch
            {
                "nl", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "Bijlage niet gevonden!" },
                    { TranslationType.WrongAttachmentExtension, "Verkeerde bestandsextensie! De plug-in werkt alleen met .xlsx en .xls bestanden!" },
                    { TranslationType.NumericFormatValidation, "Regel {0}, Kolom {1} -> Ongeldig numeriek formaat" },
                    { TranslationType.MaxLengthValidation, "Regel {0}, Kolom {1} -> Waarde overschrijdt de maximaal toegestane lengte ({2})" },
                    { TranslationType.RequiredColumnValidation, "Regel {0}, Kolom {1} -> Ontbrekende waarde in vereiste cel" },
                    { TranslationType.ExcludedValuesValidation, "Regel {0}, Kolom {1}, {2} -> De waarde is niet toegestaan volgens de configuratie" },
                    { TranslationType.IncludeDictionaryValidation, "Regel {0}, Kolom {1}, {2} -> De waarde werd niet gevonden in het systeem" },
                    { TranslationType.UniqueKeyValidation, "Regel {0} -> De unieke sleutel ({1}) is gedupliceerd in het geüploade bestand, eerder gevonden op regel: {2}" }
                }
            },

            // Continue with data for Hindi
            {
                "hi", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "अटैचमेंट नहीं मिला!" },
                    { TranslationType.WrongAttachmentExtension, "गलत फ़ाइल एक्सटेंशन! प्लगइन केवल .xlsx और .xls फ़ाइलों के साथ काम करता है!" },
                    { TranslationType.NumericFormatValidation, "लाइन {0}, कॉलम {1} -> अमान्य संख्यात्मक प्रारूप" },
                    { TranslationType.MaxLengthValidation, "लाइन {0}, कॉलम {1} -> मूल्य अधिकतम अनुमति दी गई लंबाई ({2}) से अधिक है" },
                    { TranslationType.RequiredColumnValidation, "लाइन {0}, कॉलम {1} -> आवश्यक सेल में मूल्य गुम है" },
                    { TranslationType.ExcludedValuesValidation, "लाइन {0}, कॉलम {1}, {2} -> मूल्य विन्यास के अनुसार अनुमति नहीं है" },
                    { TranslationType.IncludeDictionaryValidation, "लाइन {0}, कॉलम {1}, {2} -> मूल्य प्रणाली में नहीं मिला" },
                    { TranslationType.UniqueKeyValidation, "लाइन {0} -> अद्वितीय कुंजी ({1}) अपलोड की गई फ़ाइल में डुप्लिकेट है, पहले लाइन पर पाई गई: {2}" }
                }
            },

            // Continue with data for Arabic
            {
                "ar", new Dictionary<TranslationType, string>
                {
                    { TranslationType.AttachmentNotFound, "المرفق غير موجود!" },
                    { TranslationType.WrongAttachmentExtension, "امتداد الملف خاطئ! يعمل الإضافة فقط مع ملفات .xlsx و .xls!" },
                    { TranslationType.NumericFormatValidation, "السطر {0}, العمود {1} -> تنسيق رقمي غير صالح" },
                    { TranslationType.MaxLengthValidation, "السطر {0}, العمود {1} -> القيمة تتجاوز الطول الأقصى المسموح به ({2})" },
                    { TranslationType.RequiredColumnValidation, "السطر {0}, العمود {1} -> القيمة مفقودة في الخلية المطلوبة" },
                    { TranslationType.ExcludedValuesValidation, "السطر {0}, العمود {1}, {2} -> القيمة غير مسموح بها وفقًا للتكوين" },
                    { TranslationType.IncludeDictionaryValidation, "السطر {0}, العمود {1}, {2} -> لم يتم العثور على القيمة في النظام" },
                    { TranslationType.UniqueKeyValidation, "السطر {0} -> المفتاح الفريد ({1}) مكرر في الملف المحمّل، وجد سابقا في السطر: {2}" }
                }
            }

        };
    }
}
