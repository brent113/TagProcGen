using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using OutputRowEntryDictionary = System.Collections.Generic.Dictionary<int, string>;
using OutputList = System.Collections.Generic.List<System.Collections.Generic.Dictionary<int, string>>;

namespace TagProcGen
{

    /// <summary>

/// Builds the SCADA worksheet and handles tag name formatting and merging

/// </summary>
    public class ScadaWorksheet
    {
        private Dictionary<string, string> _Pointers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
    /// Key: Pointer Name. Value: Cell Reference
    /// </summary>
        public Dictionary<string, string> Pointers
        {
            get
            {
                return _Pointers;
            }
        }

        private Excel.Worksheet _xlSheet;
        /// <summary>
    /// Excel worksheet corresponding to the SCADA template
    /// </summary>
        public Excel.Worksheet xlSheet
        {
            get
            {
                return _xlSheet;
            }
        }

        /// <summary>
    /// Create a new instance
    /// </summary>
    /// <param name="xlSheet">Excel worksheet corresponding to the SCADA template</param>
        public ScadaWorksheet(Excel.Worksheet xlSheet)
        {
            _xlSheet = xlSheet;
        }

        /// <summary>
    /// Template format to join the IED and Point names into the SCADA name.
    /// </summary>
        public string ScadaNameTemplate { get; set; }

        /// <summary>
    /// Join IED and Point names to form final SCADA name
    /// </summary>
    /// <param name="iedName">SCADA device name</param>
    /// <param name="pointName">SCADA point name</param>
    /// <returns>SCADA Name</returns>
        public string ScadaNameGenerator(string iedName, string pointName)
        {
            return ScadaNameTemplate.Replace(Keywords.IED_NAME_KEYWORD, iedName).Replace(Keywords.POINT_NAME_KEYWORD, pointName);
        }

        private Dictionary<string, ScadaTagPrototype> _ScadaTagPrototypes = new Dictionary<string, ScadaTagPrototype>();
        /// <summary>
    /// Dictionary of all SCADA tag prototypes. Key: SCADA type name, Value: Prototype
    /// </summary>
        public Dictionary<string, ScadaTagPrototype> ScadaTagPrototypes
        {
            get
            {
                return _ScadaTagPrototypes;
            }
        }

        /// <summary>
    /// SCADA tag prototype containing type-specific data.
    /// </summary>
        public class ScadaTagPrototype
        {
            /// <summary>
        /// Standard data all SCADA tags of this type have.
        /// </summary>
            public OutputRowEntryDictionary StandardColumns = new OutputRowEntryDictionary();

            /// <summary>
        /// Format to generate key from address.
        /// </summary>
            public string KeyFormat;

            /// <summary>
        /// Header row of output CSV.
        /// </summary>
            public string CsvHeader;

            /// <summary>
        /// Default data equivalent to a new blank record from DataExplorer to merge custom data into.
        /// </summary>
            public string CsvRowDefaults;

            /// <summary>
        /// Column to sort on
        /// </summary>
            public int SortingColumn;
        }

        /// <summary>
    /// Add a new SCADA prototype entry from the given data.
    /// </summary>
    /// <param name="pointTypeName">Point type name to add a prototype for.</param>
    /// <param name="defaultColumnData">Column data all SCADA points of this type have.</param>
    /// <param name="keyFormat">Format to generate key from address.</param>
    /// <param name="csvHeader">Header row of output CSV.</param>
    /// <param name="csvRowDefaults">Default values to use if column data is not specified.</param>
    /// <param name="sortingColumn">Column to sort output by.</param>
        public void AddTagPrototypeEntry(string pointTypeName, string defaultColumnData, string keyFormat, string csvHeader, string csvRowDefaults, int sortingColumn
    )
        {
            var pointTypeInfo = new PointTypeInfo(pointTypeName);
            var scadaTagPrototype = new ScadaTagPrototype();
            defaultColumnData.ParseColumnDataPairs(scadaTagPrototype.StandardColumns);

            scadaTagPrototype.KeyFormat = keyFormat;
            scadaTagPrototype.CsvHeader = csvHeader;
            scadaTagPrototype.CsvRowDefaults = csvRowDefaults;
            scadaTagPrototype.SortingColumn = sortingColumn;

            ScadaTagPrototypes.Add(pointTypeInfo.ToString(), scadaTagPrototype);
        }

        /// <summary>
    /// Rows of entries used to build the SCADA worksheet. Each row is a column dictionary.
    /// </summary>
    /// <remarks>Main output of this template</remarks>
        private Dictionary<string, OutputList> _ScadaOutputList = new Dictionary<string, OutputList>();

        private int _MaxValidatedTagLength = 0;
        /// <summary>
    /// Length of the longest tag name.
    /// </summary>
        public int MaxValidatedTagLength
        {
            get
            {
                return _MaxValidatedTagLength;
            }
        }

        private string _MaxValidatedTag;
        /// <summary>
    /// Name of the longest tag.
    /// </summary>
        public string MaxValidatedTag
        {
            get
            {
                return _MaxValidatedTag;
            }
        }

        /// <summary>
    /// Keywords that get replaced with other values.
    /// </summary>
        public class Keywords
        {
            public const string FULL_NAME_KEYWORD = "{NAME}";
            public const string IED_NAME_KEYWORD = "{DEVICENAME}";
            public const string POINT_NAME_KEYWORD = "{POINTNAME}";
            public const string ADDRESS_KEYWORD = "{ADDRESS}";
            public const string KEY_KEYWORD = "{KEY}";
            public const string RECORD_KEYWORD = "{RECORD}";
        }

        /// <summary>
    /// Throws error is tag name is invalid. Letters, numbers and space only; no symbols.
    /// </summary>
    /// <param name="tagName">Tag name to validate.</param>
        public void ValidateTagName(string tagName)
        {
            var r = Regex.Match(tagName, "^[A-Za-z0-9 ]+$", RegexOptions.None);
            if (!r.Success)
                throw new ArgumentException("Invalid tag name: " + tagName);

            if (tagName.Length > Conversions.ToInteger(Pointers[Constants.TPL_SCADA_MAX_NAME_LENGTH]))
                throw new ArgumentException("Tag name too long: " + tagName);
            if (tagName.Length > _MaxValidatedTagLength)
            {
                _MaxValidatedTagLength = tagName.Length;
                _MaxValidatedTag = tagName;
            }
        }

        /// <summary>
    /// Substitute SCADA point name and address placeholders with specified values.
    /// </summary>
    /// <param name="scadaRowEntry">SCADA entry to find and replace</param>
    /// <param name="name">Name to substitute into placeholder</param>
    /// <param name="address">Address to substitute into placeholder. Handles offset here.</param>
    /// <param name="keyFormat">Format to generate key from address with.</param>
    /// <param name="keyAddress">Optional parameter to use a different address when generating a key. Do not apply any offset.</param>
        public void ReplaceScadaKeywords(OutputRowEntryDictionary scadaRowEntry, string name, int address, string keyFormat, int keyAddress = -1)
        {
            int adjustedAddress = address + Convert.ToInt32(Pointers[Constants.TPL_SCADA_ADDRESS_OFFSET]);
            keyAddress = keyAddress > 0 ? keyAddress + Convert.ToInt32(Pointers[Constants.TPL_SCADA_ADDRESS_OFFSET]) : adjustedAddress;
            string c = string.Format(keyFormat, adjustedAddress);
            var replacements = new Dictionary<string, string>()
            {
                {
                    Keywords.FULL_NAME_KEYWORD,
                    name
                },
                {
                    Keywords.ADDRESS_KEYWORD,
                    adjustedAddress.ToString()
                },
                {
                    Keywords.KEY_KEYWORD,
                    string.Format(keyFormat, keyAddress)
                }
            };

            scadaRowEntry.ReplaceTagKeywords(replacements);
        }

        /// <summary>
    /// Add row entry to output.
    /// </summary>
    /// <param name="pointTypeInfoName">Name of the point type. </param>
    /// <param name="scadaRowEntry">Row to add to output.</param>
        public void AddScadaTagOutput(string pointTypeInfoName, OutputRowEntryDictionary scadaRowEntry)
        {
            if (!_ScadaOutputList.ContainsKey(pointTypeInfoName))
                _ScadaOutputList[pointTypeInfoName] = new OutputList();

            _ScadaOutputList[pointTypeInfoName].Add(scadaRowEntry);
        }

        /// <summary>
    /// Write all SCADA tag types to CSV.
    /// </summary>
    /// <param name="path">Source filename to append output suffix on.</param>
        public void WriteAllSCADATags(string path)
        {
            foreach (var tagGroup in _ScadaOutputList)
                WriteScadaTagCSV(tagGroup, path);
        }

        /// <summary>
    /// Write the scada worksheet out to CSV.
    /// </summary>
    /// <param name="type">Tag type to write out.</param>
    /// <param name="path">Source filename to append output suffix on.</param>
        private void WriteScadaTagCSV(KeyValuePair<string, OutputList> type, string path)
        {
            string typeName = type.Key;
            var tagGroup = type.Value;

            var comparer = new BySortingColumn(ScadaTagPrototypes[typeName].SortingColumn);
            tagGroup.Sort(comparer);

            if (!ScadaTagPrototypes.ContainsKey(typeName))
                throw new ArgumentException("Unable to locate tag prototype.\r\n\r\nMissing: \"" + typeName + "\" in tag prototype.");

            string csvPath = System.IO.Path.GetDirectoryName(path) + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileNameWithoutExtension(path) + "_ScadaTags_" + typeName + ".csv";
            using (var csvStreamWriter = new System.IO.StreamWriter(csvPath, false))
            {
                var csvWriter = new CsvHelper.CsvWriter(csvStreamWriter);

                // Write header
                ScadaTagPrototypes[typeName].CsvHeader.Split(',').ToList().ForEach(x => csvWriter.WriteField(x));
                csvWriter.NextRecord();

                // Parse default columns and types to substitute data into
                var newRow = ScadaTagPrototypes[typeName].CsvRowDefaults.Split(',').ToList().Select(s =>
                {
                    var r = new { Value = s, isString = !int.TryParse(s, out int i) };
                    if (r.isString)
                        r.Value = r.Value.Replace(Conversions.ToString('"'), "");
                    return r;
                }
    ).ToList();

                // Write out to CSV
                int record = 1;
                foreach (var c in tagGroup)
                {
                    for (int i = 1, loopTo = newRow.Count(); i <= loopTo; i++)
                    {
                        if (c.ContainsKey(i) && !string.IsNullOrWhiteSpace(c[i]))
                        {
                            if ((c[i] ?? "") == Keywords.RECORD_KEYWORD)
                                c[i] = Conversions.ToString(record);
                            csvWriter.WriteField(c[i], newRow[i - 1].isString);
                        }
                        else
                            csvWriter.WriteField(newRow[i - 1].Value, newRow[i - 1].isString);
                    }
                    csvWriter.NextRecord();
                    record += 1;
                }
            }
        }
    }
}
