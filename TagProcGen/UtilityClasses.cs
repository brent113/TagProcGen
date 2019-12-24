using System.Linq;
using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using OutputRowEntryDictionary = System.Collections.Generic.Dictionary<int, string>;

namespace TagProcGen
{

    /// <summary>
    /// Global constants.
    /// </summary>
    public class Constants
    {
        // Initial Pointer
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
        public const string TPL_DEF = "A3";

        // Important Excel worksheet names
        public const string TPL_SHEET_PREFIX = "TPL_";
        public const string TPL_DEF_SHEET = "TPL_Def";
        public const string TPL_RTAC_SHEET = "TPL_RTAC";
        public const string TPL_SCADA_SHEET = "TPL_SCADA";

        // Global Template Reference Constants
        public const string TPL_RTAC_DEF = "TPL_RTAC_DEF";
        public const string TPL_SCADA_DEF = "TPL_SCADA_DEF";
        public const string TPL_IED_DEF = "TPL_IED_DEF";

        // Custom Map Constants
        public const string TPL_TAG_MAP = "TPL_TAG_MAP";
        public const string TPL_RTAC_TAGS = "TPL_RTAC_TAGS";

        // RTAC Template Reference Constants
        public const string TPL_RTAC_TAG_PROTO = "TPL_RTAC_TAG_PROTO";
        public const string TPL_RTAC_MAP_NAME = "TPL_RTAC_MAP_NAME";
        public const string TPL_RTAC_TAG_MAP = "TPL_RTAC_TAG_MAP";
        public const string TPL_RTAC_ALIAS_SUB = "TPL_RTAC_ALIAS_SUB";
        public const string TPL_RTAC_TAG_PROC_COLS = "TPL_RTAC_TAG_PROC_COLS";
        public const string TPL_RTAC_TAG_PROC_WRAP_MODE = "TPL_RTAC_TAG_PROC_WRAP_MODE";

        // SCADA Template Reference Constants
        public const string TPL_SCADA_NAME_FORMAT = "TPL_SCADA_NAME_FORMAT";
        public const string TPL_SCADA_MAX_NAME_LENGTH = "TPL_SCADA_MAX_NAME_LENGTH";
        public const string TPL_SCADA_TAG_PROTO = "TPL_SCADA_TAG_PROTO";
        public const string TPL_SCADA_ADDRESS_OFFSET = "TPL_SCADA_ADDRESS_OFFSET";

        // IED Template Reference Constants
        public const string TPL_DATA = "TPL_DATA";
        public const string TPL_IED_NAMES = "TPL_IED_NAMES";
        public const string TPL_OFFSETS = "TPL_OFFSETS";

        // Point Type Constants
        public const string STATUS_BINARY = "STATUSBINARY";
        public const string STATUS_ANALOG = "STATUSANALOG";
        public const string CONTROL_BINARY = "CONTROLBINARY";
        public const string CONTROL_ANALOG = "CONTROLANALOG";
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
    }

    /// <summary>
    /// Represents points that are analog or binary, both status and control.
    /// </summary>
    public class PointTypeInfo
    {
        private readonly string _pointTypeName;

        /// <summary>Returns True if the point is a status type.</summary>
        public bool IsStatus { get; }
        /// <summary>Returns True if the point is a control type.</summary>
        public bool IsControl => !IsStatus;
        /// <summary>Returns True if the point is a binary type.</summary>
        public bool IsBinary { get; }
        /// <summary>Returns True if the point is an analog type.</summary>
        public bool IsAnalog => !IsBinary;

        /// <summary>
        /// Initialize a new instance of PointTypeInfo from text.
        /// </summary>
        /// <param name="pointTypeText">Text to parse point type information for.</param>
        /// <remarks>Text should be like: StatusAnalog, or ControlBinary</remarks>
        public PointTypeInfo(string pointTypeText)
        {
            _pointTypeName = pointTypeText.ToUpper();

            var AllTypes = new[] {
                Constants.STATUS_BINARY,
                Constants.STATUS_ANALOG,
                Constants.CONTROL_BINARY,
                Constants.CONTROL_ANALOG
            };
            var StatusTypes = new[] { Constants.STATUS_BINARY, Constants.STATUS_ANALOG };
            var BinaryTypes = new[] { Constants.STATUS_BINARY, Constants.CONTROL_BINARY };

            if (!AllTypes.Contains(_pointTypeName))
                throw new Exception(string.Format("Point type {0} is not a valid point type.", _pointTypeName));

            IsStatus = StatusTypes.Contains(_pointTypeName);
            IsBinary = BinaryTypes.Contains(_pointTypeName);
        }

        /// <summary>
        /// Initialize a new instance of PointTypeInfo from values.
        /// </summary>
        /// <param name="isStatus">Indicates the point is a status point.</param>
        /// <param name="isBinary">Indicates the point is a binary point.</param>
        public PointTypeInfo(bool isStatus, bool isBinary)
        {
            IsStatus = isStatus;
            IsBinary = isBinary;
        }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>The point type name as a string.</returns>
        public override string ToString()
        {
            return _pointTypeName;
        }
    }

    /// <summary>
    /// Custom comparer that sorts a list of output rows by the given sorting column alphanumerically.
    /// </summary>
    public class BySortingColumn : IComparer<OutputRowEntryDictionary>
    {
        private readonly int _sortingColumn;

        /// <summary>
        /// Initialize a new instance of BySortingColumn.
        /// </summary>
        /// <param name="sortingColumn">Column number to sort by.</param>
        public BySortingColumn(int sortingColumn) => _sortingColumn = sortingColumn;


        /// <summary>
        /// Compare two values.
        /// </summary>
        /// <param name="x">Value1 to compare.</param>
        /// <param name="y">Value2 to compare.</param>
        /// <returns>X less than Y: Less than 0. X=Y: 0. X greater than Y: Greater than 0.</returns>
        public int Compare(OutputRowEntryDictionary x, OutputRowEntryDictionary y)
        {
            double xVal = Convert.ToDouble(x[_sortingColumn]);
            double YVal = Convert.ToDouble(y[_sortingColumn]);

            return xVal.CompareTo(YVal);
        }
    }

    /// <summary>
    /// Utilities that are used by various function throughout the program.
    /// </summary>
    public static class SharedUtils
    {
        /// <summary>
        /// Read 2 columns of data into a Key: Value structure. Optional list of parameters to verify were successfully read in.
        /// </summary>
        /// <param name="start">Excel range to begin reading data pairs at.</param>
        /// <param name="dict">Dictionary to store data pairs in.</param>
        /// <param name="ExpectedParameters">List of parameters that must be in the dictionary or an error will be thrown.</param>
        public static void ReadPairRange(Excel.Range start, Dictionary<string, string> dict, params string[] ExpectedParameters)
        {
            while (!string.IsNullOrEmpty((string)start.Value))
            {
                dict[(string)start.Value] = (string)start.get_Offset(0, 1).Text;
                start = start.get_Offset(1, 0);
            }

            foreach (var p in ExpectedParameters)
            {
                if (!dict.ContainsKey(p))
                    throw new Exception("Unable to locate pointer.\r\n\r\nMissing: \"" + p + "\" from " + start.Parent.Name);
            }
        }

        /// <summary>
        /// Convert an output row dictionary into a sparse string array where the 1-based output row column indices are converted to a 0-based string array index.
        /// </summary>
        /// <param name="outputRow">Row data to convert into a string array.</param>
        /// <returns>Sparsely populated string array.</returns>
        public static string[] OutputRowEntryDictionaryToArray(OutputRowEntryDictionary outputRow)
        {
            // Create string array (0 based) from max column index (1 based)
            int arrayUBound = outputRow.Max(kv => kv.Key) - 1;
            var s = new string[arrayUBound + 1];

            // Store column values in string array with 1 base to 0 base index conversion
            foreach (var c in outputRow)
                s[c.Key - 1] = c.Value;

            return s;
        }
    }

    /// <summary>
    /// Contains extension methods
    /// </summary>
    public static class ExtensionMethods
    {
        /// <summary>
        /// Parse a string containing formatted column / data pairs into the output row dictionary format.
        /// </summary>
        /// <param name="columnDataPairString">String to parse. Ex: [2, {NAME}];[3, {ADDRESS}];[5, {ALIAS}]</param>
        /// <param name="columnDataDict">Output dictionary to store parsed data in.</param>
        public static void ParseColumnDataPairs(this string columnDataPairString, OutputRowEntryDictionary columnDataDict)
        {
            if (columnDataPairString.Length == 0)
                return;
            // Split col / data pairs - example format: [1, True];[2, {NAME}]
            foreach (var colPair in columnDataPairString.Split(';'))
            {
                if (colPair.Length == 0)
                    throw new Exception("Malformed column / data pair: " + columnDataPairString);
                // strip [ and ]
                if (colPair[0] != '[' | colPair[colPair.Length - 1] != ']')
                    throw new Exception("Malformed column / data pair: " + colPair);
                var t = colPair.Substring(1, colPair.Length - 2).Split(',');

                if (!int.TryParse(t[0].Trim(), out int colIndex))
                    throw new Exception("Invalid Column Index: unable to convert \"" + t[0].Trim() + "\" to an integer");

                string colData = t[1].Trim();

                columnDataDict.Add(colIndex, colData);
            }
        }

        /// <summary>
        /// Apply replacements to column keywords like {NAME} and {ADDRESS}
        /// </summary>
        /// <param name="columns">Column data pair dictionary to update</param>
        /// <param name="replacements">Dictionary of keywords (like {NAME}) and their replacement</param>
        public static void ReplaceTagKeywords(this OutputRowEntryDictionary columns, Dictionary<string, string> replacements)
        {
            var keys = new List<int>(columns.Keys.ToList());
            foreach (var columnKey in keys)
            {
                foreach (var rep in replacements)
                    columns[columnKey] = columns[columnKey].Replace(rep.Key, rep.Value);
            }
        }

        /// <summary>
        /// Search a string for a character and return the Nth character.
        /// </summary>
        /// <param name="s">String to search.</param>
        /// <param name="t">Character to search for.</param>
        /// <param name="n">Instance number of the character.</param>
        /// <returns>Returns the Nth index as an Integer.</returns>
        public static int GetNthIndex(this string s, char t, int n)
        {
            int count = 0;
            for (int i = 0, loopTo = s.Length - 1; i <= loopTo; i++)
            {
                if (s[i] == t)
                {
                    count += 1;
                    if (count == n)
                        return i;
                }
            }
            return -1;
        }
    }
}
