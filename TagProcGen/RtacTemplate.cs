using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using OutputRowEntryDictionary = System.Collections.Generic.Dictionary<int, string>;
using OutputList = System.Collections.Generic.List<System.Collections.Generic.Dictionary<int, string>>;

namespace TagProcGen
{

    /// <summary>
    /// Generates RTAC tags. Stores tag prototypes, handles server tag generation.
    /// </summary>
    public class RtacTemplate
    {
        /// <summary>
        /// Create a new instance
        /// </summary>
        /// <param name="templateName">RTAC template sheet name</param>
        public RtacTemplate(string templateName) => TemplateName = templateName;

        /// <summary>
        /// Key: Pointer Name. Value: Cell Reference
        /// </summary>
        public Dictionary<string, string> Pointers { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// RTAC template worksheet name
        /// </summary>
        public string TemplateName { get; }

        /// <summary>
        /// Name of the SCADA server object in the RTAC.
        /// </summary>
        public string RtacServerName { get; set; }

        /// <summary>
        /// Server tag alias template.
        /// </summary>
        public string AliasNameTemplate { get; set; }

        /// <summary>
        /// Dictionary of server tag prototypes. Key: Server type name, Value: Prototype Root Type
        /// </summary>
        public Dictionary<string, ServerTagRootPrototype> RtacTagPrototypes { get; } = new Dictionary<string, ServerTagRootPrototype>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Starting value of next IED's tags. Incremented by offsets
        /// </summary>
        public Dictionary<string, int> TagTypeRunningAddressOffset { get; } = new Dictionary<string, int>();

        /// <remarks>Contains 1:1 map to prototype. Key: Device tag type, Value: Server tag map info.</remarks>
        private readonly Dictionary<string, ServerTagMapInfo> _IedToServerTypeMap = new Dictionary<string, ServerTagMapInfo>();

        /// <summary>
        /// Add a new entry into the device-server tag map.
        /// </summary>
        /// <param name="iedTypeName">Device type name.</param>
        /// <param name="serverTypeName">Server type name.</param>
        /// <param name="performQualityWrapping">Indicates wheter to substitute nominal data with the source tag quality is bad.</param>
        public void AddIedServerTagMap(string iedTypeName, string serverTypeName, bool performQualityWrapping)
        {
            var tagMapInfo = new ServerTagMapInfo() { ServerTagTypeName = serverTypeName, PerformQualityWrapping = performQualityWrapping };
            _IedToServerTypeMap[iedTypeName] = tagMapInfo;
        }

        /// <summary>
        /// Get server tag map information for a given device type name.
        /// </summary>
        /// <param name="iedTypeName">Device type name.</param>
        /// <returns>Server tag map information, or nothing if no entry exists.</returns>
        public ServerTagMapInfo GetServerTypeByIedType(string iedTypeName)
        {
            if (!_IedToServerTypeMap.ContainsKey(iedTypeName))
                return null;
            return _IedToServerTypeMap[iedTypeName];
        }

        private readonly Dictionary<string, string> _TagAliasSubstitutes = new Dictionary<string, string>();
        /// <summary>
        /// Placeholders to search for and replace with associated value.
        /// </summary>
        /// <remarks>Key: Find, Value: Replace</remarks>
        public Dictionary<string, string> TagAliasSubstitutes
        {
            get
            {
                return _TagAliasSubstitutes;
            }
        }

        /// <summary>
        /// List of all server tags by type. Generated from device templates.
        /// </summary>
        private readonly Dictionary<string, OutputList> RtacOutputList = new Dictionary<string, OutputList>();

        /// <summary>
        /// Keywords that get replaced with other values.
        /// </summary>
        private class Keywords
        {
            /// <summary>Name</summary>
            public const string NameKeyword = "{NAME}";
            /// <summary>Address</summary>
            public const string AddressKeyword = "{ADDRESS}";
            /// <summary>Alias</summary>
            public const string AliasKeyword = "{ALIAS}";
            /// <summary>Control</summary>
            public const string ControlKeyword = "{CTRL}";
        }

        /// <summary>
        /// Load a new server tag prototype entry. Creates new prototype or adds information to existing array prototype.
        /// </summary>
        /// <param name="tagInfo">Tag name and index information.</param>
        /// <param name="nameTemplate">Formatting template for generated tags.</param>
        /// <param name="defaultColumnData">Default data all tags have.</param>
        /// <param name="sortingColumn">Column to sort alphanumerically on before writing out. Only needs to be specified once per prototype.</param>
        /// <param name="pointTypeText">Type of the point, either binary / analog and status / control.</param>
        /// <param name="nominalColumns">String denoting the presence of nominal columns in a format like "23" or "12:25".</param>
        public void AddTagPrototypeEntry(ServerTagInfo tagInfo, string nameTemplate, string defaultColumnData,
                                         int sortingColumn, string pointTypeText, string nominalColumns)
        {
            tagInfo.ThrowIfNull(nameof(tagInfo));
            nameTemplate.ThrowIfNull(nameof(nameTemplate));
            defaultColumnData.ThrowIfNull(nameof(defaultColumnData));
            pointTypeText.ThrowIfNull(nameof(pointTypeText));
            nominalColumns.ThrowIfNull(nameof(nominalColumns));


            // Get existing tag, or create new
            ServerTagRootPrototype tagRootPrototype;
            if (RtacTagPrototypes.ContainsKey(tagInfo.RootServerTagTypeName))
                tagRootPrototype = RtacTagPrototypes[tagInfo.RootServerTagTypeName];
            else
            {
                tagRootPrototype = new ServerTagRootPrototype() { SortingColumn = -1 };
                RtacTagPrototypes.Add(tagInfo.RootServerTagTypeName, tagRootPrototype);
            }

            // Create entry in TagGenerationAddressBase
            TagTypeRunningAddressOffset[tagInfo.RootServerTagTypeName] = 0;

            // Add sorting column to root prototype if it is valid
            if (sortingColumn > -1)
                tagRootPrototype.SortingColumn = sortingColumn;

            // Add data direction to root prototype if it is valid
            if (pointTypeText.Length > 0)
                tagRootPrototype.PointType = new PointTypeInfo(pointTypeText);

            // Parse nominal column information
            if (nominalColumns.Length > 0)
            {
                var colonSplit = nominalColumns.Split(new[] { '.', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);

                if (colonSplit.Length == 1)
                    tagRootPrototype.NominalColumns = (Convert.ToInt32(colonSplit[0]), Convert.ToInt32(colonSplit[0]));
                else if (colonSplit.Length == 2)
                {
                    tagRootPrototype.NominalColumns = (Convert.ToInt32(colonSplit[0]), Convert.ToInt32(colonSplit[1]));

                    // Check for even number of columns - analog limits come in pairs. 1:10 = 10-1, should be odd
                    if ((tagRootPrototype.NominalColumns.upper - tagRootPrototype.NominalColumns.lower) % 2 == 0)
                        throw new TagGenerationException($"Tag prototype {tagInfo.RootServerTagTypeName} has an odd number of nominal value columns. Only even number of columns allowed.");
                }
                else
                    throw new TagGenerationException("Invalid analog limit column range. Expecting format like '10 or [11..20]'");
            }

            // Ensure the array has a placeholder for the incoming index
            for (int i = tagRootPrototype.TagPrototypeEntries.Count, loopTo = tagInfo.Index; i <= loopTo; i++)
                tagRootPrototype.TagPrototypeEntries.Add(null);// Add placeholders

            // Store prototype entry
            var newTagPrototypeEntry = new ServerTagPrototypeEntry();
            {
                newTagPrototypeEntry.ServerTagNameTemplate = nameTemplate;
                defaultColumnData.ParseColumnDataPairs(newTagPrototypeEntry.StandardColumns);
            }

            // Store new prototype entry
            tagRootPrototype.TagPrototypeEntries[tagInfo.Index] = newTagPrototypeEntry;
        }

        /// <summary>
        /// Ensure all loaded tag prototypes have a valid sorting column, point information,
        /// and status points have a nominal indication column.
        /// </summary>
        public void ValidateTagPrototypes()
        {
            foreach (var ta in RtacTagPrototypes)
            {
                if (ta.Value.SortingColumn < 0)
                    throw new TagGenerationException(string.Format("Tag prototype {0} is missing a valid sorting column.", ta.Key));
                if (ta.Value.PointType == null)
                    throw new TagGenerationException(string.Format("Tag prototype {0} is missing a valid data direction.", ta.Key));
                if (ta.Value.PointType.IsStatus && (ta.Value.NominalColumns.lower == -1 || ta.Value.NominalColumns.upper == -1))
                    throw new TagGenerationException(string.Format("Tag prototype {0} is a status type but is missing valid nominal columns.", ta.Key));
            }
        }

        /// <summary>
        /// Returns the information of a valid tag, otherwise throws an exception.
        /// </summary>
        /// <param name="iedTagName">Device tag name to validate.</param>
        /// <returns>TagInfo structure of valid tag.</returns>
        public ServerTagInfo ValidateTag(string iedTagName)
        {
            var tagMapInfo = GetServerTypeByIedType(iedTagName);
            if (tagMapInfo == null)
                throw new TagGenerationException("Unable to locate tag mapping.\r\n\r\nMissing: \"" + iedTagName + "\" in tag map.");

            var Tag = new ServerTagInfo(tagMapInfo.ServerTagTypeName);
            if (!RtacTagPrototypes.ContainsKey(Tag.RootServerTagTypeName))
                throw new TagGenerationException("Unable to locate tag prototype.\r\n\r\nMissing: \"" + Tag.RootServerTagTypeName + "\" in tag prototype.");

            return Tag;
        }

        /// <summary>
        /// Return the server tag associated with th given device tag.
        /// </summary>
        /// <param name="iedTagType">Name of device tag to get server info for.</param>
        /// <returns>Server tag information.</returns>
        /// <remarks>Ex: operSPC-T -> DNPC, in TagInfo container</remarks>
        public ServerTagInfo GetServerTagInfoByDevice(string iedTagType)
        {
            return ValidateTag(iedTagType);
        }

        /// <summary>
        /// Return the server tag associated with th given device tag.
        /// </summary>
        /// <param name="iedTagType">Name of device tag to get server prototype for.</param>
        /// <returns>Server tag root prototype.</returns>
        /// <remarks>Ex: operSPC-T -> DNPC</remarks>
        public ServerTagRootPrototype GetServerTagPrototypeByDevice(string iedTagType)
        {
            var Tag = GetServerTagInfoByDevice(iedTagType);

            return RtacTagPrototypes[Tag.RootServerTagTypeName];
        }

        /// <summary>
        /// Returns the specific server prototype entry associated with th given device tag.
        /// </summary>
        /// <param name="iedTagType">Name of device tag to get server prototype entry for.</param>
        /// <returns>Server tag prototype entry.</returns>
        /// <remarks>Ex: operSPC-T -> DNPC[2]</remarks>
        public ServerTagPrototypeEntry GetServerTagEntryByDevice(string iedTagType)
        {
            var Tag = GetServerTagInfoByDevice(iedTagType);
            var TagInfo = GetServerTagInfoByDevice(iedTagType);

            return RtacTagPrototypes[Tag.RootServerTagTypeName].TagPrototypeEntries[TagInfo.Index];
        }

        /// <summary>
        /// Returns the text after the 2nd dot in an array tag. Returns empty string if type is not an array type.
        /// </summary>
        /// <param name="tagInfo">Tag information to get array suffix for.</param>
        /// <returns>String containing the characters after the 2nd dot in a tag format. Ex: Result of input {SERVER}.BO_{0:D5}.operLatchOn is operLatchOn.</returns>
        public string GetArraySuffix(ServerTagInfo tagInfo)
        {
            tagInfo.ThrowIfNull(nameof(tagInfo));

            string format = RtacTagPrototypes[tagInfo.RootServerTagTypeName].TagPrototypeEntries[tagInfo.Index].ServerTagNameTemplate;

            int secondDotIndex = format.GetNthIndex('.', 2);
            if (secondDotIndex < 1)
                return "";

            return format.Substring(secondDotIndex, format.Length - secondDotIndex);
        }

        /// <summary>
        /// Generate a server tag name from a given prototype name template and address.
        /// </summary>
        /// <param name="tagPrototypeEntry">Tag prototype entry's format to use.</param>
        /// <param name="address">Address to substitute in.</param>
        /// <returns>Formatted server tag name.</returns>
        public string GenerateServerTagNameByAddress(ServerTagPrototypeEntry tagPrototypeEntry, int address)
        {
            tagPrototypeEntry.ThrowIfNull(nameof(tagPrototypeEntry));

            string tagName = tagPrototypeEntry.ServerTagNameTemplate.Replace("{SERVER}", RtacServerName);
            return string.Format(tagName, address);
        }

        /// <summary>
        /// Increment generation base address by the given amount.
        /// </summary>
        /// <param name="rtacTagName">Server tag type name.</param>
        /// <param name="incrementVal">Value to increment base address by.</param>
        public void IncrementRtacTagBaseAddressByRtacTagType(string rtacTagName, int incrementVal)
        {
            TagTypeRunningAddressOffset[rtacTagName] += incrementVal;
        }

        /// <summary>
        /// Replace standard placeholders in columns.
        /// </summary>
        /// <param name="rtacDataRow">Row of data to replace placeholders.</param>
        /// <param name="rtacTagName">Server tag name.</param>
        /// <param name="tagAddress">Server tag address.</param>
        /// <param name="tagAlias">Server tag alias.</param>
        public static void ReplaceRtacKeywords(OutputRowEntryDictionary rtacDataRow, string rtacTagName, string tagAddress, string tagAlias)
        {
            var replacements = new Dictionary<string, string>()
            {
                {
                    Keywords.NameKeyword,
                    rtacTagName
                },
                {
                    Keywords.AddressKeyword,
                    tagAddress
                },
                {
                    Keywords.AliasKeyword,
                    tagAlias
                }
            };

            // Replace keywords
            rtacDataRow.ReplaceTagKeywords(replacements);
        }

        /// <summary>
        /// Add a tag row to the output type collection.
        /// </summary>
        /// <param name="rtacTagTypeName">Type name to add output to.</param>
        /// <param name="rtacRow">Data to add.</param>
        public void AddRtacTagOutput(string rtacTagTypeName, OutputRowEntryDictionary rtacRow)
        {
            if (!RtacOutputList.ContainsKey(rtacTagTypeName))
                RtacOutputList[rtacTagTypeName] = new OutputList();

            RtacOutputList[rtacTagTypeName].Add(rtacRow);
        }

        /// <summary>
        /// Return the alias of a server tag given a SCADA name and direction.
        /// </summary>
        /// <param name="scadaName">SCADA name to process.</param>
        /// <param name="pointType">Used to determine if the control suffix needs to be appended.</param>
        /// <returns>Server tag alias.</returns>
        public string GetRtacAlias(string scadaName, PointTypeInfo pointType)
        {
            pointType.ThrowIfNull(nameof(pointType));

            if (pointType.IsControl)
                scadaName += Keywords.ControlKeyword;

            foreach (var s in TagAliasSubstitutes)
                scadaName = scadaName.Replace(s.Key, s.Value);

            return AliasNameTemplate.Replace(Keywords.NameKeyword, scadaName);
        }

        /// <summary>
        /// Validate a tag alias. Throws error if invalid.
        /// </summary>
        /// <param name="tagAlias">Tag alias to validate.</param>
        public static void ValidateTagAlias(string tagAlias)
        {
            var r = Regex.Match(tagAlias, @"^[A-Za-z0-9_]+\s*$", RegexOptions.None);
            if (!r.Success)
                throw new TagGenerationException("Invalid tag name: " + tagAlias);
        }

        /// <summary>
        /// Write all servers tag types to CSV.
        /// </summary>
        /// <param name="path">Source filename to append output suffix on.</param>
        public void WriteAllServerTags(string path)
        {
            foreach (var tagGroup in RtacOutputList)
                WriteServerTagCSV(tagGroup, path);
        }

        /// <summary>
        /// Write the specified server tag type to CSV.
        /// </summary>
        /// <param name="type">Tag type to write out.</param>
        /// <param name="path">Source filename to append output suffix on.</param>
        private void WriteServerTagCSV(KeyValuePair<string, OutputList> type, string path)
        {
            string typeName = type.Key;
            var tagGroup = type.Value;

            var comparer = new BySortingColumn(RtacTagPrototypes[typeName].SortingColumn);
            tagGroup.Sort(comparer);

            if (!RtacTagPrototypes.ContainsKey(typeName))
                throw new TagGenerationException("Unable to locate tag prototype.\r\n\r\nMissing: \"" + typeName + "\" in tag prototype.");

            string csvPath = System.IO.Path.GetDirectoryName(path) + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileNameWithoutExtension(path) + "_RtacServerTags_" + typeName + ".csv";
            using (var csvStreamWriter = new System.IO.StreamWriter(csvPath, false))
            {
                using (var csvWriter = new CsvHelper.CsvWriter(csvStreamWriter))
                {
                    // Remove address hack from earlier in the generation section
                    tagGroup.ForEach(x => x[RtacTagPrototypes[typeName].SortingColumn] = Math.Truncate(Convert.ToDouble(x[RtacTagPrototypes[typeName].SortingColumn])).ToString());

                    foreach (var tag in tagGroup)
                    {
                        foreach (var s in SharedUtils.OutputRowEntryDictionaryToArray(tag))
                            csvWriter.WriteField(s);
                        csvWriter.NextRecord();
                    }
                }
            }
        }
    }

    /// <summary>
    /// Server tag prototype child structure
    /// </summary>
    public class ServerTagPrototypeEntry
    {
        /// <summary>
        /// Server tag name format with placeholder for address.
        /// </summary>
        /// <remarks>Markup supported by String.Format supported</remarks>
        public string ServerTagNameTemplate { get; set; }

        /// <summary>
        /// Standard data all server tags of this type have
        /// </summary>
        public OutputRowEntryDictionary StandardColumns { get; } = new OutputRowEntryDictionary();
    }

    /// <summary>
    /// Class to store server type map info.
    /// </summary>
    public class ServerTagMapInfo
    {
        /// <summary>Name of the server tag type.</summary>
        public string ServerTagTypeName { get; set; }
        /// <summary>Indicates whether to substitute nominal data when the source tag is bad quality.</summary>
        public bool PerformQualityWrapping { get; set; }
    }

    /// <summary>
    /// Server tag prototype root structure
    /// </summary>
    public class ServerTagRootPrototype
    {
        /// <summary>
        /// List of child prototype entries. Non-array types have 1 entry.
        /// </summary>
        /// <remarks>Array types like DNPC etc will have multiple formats</remarks>
        public List<ServerTagPrototypeEntry> TagPrototypeEntries { get; } = new List<ServerTagPrototypeEntry>();

        /// <summary>
        /// Column to sort on
        /// </summary>
        public int SortingColumn { get; set; }

        /// <summary>
        /// Point type: either binary / analog and status / control.
        /// </summary>
        public PointTypeInfo PointType { get; set; }

        /// <summary>
        /// If the type is an analog type with limits this denotes the min and max column range that stores those limits.
        /// </summary>
        /// <remarks>
        /// Used for calculating nominal analog values for quality substitution.
        /// For binary points, both tuple values are the same.
        /// </remarks>
        public (int lower, int upper) NominalColumns { get; set; } = (-1, -1);
    }

    /// <summary>
    /// Parse tag information.
    /// </summary>
    /// <remarks>Helper class to parse tag info for array-capable tags</remarks>
    public class ServerTagInfo
    {
        /// <summary>
        /// Root tag type name, such as DNPC
        /// </summary>
        public string RootServerTagTypeName { get; private set; }

        private string fullServerTagTypeName;
        /// <summary>
        /// Full tag type name, such as DNPC[2]
        /// </summary>
        public string FullServerTagTypeName
        {
            get => fullServerTagTypeName;
            set => ParseServerTagTypeInfo(value);
        }

        /// <summary>
        /// Is tag an array type such as DNPC[2]
        /// </summary>
        public bool IsArray { get; private set; }

        /// <summary>
        /// Index of array tag types such as DNPC[2]
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Initialize a new instance of TagInfo.
        /// </summary>
        public ServerTagInfo() { }

        /// <summary>
        /// Initialize a new instance of TagInfo with the given tag name.
        /// </summary>
        /// <param name="fullServerTagTypeName">Tag type name to parse</param>
        public ServerTagInfo(string fullServerTagTypeName) => ParseServerTagTypeInfo(fullServerTagTypeName);

        /// <summary>
        /// Parse given tag type name into root type name and index,
        /// </summary>
        /// <param name="fullServerTagTypeName">Tag type name to parse</param>
        private void ParseServerTagTypeInfo(string fullServerTagTypeName)
        {
            // Note (?: ) is a non capture group
            var r = Regex.Match(fullServerTagTypeName, @"(\w+)(?:\[(\d+)\])?", RegexOptions.None);
            if (!r.Success)
                throw new TagGenerationException("Invalid tag type name: " + fullServerTagTypeName);

            this.fullServerTagTypeName = fullServerTagTypeName;
            RootServerTagTypeName = r.Groups[1].Value;
            IsArray = r.Groups[2].Length > 0;
            Index = IsArray ? Convert.ToInt32(r.Groups[2].Value) : 0;
        }
    }
}
