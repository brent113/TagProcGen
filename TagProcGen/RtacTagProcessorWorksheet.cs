using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Text.RegularExpressions;
using OutputRowEntryDictionary = System.Collections.Generic.Dictionary<int, string>;
using OutputList = System.Collections.Generic.List<System.Collections.Generic.Dictionary<int, string>>;

namespace TagProcGen
{

    /// <summary>
    /// Builds the RTAC tag processor map
    /// </summary>
    public class RtacTagProcessorWorksheet
    {
        /// <summary>
        /// Keywords that get replaced with other values.
        /// </summary>
        public class Keywords
        {
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
            public const string DESTINATION = "{DESTINATION}";
            public const string DESTINATION_TYPE = "{DESTINATION_TYPE}";
            public const string SOURCE = "{SOURCE}";
            public const string SOURCE_TYPE = "{SOURCE_TYPE}";
            public const string TIME_SOURCE = "{TIME_SOURCE}";
            public const string QUALITY_SOURCE = "{QUALITY_SOURCE}";
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
        }

        /// <summary>
        /// Quality wrapping mode enumeration.
        /// </summary>
        public enum QualityWrapModeEnum
        {
            /// <summary>Don't wrap tags with quality check.</summary>
            None = 0,
            /// <summary>Check 1 tag's quality and perform data substitution based on that one tag for the entire device.</summary>
            GroupAllByDevice = 1,
            /// <summary>Individually check the first tag of a device, group the rest of the device's tags into a shared quality check.</summary>
            WrapFirstGroupRestByDevice = 2,
            /// <summary>Individually check each tag's quality and substitute data for that tag only.</summary>
            WrapIndividually = 3
        }

        /// <summary>Output list of tag processor entries.</summary>
        private List<TagProcessorMapEntry> _Map = new List<TagProcessorMapEntry>();

        /// <summary>Output list of all tag processor columns.</summary>
        private readonly OutputList _TagProcessorOutputRows = new OutputList();

        /// <summary>Tag processor columns template to generate output column order.</summary>
        public OutputRowEntryDictionary TagProcessorColumnsTemplate { get; } = new OutputRowEntryDictionary();

        /// <summary>
        /// Add new entry to the tag processor map.
        /// </summary>
        /// <param name="scadaTag">SCADA tag name</param>
        /// <param name="scadaTagDataType">SCADA tag types</param>
        /// <param name="iedTagName">Device tag name</param>
        /// <param name="iedDataType">Device tag type</param>
        /// <param name="pointType">Is the point status or control and analog or binary.</param>
        /// <param name="scadaRow">SCADA row entry. Used for calculating nominal values.</param>
        /// <param name="performQualityWrapping">Indicates wheter to substitute nominal data with the source tag quality is bad.</param>
        /// <param name="nominalValueColumns">Which columns in the SCADA data should be used to generate nominal values.</param>
        public void AddEntry(string scadaTag, string scadaTagDataType, string iedTagName, string iedDataType, PointTypeInfo pointType, OutputRowEntryDictionary scadaRow, bool performQualityWrapping, Tuple<int, int> nominalValueColumns
    )
        {
            string destTag, destType, sourceTag, sourceType;
            if (pointType.IsStatus)
            {
                // Flow: IED -> SCADA
                destTag = scadaTag;
                destType = scadaTagDataType;
                sourceTag = iedTagName;
                sourceType = iedDataType;
            }
            else if (pointType.IsControl)
            {
                // Flow: SCADA -> IED
                destTag = iedTagName;
                destType = iedDataType;
                sourceTag = scadaTag;
                sourceType = scadaTagDataType;
            }
            else
                throw new ArgumentException("Invalid Direction");

            _Map.Add(new TagProcessorMapEntry(destTag, destType, sourceTag, sourceType, pointType, scadaRow, performQualityWrapping, nominalValueColumns
    )
    );
        }

        /// <summary>
        /// The tag processor at RTAC startup sends bad quality data that results in nuisance alarms.
        /// Wrap tag processor devices in IEC 61131-3 logic to substitute bad quality data with
        /// placeholder data that will not generate nuisance alarms.
        /// 
        /// Data determined as:
        /// - For status points: SCADA Normal State.
        /// - For analog points: If alarm limits exist, something that satisfies high / low limits.
        /// 
        /// Can write out in multiple ways: By device, by first tag then device group, or by tag.
        /// By first tag then device group seems to be the best tradeoff betwen granularity and length.
        /// </summary>
        private void WrapDevicesWithQualitySubstitutions(QualityWrapModeEnum wrapMode)
        {
            if (wrapMode != (int)QualityWrapModeEnum.None)
            {
                var qualityWrappedMap = new List<TagProcessorMapEntry>();

                // Group tag processor into devices
                var qualityWrappingGroups = _Map.Where(mapEntry => mapEntry.PerformQualityWrapping
    ).GroupBy(mapEntry => mapEntry.ParsedDeviceName
    ).ToList();

                var NonQualityWrappingTags = _Map.Where(mapEntry => !mapEntry.PerformQualityWrapping
    ).ToList();

                if ((int)wrapMode == (int)QualityWrapModeEnum.GroupAllByDevice)
                {
                    // Basic - wrap every device in 1 quality group. Sometimes generates racked in/out alarms.
                    // Seems like the RTAC initializes some points earlier than other points
                    foreach (var deviceGroup in qualityWrappingGroups)
                    {
                        var deviceGroupList = deviceGroup.ToList();
                        var tagWrapper = new TagQualityWrapGenerator(deviceGroupList);

                        qualityWrappedMap.AddRange(tagWrapper.Generate());
                    }
                }
                else if ((int)wrapMode == (int)QualityWrapModeEnum.WrapFirstGroupRestByDevice)
                {
                    // Write the first tag, then write the remainder as a group. This means at least 2 tags are checked.
                    // Seems to be the best tradeoff between granularity and length in the tag processor.
                    foreach (var deviceGroup in qualityWrappingGroups)
                    {
                        var deviceGroupList = deviceGroup.ToList();

                        var first = new List<TagProcessorMapEntry> { deviceGroupList.First() };
                        var tagWrapper = new TagQualityWrapGenerator(first);
                        qualityWrappedMap.AddRange(tagWrapper.Generate());

                        var rest = deviceGroupList.Skip(1).ToList();
                        if (rest.Count > 0)
                        {
                            tagWrapper = new TagQualityWrapGenerator(rest);
                            qualityWrappedMap.AddRange(tagWrapper.Generate());
                        }
                    }
                }
                else if ((int)wrapMode == (int)QualityWrapModeEnum.WrapIndividually)
                {
                    // Write every point in its own quality tag. Guaranteed to work at the expense of tag processor length.
                    foreach (var deviceGroup in qualityWrappingGroups)
                    {
                        var deviceGroupList = deviceGroup.ToList();

                        foreach (var tagEntry in deviceGroupList.ToList())
                        {
                            var listOfOne = new List<TagProcessorMapEntry> { tagEntry };
                            var tagWrapper = new TagQualityWrapGenerator(listOfOne);

                            qualityWrappedMap.AddRange(tagWrapper.Generate());
                        }
                    }
                }

                _Map = qualityWrappedMap; // Update old map with quality wrapped map
                _Map.AddRange(NonQualityWrappingTags);
            }
        }

        /// <summary>
        /// Transformer Me._Map into output rows.
        /// </summary>
        private void GenerateOutputList()
        {
            foreach (var entry in _Map)
            {
                var outputRowEntry = new OutputRowEntryDictionary(TagProcessorColumnsTemplate);
                ReplaceRtacTagProcessorKeywords(outputRowEntry, entry.DestinationTagName, entry.DestinationTagDataType, entry.SourceExpression, entry.SourceExpressionDataType, entry.TimeSourceTagName, entry.QualitySourceTagName
    );
                _TagProcessorOutputRows.Add(outputRowEntry);
            }
        }

        /// <summary>
        /// Replace standard placeholders in columns.
        /// </summary>
        /// <param name="rtacTagProcessorRow">Output row to substitute placeholders in.</param>
        /// <param name="destination">Destination tag.</param>
        /// <param name="destinationType">Destination tag type.</param>
        /// <param name="source">Source tag.</param>
        /// <param name="sourceType">Source tag type.</param>
        /// <param name="timeSource">Time source tag.</param>
        /// <param name="qualitySource">Quality source tag.</param>
        private void ReplaceRtacTagProcessorKeywords(OutputRowEntryDictionary rtacTagProcessorRow, string destination, string destinationType, string source, string sourceType, string timeSource, string qualitySource
    )
        {
            var replacements = new Dictionary<string, string>()
            {
                {
                    Keywords.DESTINATION,
                    destination
                },
                {
                    Keywords.DESTINATION_TYPE,
                    destinationType
                },
                {
                    Keywords.SOURCE,
                    source
                },
                {
                    Keywords.SOURCE_TYPE,
                    sourceType
                },
                {
                    Keywords.TIME_SOURCE,
                    timeSource
                },
                {
                    Keywords.QUALITY_SOURCE,
                    qualitySource
                }
            };

            // Replace keywords
            rtacTagProcessorRow.ReplaceTagKeywords(replacements);
        }

        /// <summary>
        /// Write the tag processor map out to CSV
        /// </summary>
        /// <param name="path">Source filename to append output suffix on.</param>
        /// <param name="wrapMode">0-3, From no tag wrapping to wrap every tag individually.</param>
        public void WriteCsv(string path, QualityWrapModeEnum wrapMode)
        {
            WrapDevicesWithQualitySubstitutions(wrapMode);

            GenerateOutputList();

            string csvPath = System.IO.Path.GetDirectoryName(path) + Convert.ToString(System.IO.Path.DirectorySeparatorChar) + System.IO.Path.GetFileNameWithoutExtension(path) + "_TagProcessor.csv";
            using (var csvStreamWriter = new System.IO.StreamWriter(csvPath, false))
            {
                var csvWriter = new CsvHelper.CsvWriter(csvStreamWriter);

                foreach (var c in _TagProcessorOutputRows)
                {
                    foreach (var s in SharedUtils.OutputRowEntryDictionaryToArray(c))
                        csvWriter.WriteField(s);
                    csvWriter.NextRecord();
                }
            }
        }

        /// <summary>
        /// Intermediate storage format to generate quality wrapped tag processor entries.
        /// </summary>
        public class TagQualityWrapGenerator
        {
            /// <summary>Defines the conditional to be used to select bad quality points. Replace {TAG} with tag name.</summary>
            public const string QUALITY_CONDITIONAL_TEMPLATE = "IF ({TAG}.q.validity <> good) THEN";
            /// <summary>Else format.</summary>
            public const string ELSE_TEMPLATE = "ELSE";
            /// <summary>End if format.</summary>
            public const string END_IF_TEMPLATE = "END_IF";
            /// <summary>Time source template. Replace {TAG} with tag name.</summary>
            public const string TIME_SOURCE_TEMPLATE = "{TAG}.t";
            /// <summary>Quality source template. Replace {TAG} with tag name.</summary>
            public const string QUALITY_SOURCE_TEMPLATE = "{TAG}.q";
            /// <summary>Tag keyword to substitute in the templates.</summary>
            public const string TAG_KEYWORD = "{TAG}";

            /// <summary>
            /// List of tags to wrap with quality substitution.
            /// </summary>
            public List<TagProcessorMapEntry> TagsToWrap { get; }

            /// <summary>
            /// Initialize a new instance of TagQualityWrapGenerator.
            /// </summary>
            public TagQualityWrapGenerator()
            {
                TagsToWrap = new List<TagProcessorMapEntry>();
            }

            /// <summary>
            /// Initialze a new instance of TagQualityWrapGenerator.
            /// </summary>
            /// <param name="tagsToWrap">List of all tags to wrap.</param>
            public TagQualityWrapGenerator(List<TagProcessorMapEntry> tagsToWrap)
            {
                TagsToWrap = new List<TagProcessorMapEntry>(tagsToWrap);
            }

            /// <summary>
            /// Generate a new list of TagProcessorMapEntry classes that include a quality wrap with nominal values for bad quality tags.
            /// </summary>
            /// <returns>List of TagProcessorMapEntry classes with output row information.</returns>
            /// <remarks>
            /// Output format:
            /// If (tag.qual != good) then
            /// dest = nominal value
            /// Else
            /// dest = sourceExpr
            /// End_if
            /// </remarks>
            public List<TagProcessorMapEntry> Generate()
            {
                var outputTags = new List<TagProcessorMapEntry>();

                // Add first conditional for bad quality
                string firstTagName = TagsToWrap.First().ParsedTagName;
                outputTags.Add(GenerateConditionalTagEntry(QUALITY_CONDITIONAL_TEMPLATE, firstTagName));

                // Add nominal data
                foreach (var tag in TagsToWrap)
                {
                    string nominalValue = GetNominalValue(tag.PointType, tag.ScadaRow, tag.NominalValueColumns);
                    string timeSourceTag = TIME_SOURCE_TEMPLATE.Replace(TAG_KEYWORD, tag.ParsedTagName);
                    string qualitySourceTag = QUALITY_SOURCE_TEMPLATE.Replace(TAG_KEYWORD, tag.ParsedTagName);

                    var nominalTag = new TagProcessorMapEntry()
                    {
                        DestinationTagName = tag.DestinationTagName,
                        DestinationTagDataType = tag.DestinationTagDataType,
                        SourceExpression = nominalValue,
                        SourceExpressionDataType = "",
                        TimeSourceTagName = timeSourceTag,
                        QualitySourceTagName = qualitySourceTag
                    };

                    outputTags.Add(nominalTag);
                }

                // Add else
                outputTags.Add(GenerateConditionalTagEntry(ELSE_TEMPLATE));

                // Add original data mapping
                outputTags.AddRange(TagsToWrap);

                // Add end_if
                outputTags.Add(GenerateConditionalTagEntry(END_IF_TEMPLATE));

                return outputTags;
            }

            /// <summary>
            /// Generate a tag processor entry for a conditional with the given text and optional tag name.
            /// </summary>
            /// <param name="conditionalText">Conditional text to put into the source expression.</param>
            /// <param name="qualityTag">Optional placeholder to substitute.</param>
            /// <returns>Tag map entry with the given data.</returns>
            public static TagProcessorMapEntry GenerateConditionalTagEntry(string conditionalText, string qualityTag = null)
            {
                var pointType = new PointTypeInfo(true, true); // Status binary as a placeholder
                var mapEntry = new TagProcessorMapEntry()
                {
                    SourceExpression = qualityTag == null ? conditionalText : conditionalText.Replace(TAG_KEYWORD, qualityTag),
                    PointType = pointType
                };

                return mapEntry;
            }

            /// <summary>
            /// Get the nominal value of binary or analog status points from provided SCADA data.
            /// </summary>
            /// <param name="pointType">Point type information.</param>
            /// <param name="scadaColumns">SCADA data to derive nominal values from.</param>
            /// <param name="nominalValueColumns">Where to look in the SCADA data for the nominal values.</param>
            /// <returns>String that is the normal state for binaries or average of the two median analog alarms limits.</returns>
            public static string GetNominalValue(PointTypeInfo pointType, OutputRowEntryDictionary scadaColumns, Tuple<int, int> nominalValueColumns)
            {
                if (pointType.IsBinary)
                {
                    if (scadaColumns.ContainsKey(nominalValueColumns.Item1))
                        // Convert Boolean to IEC 61131-3 TRUE or FALSE
                        return Convert.ToBoolean(Convert.ToInt32(scadaColumns[nominalValueColumns.Item1])).ToString().ToUpper();
                    else
                        return "FALSE";
                }
                else
                {
                    var analogLimits = scadaColumns.Where(x => x.Key >= nominalValueColumns.Item1 & x.Key <= nominalValueColumns.Item2
                        ).Where(x => !string.IsNullOrWhiteSpace(x.Value)
                        ).OrderBy(x => Convert.ToDouble(x.Value)
                        ).ToList();
                    int middleStart = Convert.ToInt32(Math.Floor(analogLimits.Count / (double)2)) - 1;

                    // Must have conditional again in case there are no limits defined.
                    if (middleStart >= 0)
                    {
                        double average = (Convert.ToDouble(analogLimits[middleStart].Value) + Convert.ToDouble(analogLimits[middleStart + 1].Value)) / 2;
                        return average.ToString();
                    }
                    else
                        return "0";
                }
            }
        }

        /// <summary>
        /// Stores data for each tag processor map entry
        /// </summary>
        public class TagProcessorMapEntry
        {
            /// <summary>Destination tag name.</summary>
            public string DestinationTagName { get; set; }
            /// <summary>Destination tag datat type.</summary>
            public string DestinationTagDataType { get; set; }
            private string _SourceExpression;
            /// <summary>Source expression.</summary>
            public string SourceExpression
            {
                get
                {
                    return _SourceExpression;
                }
                set
                {
                    _SourceExpression = value;
                    ParseSourceExpression();
                }
            }
            /// <summary>Source expression tag type.</summary>
            public string SourceExpressionDataType { get; set; }
            /// <summary>Point type (i.e. status, analog).</summary>
            public PointTypeInfo PointType { get; set; }
            /// <summary>SCADA data associated with the source tag.</summary>
            /// <remarks>Used for generating SCADA nominal values that will not generate alarms.</remarks>
            public OutputRowEntryDictionary ScadaRow { get; set; }
            /// <summary>Indicates whether to substitute nominal data when the source tag quality is bad.</summary>
            public bool PerformQualityWrapping { get; set; }
            /// <summary>Column numbers in the SCADA row data that contains the nominal data information.</summary>
            public Tuple<int, int> NominalValueColumns { get; set; }

            /// <summary>Time source tag to use when using tag quality substitution.</summary>
            public string TimeSourceTagName { get; set; }
            /// <summary>Quality source tag to use when using tag quality substitution.</summary>
            public string QualitySourceTagName { get; set; }

            private string _ParsedDeviceName;
            /// <summary>Device name parsed from the source expression.</summary>
            public string ParsedDeviceName
            {
                get
                {
                    return _ParsedDeviceName;
                }
            }

            private string _ParsedTagName;
            /// <summary>Tag name parsed from the source expression.</summary>
            public string ParsedTagName
            {
                get
                {
                    return _ParsedTagName;
                }
            }

            private const string RegexMatch = @".*?(\p{L}\w*)\.(\p{L}\w*).*";
            private const string RegexTagName = "$1.$2";
            private const string RegexDeviceName = "$1";
            /// <summary>
            /// Locates the device name and point name from a source expression.
            /// </summary>
            private void ParseSourceExpression()
            {
                _ParsedDeviceName = Regex.Replace(SourceExpression, RegexMatch, RegexDeviceName);
                _ParsedTagName = Regex.Replace(SourceExpression, RegexMatch, RegexTagName);
            }

            /// <summary>Initialize a new instance of the TagProcessorMapEntry class.</summary>
            public TagProcessorMapEntry()
            {
            }

            /// <summary>
            /// Initialize a new instance of the TagProcessorMapEntry class.
            /// </summary>
            /// <param name="destinationTagName">Destination tag name.</param>
            /// <param name="destinationTagDataType">Destination tag data type.</param>
            /// <param name="sourceExpression">Source expression.</param>
            /// <param name="sourceExpressionDataType">Source expression datat type.</param>
            /// <param name="pointType">Point type (i.e. status, analog). Used for filtering entries in the tag processor.</param>
            /// <param name="scadaRow">SCADA data associtated with source tag.</param>
            /// <param name="performQualityWrapping">Indicates wheter to substitute nominal data with the source tag quality is bad.</param>
            /// <param name="nominalValueColumns">Which columns in the SCADA data should be used to generate nominal values.</param>
            public TagProcessorMapEntry(string destinationTagName, string destinationTagDataType, string sourceExpression, string sourceExpressionDataType, PointTypeInfo pointType, OutputRowEntryDictionary scadaRow, bool performQualityWrapping, Tuple<int, int> nominalValueColumns)
            {
                DestinationTagName = destinationTagName;
                DestinationTagDataType = destinationTagDataType;
                SourceExpression = sourceExpression;
                SourceExpressionDataType = sourceExpressionDataType;
                PointType = pointType;
                ScadaRow = scadaRow;
                PerformQualityWrapping = performQualityWrapping;
                NominalValueColumns = nominalValueColumns;

                ParseSourceExpression();
            }
        }
    }
}
