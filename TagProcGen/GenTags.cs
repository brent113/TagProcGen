using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using OutputRowEntryDictionary = System.Collections.Generic.Dictionary<int, string>;
using System.Windows.Forms;

namespace TagProcGen
{

    /// <summary>
    /// Main class that orchestrates and does the tag generation.
    /// </summary>
    public static class GenTags
    {

        /// <summary>
        /// List of global template reference lookup pairs
        /// </summary>
        private static readonly Dictionary<string, string> GlobalPointers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Worksheet References.
        /// <summary>Global Excel application reference.</summary>
        private static Excel.Application xlApp;
        /// <summary>Global Excel Workbook reference.</summary>
        private static Excel.Workbook xlWorkbook;
        /// <summary>Global template reference to definitions worksheet.</summary>
        private static Excel.Worksheet xlDef;

        // Data templates for storing data.
        /// <summary>RTAC server tags template.</summary>
        private static RtacTemplate TPL_Rtac;
        /// <summary>Every loaded device template.</summary>
        private static List<IedTemplate> IedTemplates;

        // Output worksheet generators.
        /// <summary>RTAC tag processor worksheet template</summary>
        private static RtacTagProcessorWorksheet TPL_TagProcessor;
        /// <summary>SCADA worksheet template.</summary>
        private static ScadaWorksheet TPL_Scada;


        /// <summary>
        /// Master function that orchestrates the generation process. Calls each responsible function in turn.
        /// </summary>
        /// <param name="path">Path to the Excel workbook containing the configuration.</param>
        /// <param name="logger">Log Notifier</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "This main function is the global error handler for all downstream errors")]
        public static void Generate(string path, INotifier logger)
        {
            logger.ThrowIfNull(nameof(logger));

            string ProcessingStep = "";
            try
            {
                ProcessingStep = "Initializing"; InitExcel(path);
                ProcessingStep = "Loading Templates"; LocateTemplates();
                ProcessingStep = "Loading Pointers"; LoadPointers();
                ProcessingStep = "Reading RTAC Template"; ReadRtac();
                ProcessingStep = "Reading SCADA"; ReadScada();

                foreach (var t in IedTemplates)
                {
                    ProcessingStep = "Reading Template " + t.TemplateName;
                    ReadTemplate(t);
                    t.Validate(TPL_Rtac);
                }
            }
            catch (TagGenerationException ex)
            {
                logger.Log("Could not successfully read template. Error text:\r\n\r\n" + ex.Message + "\r\n\r\nOccured while: " + ProcessingStep, "Error", LogSeverity.Error);
                return;
            }
            catch (Exception ex)
            {
                logger.Log("An unhandled exception has occured. Error text:\r\n\r\n" + ex.Message + "\r\n\r\nOccured while: " + ProcessingStep, "Error", LogSeverity.Error);
                return;
            }
            finally
            {
                xlWorkbook.Close(false);
                xlApp.Quit();
            }

            try
            {
                foreach (var t in IedTemplates)
                {
                    ProcessingStep = "Processing Template " + t.TemplateName;
                    GenIEDTagProcMap(t);
                }

                ProcessingStep = "Writing Tag Map";
                TPL_TagProcessor.WriteCsv(path, (RtacTagProcessorWorksheet.QualityWrapModeEnum)Convert.ToInt32(TPL_Rtac.Pointers[Constants.TplRtacTagProcWrapMode]));

                ProcessingStep = "Writing SCADA Tags";
                TPL_Scada.WriteAllSCADATags(path);

                ProcessingStep = "Writing RTAC Tags";
                TPL_Rtac.WriteAllServerTags(path);

                ProcessingStep = "Writing Fingerprint File";
                WriteFingerprint(path);
            }
            catch (TagGenerationException ex)
            {
                logger.Log("Could not successfully generate tag map. Error text:\r\n\r\n" + ex.Message + "\r\n\r\nOccured while: " + ProcessingStep, "Error", LogSeverity.Error);
                return;
            }
            catch (Exception ex)
            {
                logger.Log("An unhandled exception has occured. Error text:\r\n\r\n" + ex.Message + "\r\n\r\nOccured while: " + ProcessingStep, "Error", LogSeverity.Error);
                return;
            }

            logger.Log("Successfully generated tag processor map.\r\n\r\nLongest SCADA tag name: " + TPL_Scada.MaxValidatedTag + " at " + TPL_Scada.MaxValidatedTagLength.ToString() + " characters.", "Success", LogSeverity.Info);
        }

        /// <summary>
        /// Initialize an instance of Excel and load the workbook specified.
        /// </summary>
        /// <param name="Path">Path to the Excel workbook containing the configuration.</param>
        private static void InitExcel(string Path)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(Path);
        }

        /// <summary>
        /// Locate and load the worksheet templates in the workbook that need processing.
        /// </summary>
        private static void LocateTemplates()
        {
            xlDef = xlWorkbook.Sheets[Constants.TplDefSheet];

            TPL_TagProcessor = new RtacTagProcessorWorksheet();

            TPL_Rtac = new RtacTemplate(Constants.TplRtacSheet);

            TPL_Scada = new ScadaWorksheet(Constants.TplScadaSheet);

            IedTemplates = new List<IedTemplate>();
            var specialSheets = new[] { Constants.TplDefSheet, Constants.TplRtacSheet, Constants.TplScadaSheet };
            foreach (Excel.Worksheet sht in xlWorkbook.Sheets)
            {
                if (sht.Name.StartsWith(Constants.TplSheetPrefix) & !specialSheets.Contains(sht.Name))
                    IedTemplates.Add(new IedTemplate(sht.Name));
            }
        }

        /// <summary>
        /// Read the worksheet pointers from each template.
        /// </summary>
        private static void LoadPointers()
        {
            Excel.Worksheet sht;

            // Definition sheet pointers
            SharedUtils.ReadPairRange(xlDef.get_Range(Constants.TplDef), GlobalPointers, Constants.TplRtacDef,
                                      Constants.TplScadaDef, Constants.TplIedDef);

            // RTAC sheet pointers
            sht = xlWorkbook.Sheets[TPL_Rtac.TemplateName];
            SharedUtils.ReadPairRange(sht.get_Range(GlobalPointers[Constants.TplRtacDef]), TPL_Rtac.Pointers,
                                      Constants.TplRtacMapName, Constants.TplRtacTagProto, Constants.TplRtacTagMap,
                                      Constants.TplRtacAliasSub, Constants.TplRtacTagProcCols,
                                      Constants.TplRtacTagProcWrapMode);

            // SCADA sheet pointers
            sht = xlWorkbook.Sheets[TPL_Scada.TemplateName];
            SharedUtils.ReadPairRange(sht.get_Range(GlobalPointers[Constants.TplScadaDef]),
                                      TPL_Scada.Pointers, Constants.TplScadaNameFormat, Constants.TplScadaMaxNameLength,
                                      Constants.TplScadaTagProto, Constants.TplScadaAddressOffset);

            // IED pointers
            foreach (var t in IedTemplates)
            {
                // Pointers
                sht = xlWorkbook.Sheets[t.TemplateName];
                SharedUtils.ReadPairRange(sht.get_Range(GlobalPointers[Constants.TplIedDef]),
                                          t.Pointers, Constants.TplData, Constants.TplIedNames, Constants.TplOffsets);
            }
        }

        /// <summary>
        /// Read the RTAC template data.
        /// </summary>
        private static void ReadRtac()
        {
            Excel.Worksheet xlSheet = xlWorkbook.Sheets[TPL_Rtac.TemplateName];
            Excel.Range c;
            // Read name
            TPL_Rtac.RtacServerName = System.Convert.ToString(xlSheet.get_Range(TPL_Rtac.Pointers[Constants.TplRtacMapName]).Text);
            TPL_Rtac.AliasNameTemplate = System.Convert.ToString(xlSheet.get_Range(TPL_Rtac.Pointers[Constants.TplRtacMapName]).get_Offset(0, 1).Text);

            // Read tag prototypes, splitting when necessary
            c = xlSheet.get_Range(TPL_Rtac.Pointers[Constants.TplRtacTagProto]);
            while (!string.IsNullOrEmpty(Convert.ToString(c.Text)))
            {
                var tag = new ServerTagInfo(Convert.ToString(c.Text));
                string prototypeFormat = Convert.ToString(c.get_Offset(0, 1).Text);
                string colDataPairString = Convert.ToString(c.get_Offset(0, 2).Text);
                string sortingColumnRaw = Convert.ToString(c.get_Offset(0, 3).Text);
                string pointTypeInfoText = Convert.ToString(c.get_Offset(0, 4).Text);
                string analogLimitColumnRange = Convert.ToString(c.get_Offset(0, 5).Text);

                if (!int.TryParse(sortingColumnRaw, out int sortingColumn))
                    sortingColumn = -1;

                TPL_Rtac.AddTagPrototypeEntry(tag, prototypeFormat, colDataPairString, sortingColumn, pointTypeInfoText, analogLimitColumnRange);

                c = c.get_Offset(1, 0);
            }
            TPL_Rtac.ValidateTagPrototypes();

            // Read tag type map
            c = xlSheet.get_Range(TPL_Rtac.Pointers[Constants.TplRtacTagMap]);
            while (!string.IsNullOrEmpty(Convert.ToString(c.Text)))
            {
                string iedType = Convert.ToString(c.Text);
                string rtacType = Convert.ToString(c.get_Offset(0, 1).Text);
                string performQualityMappingRaw = Convert.ToString(c.get_Offset(0, 2).Text);

                bool parseSuccess = bool.TryParse(performQualityMappingRaw, out bool performQualityMapping);
                if (!parseSuccess)
                    throw new TagGenerationException(string.Format("Invalid quality wrapping flag for IED Type map entry {0}", iedType));

                TPL_Rtac.AddIedServerTagMap(iedType, rtacType, performQualityMapping);

                c = c.get_Offset(1, 0);
            }

            // Read Tag Alias Substitutions
            SharedUtils.ReadPairRange(xlSheet.get_Range(TPL_Rtac.Pointers[Constants.TplRtacAliasSub]), TPL_Rtac.TagAliasSubstitutes);

            // Read Tag processor Columns
            ((string)xlSheet.get_Range(TPL_Rtac.Pointers[Constants.TplRtacTagProcCols]).Text).ParseColumnDataPairs(TPL_TagProcessor.TagProcessorColumnsTemplate);
        }

        /// <summary>
        /// Read the SCADA template data.
        /// </summary>
        private static void ReadScada()
        {
            Excel.Worksheet xlSheet = xlWorkbook.Sheets[TPL_Scada.TemplateName];

            // Read name format
            var c = xlSheet.get_Range(TPL_Scada.Pointers[Constants.TplScadaNameFormat]);
            TPL_Scada.ScadaNameTemplate = Convert.ToString(c.Text);

            // Read SCADA prototypes
            c = xlSheet.get_Range(TPL_Scada.Pointers[Constants.TplScadaTagProto]);
            while (!string.IsNullOrEmpty(Convert.ToString(c.Text)))
            {
                string pointTypeName = Convert.ToString(c.Text);
                string defaultColumnData = Convert.ToString(c.get_Offset(0, 1).Text);
                string keyFormat = Convert.ToString(c.get_Offset(0, 2).Text);
                string csvHeader = Convert.ToString(c.get_Offset(0, 3).Text);
                string csvRowDefaults = Convert.ToString(c.get_Offset(0, 4).Text);
                string sortingColumnRaw = Convert.ToString(c.get_Offset(0, 5).Text);

                if (!int.TryParse(sortingColumnRaw, out int sortingColumn))
                    sortingColumn = -1;

                if (sortingColumn < 0)
                    throw new TagGenerationException(string.Format("SCADA prototype {0} is missing a valid sorting column.", pointTypeName));

                TPL_Scada.AddTagPrototypeEntry(pointTypeName, defaultColumnData, keyFormat, csvHeader, csvRowDefaults, sortingColumn);

                c = c.get_Offset(1, 0);
            }
        }

        /// <summary>
        /// Read the specified device template data.
        /// </summary>
        /// <param name="t">Device template to read.</param>
        private static void ReadTemplate(IedTemplate t)
        {
            t.ThrowIfNull(nameof(t));

            Excel.Worksheet xlSheet = xlWorkbook.Sheets[t.TemplateName];
            Excel.Range c;

            // Read offsets
            SharedUtils.ReadPairRange(xlSheet.get_Range(t.Pointers[Constants.TplOffsets]), t.Offsets);

            // Read IED and SCADA names
            c = xlSheet.get_Range(t.Pointers[Constants.TplIedNames]);
            while (!string.IsNullOrEmpty(Convert.ToString(c.Text)))
            {
                t.IedScadaNames.Add(new IedScadaNamePair()
                {
                    IedName = Convert.ToString(c.Text),
                    ScadaName = Convert.ToString(c.get_Offset(0, 1).Text)
                });

                c = c.get_Offset(1, 0);
            }

            // Read tag data
            c = xlSheet.get_Range(t.Pointers[Constants.TplData]);
            // for speed locate the last row, then do 1 large read
            while (!string.IsNullOrEmpty(Convert.ToString(c.get_Offset(10, 0).Text))) // read by 10s
                c = c.get_Offset(10, 0);
            while (!string.IsNullOrEmpty(Convert.ToString(c.get_Offset(1, 0).Text))) // read by 1s
                c = c.get_Offset(1, 0);
            var dataTable = xlSheet.get_Range(xlSheet.get_Range(t.Pointers[Constants.TplData]).Address + ":" + c.get_Offset(0, 7).Address).Value2;

            for (int i = 1, loopTo = dataTable.GetLength(0); i <= loopTo; i++)
            {
                bool Process = (dataTable[i, 1].ToString().ToUpper() ?? "") == "TRUE";
                if (Process)
                {
                    string filterRaw = dataTable[i, 2] != null ? dataTable[i, 2].ToString() : "";
                    string pointNumberRaw = dataTable[i, 3] != null ? dataTable[i, 3].ToString() : "";
                    string iedTagName = dataTable[i, 4] != null ? dataTable[i, 4].ToString() : "";
                    string iedTagType = dataTable[i, 5] != null ? dataTable[i, 5].ToString() : "";
                    string rtacColumns = dataTable[i, 6] != null ? dataTable[i, 6].ToString() : "";
                    string scadaPointName = dataTable[i, 7] != null ? dataTable[i, 7].ToString() : "";
                    string scadaColumns = dataTable[i, 8] != null ? dataTable[i, 8].ToString() : "";

                    int pointNumber;
                    bool pointNumberIsAbsolute;
                    if (pointNumberRaw.Length == 0)
                        throw new TagGenerationException("Point number missing");
                    pointNumberIsAbsolute = pointNumberRaw.Substring(0, 1) == "=";
                    pointNumber = Convert.ToInt32(pointNumberIsAbsolute ? pointNumberRaw.Substring(1) : pointNumberRaw);

                    var filter = new FilterInfo(filterRaw);
                    var dataEntry = t.GetOrCreateTagEntry(iedTagType, filter, pointNumber, TPL_Rtac);
                    dataEntry.DeviceFilter = filter;
                    dataEntry.PointNumber = pointNumber;
                    dataEntry.PointNumberIsAbsolute = pointNumberIsAbsolute;
                    dataEntry.IedTagNameTypeList.Add(new IedTagNameTypePair() { IedTagName = iedTagName, IedTagTypeName = iedTagType });

                    if (rtacColumns.Length > 0)
                        rtacColumns.ParseColumnDataPairs((OutputRowEntryDictionary)dataEntry.RtacColumns);
                    if (scadaPointName.Length > 0)
                        dataEntry.ScadaPointName = scadaPointName;
                    if (scadaColumns.Length > 0)
                        scadaColumns.ParseColumnDataPairs((OutputRowEntryDictionary)dataEntry.ScadaColumns);
                }
            }
        }

        /// <summary>
        /// This function does a few things:
        /// Generate SCADA output rows
        /// Generate RTAC output rows
        /// Generate Tag Map
        /// </summary>
        /// <param name="t">Template to generate data for.</param>
        private static void GenIEDTagProcMap(IedTemplate t)
        {
            t.ThrowIfNull(nameof(t));

            foreach (var iedScadaNamePair in t.IedScadaNames)
            {
                // Generate Data tag map and server tags from IEDs
                foreach (var tag in t.IedTagEntryList)
                {
                    tag.ScadaPointName.ThrowIfNull(nameof(tag.ScadaPointName));

                    // Skip tag generation if the name is blank or this tag is filtered out for this specific device name
                    if (tag.ScadaPointName == "--" || !tag.DeviceFilter.ShouldPointBeGenerated(iedScadaNamePair.IedName))
                        continue;

                    // Lookup RTAC tag info and prototype for later
                    string rtacTagInfoRootName = TPL_Rtac.GetServerTagInfoByDevice(tag.IedTagNameTypeList.First().IedTagTypeName).RootServerTagTypeName;
                    var newTagRootPrototype = TPL_Rtac.RtacTagPrototypes[rtacTagInfoRootName];
                    int addressBase = TPL_Rtac.TagTypeRunningAddressOffset[rtacTagInfoRootName];

                    // Calc in advance some basic info
                    int address = Convert.ToInt32(tag.PointNumberIsAbsolute ? tag.PointNumber : addressBase + tag.PointNumber);

                    string scadaFullName = "";

                    scadaFullName = TPL_Scada.ScadaNameGenerator(iedScadaNamePair.ScadaName, tag.ScadaPointName);

                    TPL_Scada.ValidateTagName(scadaFullName);

                    var scadaTagPrototype = TPL_Scada.ScadaTagPrototypes[newTagRootPrototype.PointType.ToString()];
                    var scadaColumns = new OutputRowEntryDictionary(scadaTagPrototype.StandardColumns); // Default SCADA columns

                    GenerateScadaTag(t, iedScadaNamePair, tag, newTagRootPrototype, address, scadaFullName, scadaTagPrototype, scadaColumns);
                    GenerateRtacTags(iedScadaNamePair, tag, rtacTagInfoRootName, newTagRootPrototype, address, scadaFullName, scadaColumns);
                }

                // Increment Server tag starting value by type offsets
                foreach (var offset in t.Offsets)
                    TPL_Rtac.IncrementRtacTagBaseAddressByRtacTagType(offset.Key, Convert.ToInt32(offset.Value));
            }
        }

        /// <summary>
        /// Generate SCADA Tags: format the data and store in the SCADA Tag Output
        /// </summary>
        /// <param name="t">IedTemplate</param>
        /// <param name="iedScadaNamePair">Which SCADA Device Name / IED Device Name</param>
        /// <param name="tag">Which tag</param>
        /// <param name="newTagRootPrototype">Which rtac prototype for the data type, to tell if it's a status or control data flow</param>
        /// <param name="address">Point address</param>
        /// <param name="scadaFullName">Full SCADA Name</param>
        /// <param name="scadaTagPrototype">SCADA tag prototype</param>
        /// <param name="scadaColumns">SCADA data</param>
        private static void GenerateScadaTag(IedTemplate t, IedScadaNamePair iedScadaNamePair, IedTagEntry tag, ServerTagRootPrototype newTagRootPrototype, int address, string scadaFullName, ScadaTagPrototype scadaTagPrototype, OutputRowEntryDictionary scadaColumns)
        {
            try
            {
                tag.ScadaColumns.ToList().ForEach(c => scadaColumns.Add(c.Key, c.Value)); // Custom SCADA columns
            }
            catch
            {
                throw new TagGenerationException("Invalid SCADA column definitions - duplicate columns present.");
            }

            if (newTagRootPrototype.PointType.IsStatus)
                // Replace keywords and add SCADA columns to output
                TPL_Scada.ReplaceScadaKeywords(scadaColumns, scadaFullName, address, scadaTagPrototype.KeyFormat);
            else
            {
                // Replace keywords. Specify separate key link address based on linked status point
                var linkedStatusPoint = t.GetLinkedStatusPoint(iedScadaNamePair.IedName, tag.ScadaPointName, TPL_Rtac);
                string linkedTagRootPrototype = TPL_Rtac.GetServerTagInfoByDevice(linkedStatusPoint.IedTagNameTypeList.First().IedTagTypeName).RootServerTagTypeName;
                int linkedAddressBase = TPL_Rtac.TagTypeRunningAddressOffset[linkedTagRootPrototype];
                int linkedAddress = Convert.ToInt32(linkedStatusPoint.PointNumberIsAbsolute ? linkedStatusPoint.PointNumber : linkedAddressBase + linkedStatusPoint.PointNumber);

                TPL_Scada.ReplaceScadaKeywords(scadaColumns, scadaFullName, address, scadaTagPrototype.KeyFormat, linkedAddress);
            }

            TPL_Scada.AddScadaTagOutput(newTagRootPrototype.PointType.ToString(), scadaColumns);
        }

        /// <summary>
        /// Generate RTAC Tags: format the data and store in the RTAC Tag Output
        /// </summary>
        /// <param name="iedScadaNamePair">Which SCADA Device Name / IED Device Name</param>
        /// <param name="tag">Which tag</param>
        /// <param name="rtacTagInfoRootName">RTAC tag data type, such as MV or DNPC</param>
        /// <param name="newTagRootPrototype">RTAC tag prototype</param>
        /// <param name="address">Point address</param>
        /// <param name="scadaFullName">SCADA tag name, used to generate the RTAC tag name</param>
        /// <param name="scadaColumns">SCADA data, for the purpose of nominal data generation during bad quality</param>
        private static void GenerateRtacTags(IedScadaNamePair iedScadaNamePair, IedTagEntry tag, string rtacTagInfoRootName, ServerTagRootPrototype newTagRootPrototype, int address, string scadaFullName, OutputRowEntryDictionary scadaColumns)
        {
            string rtacAlias = TPL_Rtac.GetRtacAlias(scadaFullName, newTagRootPrototype.PointType);
            RtacTemplate.ValidateTagAlias(rtacAlias);

            // Loop over each prototype, some like DNPC have several
            for (int index = 0, loopTo = newTagRootPrototype.TagPrototypeEntries.Count - 1; index <= loopTo; index++)
            {
                // Copy of default RTAC data
                var rtacColumns = new OutputRowEntryDictionary(newTagRootPrototype.TagPrototypeEntries[index].StandardColumns);

                try
                {
                    tag.RtacColumns.ToList().ForEach(c => rtacColumns.Add(c.Key, c.Value)); // Tag specific RTAC data
                }
                catch (ArgumentException)
                {
                    throw new TagGenerationException("Invalid RTAC column definitions - duplicate columns present.");
                }

                // Point name from format
                string tagName = TPL_Rtac.GenerateServerTagNameByAddress(newTagRootPrototype.TagPrototypeEntries[index], address);

                // Begin tag map processing
                // Check if there's an IED tag that maps to the current tag prototype
                int idx = index; // Required because iteration variables cannot be used in queries
                var iedTag = tag.IedTagNameTypeList
                    .Where(ied => TPL_Rtac.GetServerTagInfoByDevice(ied.IedTagTypeName).Index == idx)
                    .ToList();

                if (iedTag.Count > 1)
                    throw new TagGenerationException("Too many tags that map to " + rtacTagInfoRootName + ". Tag = " + iedTag.First().IedTagName);

                if (iedTag.Count == 1)
                {
                    string iedTagName = IedTemplate.SubstituteTagName(iedTag[0].IedTagName, iedScadaNamePair.IedName);
                    string iedTagTypeName = iedTag[0].IedTagTypeName;

                    var rtacTagInfo = TPL_Rtac.GetServerTagInfoByDevice(iedTagTypeName);
                    string rtacTagSuffix = TPL_Rtac.GetArraySuffix(rtacTagInfo);

                    var rtacServerTagTypeMap = TPL_Rtac.GetServerTypeByIedType(iedTagTypeName);

                    string rtacServerTagName = "Tags." + rtacAlias + rtacTagSuffix;
                    string rtacServerTagType = rtacServerTagTypeMap.ServerTagTypeName;

                    TPL_TagProcessor.AddEntry(rtacServerTagName, rtacServerTagType, iedTagName, iedTagTypeName, newTagRootPrototype.PointType, scadaColumns, rtacServerTagTypeMap.PerformQualityWrapping, newTagRootPrototype.NominalColumns);
                }
                // Tag map processing done

                // Calculate address fractional addition below to maintain sort order later 
                // for when potentially duplicate addresses get sorted, ie: array type
                double fractionalAddress = Convert.ToDouble(index) / Convert.ToDouble(newTagRootPrototype.TagPrototypeEntries.Count);

                RtacTemplate.ReplaceRtacKeywords(rtacColumns, tagName, Convert.ToString((double)address + fractionalAddress), rtacAlias);
                TPL_Rtac.AddRtacTagOutput(rtacTagInfoRootName, rtacColumns);
            }
        }

        /// <summary>
        /// Write out a file containing a fingerprint capable of showing changes between template files.
        /// To solve the issue of IED specific filters, only write lines with an "ALL" filter
        /// </summary>
        /// <param name="path">Template path</param>
        private static void WriteFingerprint(string path)
        {
            return;

            string csvPath = System.IO.Path.GetDirectoryName(path) + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileNameWithoutExtension(path) + "_Fingerprint.csv";
            using (var csvStreamWriter = new System.IO.StreamWriter(csvPath, false))
            {
                using (var csvWriter = new CsvHelper.CsvWriter(csvStreamWriter))
                {
                    var header = "";
                    header.Split(',').ToList().ForEach(x => csvWriter.WriteField(x));
                    csvWriter.NextRecord();

                    foreach (var t in IedTemplates)
                    {

                    }
                }
            }
        }
    }
}
