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
        public static Dictionary<string, string> GlobalPointers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Worksheet References.
        /// <summary>Global Excel application reference.</summary>
        public static Excel.Application xlApp;
        /// <summary>Global Excel Workbook reference.</summary>
        public static Excel.Workbook xlWorkbook;
        /// <summary>Global template reference to definitions worksheet.</summary>
        public static Excel.Worksheet xlDef;

        // Data templates for storing data.
        /// <summary>RTAC server tags template.</summary>
        public static RtacTemplate TPL_Rtac;
        /// <summary>Every loaded device template.</summary>
        public static List<IedTemplate> IedTemplates;

        // Output worksheet generators.
        /// <summary>RTAC tag processor worksheet template</summary>
        public static RtacTagProcessorWorksheet TPL_TagProcessor;
        /// <summary>SCADA worksheet template.</summary>
        public static ScadaWorksheet TPL_Scada;

        /// <summary>
        /// Master function that orchestrates the generation process. Calls each responsible function in turn.
        /// </summary>
        /// <param name="path">Path to the Excel workbook containing the configuration.</param>
        public static void Generate(string path)
        {
            string Processing = "";
            try
            {
                Processing = "Initializing"; InitExcel(path);
                Processing = "Loading Templates"; LocateTemplates();
                Processing = "Loading Pointers"; LoadPointers();
                Processing = "Reading RTAC Template"; ReadRtac();
                Processing = "Reading SCADA"; ReadScada();

                foreach (var t in IedTemplates)
                {
                    Processing = "Reading Template " + t.xlSheet.Name;
                    ReadTemplate(t);
                    t.Validate(TPL_Rtac);
                }

                foreach (var t in IedTemplates)
                {
                    Processing = "Processing Template " + t.xlSheet.Name;
                    GenIEDTagProcMap(t);
                }

                Processing = "Writing Tag Map";
                TPL_TagProcessor.WriteCsv(path, (RtacTagProcessorWorksheet.QualityWrapModeEnum)TPL_Rtac.Pointers[Constants.TPL_RTAC_TAG_PROC_WRAP_MODE]);

                Processing = "Writing SCADA Tags";
                TPL_Scada.WriteAllSCADATags(path);

                Processing = "Writing RTAC Tags";
                TPL_Rtac.WriteAllServerTags(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not successfully generate tag map. Error text:\r\n\r\n" + ex.Message + "\r\n\r\nOccured while: " + Processing, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                xlWorkbook.Close(false);
            }

            MessageBox.Show("Successfully generated tag processor map.\r\n\r\nLongest SCADA tag name: " + TPL_Scada.MaxValidatedTag + " at " + TPL_Scada.MaxValidatedTagLength.ToString() + " characters.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Initialize an instance of Excel and load the workbook specified.
        /// </summary>
        /// <param name="Path">Path to the Excel workbook containing the configuration.</param>
        public static void InitExcel(string Path)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(Path);
        }

        /// <summary>
        /// Locate and load the worksheet templates in the workbook that need processing.
        /// </summary>
        public static void LocateTemplates()
        {
            xlDef = xlWorkbook.Sheets[Constants.TPL_DEF_SHEET];

            TPL_TagProcessor = new RtacTagProcessorWorksheet();

            TPL_Rtac = new RtacTemplate(xlWorkbook.Sheets[Constants.TPL_RTAC_SHEET]);

            TPL_Scada = new ScadaWorksheet(xlWorkbook.Sheets[Constants.TPL_SCADA_SHEET]);

            IedTemplates = new List<IedTemplate>();
            var specialSheets = new[] { Constants.TPL_DEF_SHEET, Constants.TPL_RTAC_SHEET, Constants.TPL_SCADA_SHEET };
            foreach (Excel.Worksheet sht in xlWorkbook.Sheets)
            {
                if (sht.Name.StartsWith(Constants.TPL_SHEET_PREFIX) & !specialSheets.Contains(sht.Name))
                    IedTemplates.Add(new IedTemplate(sht));
            }
        }

        /// <summary>
        /// Read the worksheet pointers from each template.
        /// </summary>
        public static void LoadPointers()
        {
            // Definition sheet pointers
            SharedUtils.ReadPairRange(xlDef.get_Range(Constants.TPL_DEF), GlobalPointers, Constants.TPL_RTAC_DEF, Constants.TPL_SCADA_DEF, Constants.TPL_IED_DEF);

            // RTAC sheet pointers
            SharedUtils.ReadPairRange(TPL_Rtac.xlSheet.get_Range(GlobalPointers[Constants.TPL_RTAC_DEF]), TPL_Rtac.Pointers, Constants.TPL_RTAC_MAP_NAME, Constants.TPL_RTAC_TAG_PROTO, Constants.TPL_RTAC_TAG_MAP, Constants.TPL_RTAC_ALIAS_SUB, Constants.TPL_RTAC_TAG_PROC_COLS, Constants.TPL_RTAC_TAG_PROC_WRAP_MODE);

            // SCADA sheet pointers
            SharedUtils.ReadPairRange(TPL_Scada.xlSheet.get_Range(GlobalPointers[Constants.TPL_SCADA_DEF]), TPL_Scada.Pointers, Constants.TPL_SCADA_NAME_FORMAT, Constants.TPL_SCADA_MAX_NAME_LENGTH, Constants.TPL_SCADA_TAG_PROTO, Constants.TPL_SCADA_ADDRESS_OFFSET);

            // IED pointers
            foreach (var t in IedTemplates)
                // Pointers
                SharedUtils.ReadPairRange(t.xlSheet.get_Range(GlobalPointers[Constants.TPL_IED_DEF]), t.Pointers, Constants.TPL_DATA, Constants.TPL_IED_NAMES, Constants.TPL_OFFSETS);
        }

        /// <summary>
        /// Read the RTAC template data.
        /// </summary>
        public static void ReadRtac()
        {
            Excel.Range c;
            {
                var withBlock = TPL_Rtac;
                // Read name
                withBlock.RtacServerName = Conversions.ToString(withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_RTAC_MAP_NAME]).Text);
                withBlock.AliasNameTemplate = Conversions.ToString(withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_RTAC_MAP_NAME]).get_Offset(0, 1).Text);

                // Read tag prototypes, splitting when necessary
                c = withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_RTAC_TAG_PROTO]);
                while (!string.IsNullOrEmpty(Conversions.ToString(c.Text)))
                {
                    var tag = new RtacTemplate.ServerTagInfo(Conversions.ToString(c.Text));
                    string prototypeFormat = Conversions.ToString(c.get_Offset(0, 1).Text);
                    string colDataPairString = Conversions.ToString(c.get_Offset(0, 2).Text);
                    string sortingColumnRaw = Conversions.ToString(c.get_Offset(0, 3).Text);
                    string pointTypeInfoText = Conversions.ToString(c.get_Offset(0, 4).Text);
                    string analogLimitColumnRange = Conversions.ToString(c.get_Offset(0, 5).Text);

                    int sortingColumn = -1;
                    if (!int.TryParse(sortingColumnRaw, out sortingColumn))
                        sortingColumn = -1;

                    withBlock.AddTagPrototypeEntry(tag, prototypeFormat, colDataPairString, sortingColumn, pointTypeInfoText, analogLimitColumnRange
    );

                    c = c.get_Offset(1, 0);
                }
                TPL_Rtac.ValidateTagPrototypes();

                // Read tag type map
                c = withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_RTAC_TAG_MAP]);
                while (!string.IsNullOrEmpty(Conversions.ToString(c.Text)))
                {
                    string iedType = Conversions.ToString(c.Text);
                    string rtacType = Conversions.ToString(c.get_Offset(0, 1).Text);
                    string performQualityMappingRaw = Conversions.ToString(c.get_Offset(0, 2).Text);

                    bool performQualityMapping;
                    bool parseSuccess = bool.TryParse(performQualityMappingRaw, out performQualityMapping);
                    if (!parseSuccess)
                        throw new Exception(string.Format("Invalid quality wrapping flag for IED Type map entry {0}", iedType));

                    withBlock.AddIedServerTagMap(iedType, rtacType, performQualityMapping);

                    c = c.get_Offset(1, 0);
                }

                // Read Tag Alias Substitutions
                SharedUtils.ReadPairRange(withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_RTAC_ALIAS_SUB]), withBlock.TagAliasSubstitutes);

                // Read Tag processor Columns
                ((string)withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_RTAC_TAG_PROC_COLS]).Text).ParseColumnDataPairs(TPL_TagProcessor.TagProcessorColumnsTemplate);
            }
        }

        /// <summary>
        /// Read the SCADA template data.
        /// </summary>
        public static void ReadScada()
        {
            {
                var withBlock = TPL_Scada;
                // Read name format
                var c = withBlock.xlSheet.get_Range(TPL_Scada.Pointers[Constants.TPL_SCADA_NAME_FORMAT]);
                withBlock.ScadaNameTemplate = Conversions.ToString(c.Text);

                // Read SCADA prototypes
                c = withBlock.xlSheet.get_Range(withBlock.Pointers[Constants.TPL_SCADA_TAG_PROTO]);
                while (!string.IsNullOrEmpty(Conversions.ToString(c.Text)))
                {
                    string pointTypeName = Conversions.ToString(c.Text);
                    string defaultColumnData = Conversions.ToString(c.get_Offset(0, 1).Text);
                    string keyFormat = Conversions.ToString(c.get_Offset(0, 2).Text);
                    string csvHeader = Conversions.ToString(c.get_Offset(0, 3).Text);
                    string csvRowDefaults = Conversions.ToString(c.get_Offset(0, 4).Text);
                    string sortingColumnRaw = Conversions.ToString(c.get_Offset(0, 5).Text);

                    int sortingColumn = -1;
                    if (!int.TryParse(sortingColumnRaw, out sortingColumn))
                        sortingColumn = -1;

                    if (sortingColumn < 0)
                        throw new Exception(string.Format("SCADA prototype {0} is missing a valid sorting column.", pointTypeName));

                    withBlock.AddTagPrototypeEntry(pointTypeName, defaultColumnData, keyFormat, csvHeader, csvRowDefaults, sortingColumn);

                    c = c.get_Offset(1, 0);
                }
            }
        }

        /// <summary>
        /// Read the specified device template data.
        /// </summary>
        /// <param name="t">Device template to read.</param>
        public static void ReadTemplate(IedTemplate t)
        {
            Excel.Range c;

            // Read offsets
            SharedUtils.ReadPairRange(t.xlSheet.get_Range(t.Pointers[Constants.TPL_OFFSETS]), t.Offsets);

            // Read IED and SCADA names
            c = t.xlSheet.get_Range(t.Pointers[Constants.TPL_IED_NAMES]);
            while (!string.IsNullOrEmpty(Conversions.ToString(c.Text)))
            {
                t.IedScadaNames.Add(new IedTemplate.IedScadaNamePair()
                {
                    IedName = Conversions.ToString(c.Text),
                    ScadaName = Conversions.ToString(c.get_Offset(0, 1).Text)
                });

                c = c.get_Offset(1, 0);
            }

            // Read tag data
            c = t.xlSheet.get_Range(t.Pointers[Constants.TPL_DATA]);
            // for speed locate the last row, then do 1 large read
            while (!string.IsNullOrEmpty(Conversions.ToString(c.get_Offset(10, 0).Text))) // read by 10s
                c = c.get_Offset(10, 0);
            while (!string.IsNullOrEmpty(Conversions.ToString(c.get_Offset(1, 0).Text))) // read by 1s
                c = c.get_Offset(1, 0);
            var dataTable = t.xlSheet.get_Range(t.xlSheet.get_Range(t.Pointers[Constants.TPL_DATA]).Address + ":" + c.get_Offset(0, 7).Address).Value2;

            for (int i = 1, loopTo = dataTable.GetLength(0); i <= loopTo; i++)
            {
                bool Process = (dataTable(i, 1).ToString().ToUpper() ?? "") == "TRUE";
                if (Process)
                {
                    string filterRaw = dataTable(i, 2) != null ? dataTable(i, 2).ToString() : "";
                    string pointNumberRaw = dataTable(i, 3) != null ? dataTable(i, 3).ToString() : "";
                    string iedTagName = dataTable(i, 4) != null ? dataTable(i, 4).ToString() : "";
                    string iedTagType = dataTable(i, 5) != null ? dataTable(i, 5).ToString() : "";
                    string rtacColumns = dataTable(i, 6) != null ? dataTable(i, 6).ToString() : "";
                    string scadaPointName = dataTable(i, 7) != null ? dataTable(i, 7).ToString() : "";
                    string scadaColumns = dataTable(i, 8) != null ? dataTable(i, 8).ToString() : "";

                    int pointNumber;
                    bool pointNumberIsAbsolute;
                    if (pointNumberRaw.Length == 0)
                        throw new Exception("Point number missing");
                    pointNumberIsAbsolute = (pointNumberRaw.Substring(0, 1) ?? "") == "=";
                    pointNumber = Conversions.ToInteger(pointNumberIsAbsolute ? pointNumberRaw.Substring(1) : pointNumberRaw);

                    var filter = new IedTemplate.FilterInfo(filterRaw);
                    var dataEntry = t.GetOrCreateTagEntry(iedTagType, filter, pointNumber, TPL_Rtac);
                    {
                        var withBlock = dataEntry;
                        withBlock.DeviceFilter = new IedTemplate.FilterInfo(filterRaw);
                        withBlock.PointNumber = pointNumber;
                        withBlock.PointNumberIsAbsolute = pointNumberIsAbsolute;
                        withBlock.IedTagNameTypeList.Add(new IedTemplate.IedTagNameTypePair() { IedTagName = iedTagName, IedTagTypeName = iedTagType });

                        if (rtacColumns.Length > 0)
                            rtacColumns.ParseColumnDataPairs(withBlock.RtacColumns);
                        if (scadaPointName.Length > 0)
                            withBlock.ScadaPointName = scadaPointName;
                        if (scadaColumns.Length > 0)
                            scadaColumns.ParseColumnDataPairs(withBlock.ScadaColumns);
                    }
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
        public static void GenIEDTagProcMap(IedTemplate t)
        {
            foreach (var iedScadaNamePair in t.IedScadaNames)
            {
                // Generate Data tag map and server tags from IEDs
                foreach (var tag in t.AllIedPoints)
                {
                    // Skip tag generation if this tag is filtered out
                    if (!tag.DeviceFilter.ShouldPointBeGenerated(iedScadaNamePair.IedName))
                        continue;

                    // Begin calc in advance
                    // the address and alias. Calc the name for each format entry in the loop for each format

                    // Lookup RTAC tag info and prototype for later
                    string rtacTagInfoRootName = TPL_Rtac.GetServerTagInfoByDevice(tag.IedTagNameTypeList.First().IedTagTypeName).RootServerTagTypeName;
                    var newTagRootPrototype = TPL_Rtac.RtacTagPrototypes[rtacTagInfoRootName];
                    int addressBase = TPL_Rtac.TagTypeRunningAddressOffset[rtacTagInfoRootName];

                    // Calc in advance some basic info
                    int address = Conversions.ToInteger(Interaction.IIf(tag.PointNumberIsAbsolute, tag.PointNumber, addressBase + tag.PointNumber));
                    bool ProcessScada = (tag.ScadaPointName ?? "") != "--";
                    string scadaFullName = "";
                    string rtacAlias = "";
                    if (ProcessScada)
                    {
                        scadaFullName = TPL_Scada.ScadaNameGenerator(iedScadaNamePair.ScadaName, tag.ScadaPointName);

                        TPL_Scada.ValidateTagName(scadaFullName);

                        rtacAlias = TPL_Rtac.GetRtacAlias(scadaFullName, newTagRootPrototype.PointType);
                        TPL_Rtac.ValidateTagAlias(rtacAlias);
                    }
                    // End calc in advance

                    var scadaTagPrototype = TPL_Scada.ScadaTagPrototypes[newTagRootPrototype.PointType.ToString()];
                    var scadaColumns = new OutputRowEntryDictionary(scadaTagPrototype.StandardColumns); // Default SCADA columns
                    if (ProcessScada)
                    {
                        // Begin SCADA column processing
                        try
                        {
                            tag.ScadaColumns.ToList().ForEach(c => scadaColumns.Add(c.Key, c.Value)); // Custom SCADA columns
                        }
                        catch (Exception e)
                        {
                            throw new Exception("Invalid SCADA column definitions - duplicate columns present.");
                        }

                        // todo: replace with linked address if control
                        if (newTagRootPrototype.PointType.IsStatus)
                            // Replace keywords and add SCADA columns to output
                            TPL_Scada.ReplaceScadaKeywords(scadaColumns, scadaFullName, address, scadaTagPrototype.KeyFormat);
                        else
                        {
                            // Replace keywords. Specify separate key link address based on linked status point
                            var linkedStatusPoint = t.GetLinkedStatusPoint(iedScadaNamePair.IedName, tag.ScadaPointName, TPL_Rtac);
                            string linkedTagRootPrototype = TPL_Rtac.GetServerTagInfoByDevice(linkedStatusPoint.IedTagNameTypeList.First().IedTagTypeName).RootServerTagTypeName;
                            int linkedAddressBase = TPL_Rtac.TagTypeRunningAddressOffset[linkedTagRootPrototype];
                            int linkedAddress = Conversions.ToInteger(Interaction.IIf(linkedStatusPoint.PointNumberIsAbsolute, linkedStatusPoint.PointNumber, linkedAddressBase + linkedStatusPoint.PointNumber));

                            TPL_Scada.ReplaceScadaKeywords(scadaColumns, scadaFullName, address, scadaTagPrototype.KeyFormat, linkedAddress);
                        }


                        TPL_Scada.AddScadaTagOutput(newTagRootPrototype.PointType.ToString(), scadaColumns);
                    }

                    // Begin RTAC column processing
                    for (int index = 0, loopTo = newTagRootPrototype.TagPrototypeEntries.Count - 1; index <= loopTo; index++)
                    {
                        var rtacColumns = new OutputRowEntryDictionary(newTagRootPrototype.TagPrototypeEntries[index].StandardColumns); // Default RTAC Columns

                        try
                        {
                            tag.RtacColumns.ToList().ForEach(c => rtacColumns.Add(c.Key, c.Value)); // Custom RTAC columns
                        }
                        catch (Exception e)
                        {
                            throw new Exception("Invalid RTAC column definitions - duplicate columns present.");
                        }

                        // Point name from format
                        string tagName = TPL_Rtac.GenerateServerTagNameByAddress(newTagRootPrototype.TagPrototypeEntries[index], address);

                        if (ProcessScada)
                        {
                            // Begin tag map processing
                            // Check if there's an IED tag that maps to the current tag prototype
                            int idx = index; // Required because iteration variables cannot be used in queries
                            var iedTag = (from ied in tag.IedTagNameTypeList
                                          where TPL_Rtac.GetServerTagInfoByDevice(ied.IedTagTypeName).Index == idx
                                          select ied
    ).ToList();

                            if (iedTag.Count > 1)
                                throw new Exception("Too many tags that map to " + rtacTagInfoRootName + ". Tag = " + iedTag.First().IedTagName);

                            if (iedTag.Count == 1)
                            {
                                string iedTagName = IedTemplate.SubstituteTagName(iedTag[0].IedTagName, iedScadaNamePair.IedName);
                                string iedTagTypeName = iedTag[0].IedTagTypeName;

                                var rtacTagInfo = TPL_Rtac.GetServerTagInfoByDevice(iedTagTypeName);
                                string rtacTagSuffix = TPL_Rtac.GetArraySuffix(rtacTagInfo);

                                var rtacServerTagTypeMap = TPL_Rtac.GetServerTypeByIedType(iedTagTypeName);

                                string rtacServerTagName = "Tags." + rtacAlias + rtacTagSuffix;
                                string rtacServerTagType = rtacServerTagTypeMap.ServerTagTypeName;

                                TPL_TagProcessor.AddEntry(rtacServerTagName, rtacServerTagType, iedTagName, iedTagTypeName, newTagRootPrototype.PointType, scadaColumns, rtacServerTagTypeMap.PerformQualityWrapping, newTagRootPrototype.NominalColumns
    );
                            }
                            // Tag map processing done

                            // Calculate address fractional addition below to maintain sort order later 
                            // for when potentially duplicate addresses get sorted, ie: array type
                            double fractionalAddress = Conversions.ToDouble(index) / Conversions.ToDouble(newTagRootPrototype.TagPrototypeEntries.Count);

                            TPL_Rtac.ReplaceRtacKeywords(rtacColumns, tagName, Conversions.ToString(Conversions.ToDouble(address) + fractionalAddress), rtacAlias);
                            TPL_Rtac.AddRtacTagOutput(rtacTagInfoRootName, rtacColumns);
                        }
                    }
                }

                // Increment Server tag starting value by type offsets
                foreach (var offset in t.Offsets)
                    TPL_Rtac.IncrementRtacTagBaseAddressByRtacTagType(offset.Key, Conversions.ToInteger(offset.Value));
            }
        }
    }
}
