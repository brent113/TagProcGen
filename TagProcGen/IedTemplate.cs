﻿using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using OutputRowEntryDictionary = System.Collections.Generic.Dictionary<int, string>;

namespace TagProcGen
{

    /// <summary>
    /// Stores information used to generate server tags, SCADA tags, and the map between them
    /// </summary>
    public class IedTemplate
    {
        /// <summary>
        /// Create a new instance
        /// </summary>
        /// <param name="templateName">Name of template</param>
        public IedTemplate(string templateName) => TemplateName = templateName;

        /// <summary>
        /// Key: Pointer Name. Value: Cell Reference
        /// </summary>
        public Dictionary<string, string> Pointers { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// IED template worksheet name
        /// </summary>
        public string TemplateName { get; }

        /// <summary>
        /// Stores device address alignment (i.e. 50 status points per device).
        /// </summary>
        /// <remarks>Make sure to not use more device addresses than allocated.</remarks>
        public Dictionary<string, string> Offsets { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Contains the SCADA and device data for all points in the template.
        /// </summary>
        /// <remarks>Main output of this template</remarks>
        public List<IedTagEntry> IedTagEntryList { get; } = new List<IedTagEntry>();

        /// <summary>
        /// List of device and SCADA name pairs to generate point lists for.
        /// </summary>
        public List<IedScadaNamePair> IedScadaNames { get; } = new List<IedScadaNamePair>();

        /// <summary>
        /// Keywords that get replaced with other values.
        /// </summary>
        private class Keywords
        {
            /// <summary>IED Name</summary>
            public const string IedNameKeyword = "{IED}";
        }

        /// <summary>
        /// If invalid data is found throw an error.
        /// Checks for:
        /// - Entry has multiple entries that map to the same prototype entry (by index).
        /// - Maximum point number in a type is less than the offset.
        /// - Analog limits are present in pairs.
        /// - Binary nominals are either 1, 0, or -1
        /// - All filters refer to device contained in the template
        /// - Controls reference another point in the template
        /// </summary>
        /// <param name="rtacTemplate">RTAC template, used for tag information</param>
        public void Validate(RtacTemplate rtacTemplate)
        {
            rtacTemplate.ThrowIfNull(nameof(rtacTemplate));

            // Select list of groups of tags with multiple tag names that map to the same server tag entry.
            // In each iedTagEntry group the name / type list by the server mapped full type, ie DNPC[2].
            // There should only be 1 entry, if there's more that means 2 things map to the same MV for example.
            var tagNameTypeEntriesThatMapToSamePrototypeEntry = 
                IedTagEntryList.SelectMany(iedTagEntry => iedTagEntry.IedTagNameTypeList.GroupBy(iedTagNameTypePair => rtacTemplate.GetServerTagInfoByDevice(iedTagNameTypePair.IedTagTypeName).FullServerTagTypeName)
                                                                                        .Where(tagGroups => tagGroups.Count() > 1))
                .ToList();

            if (tagNameTypeEntriesThatMapToSamePrototypeEntry.Count > 0)
                throw new TagGenerationException(string.Format("Template {0} contains a multiple tags of the same type: \r\n{1} of type {2} and \r\n{3} of type {4}.",
                                                               rtacTemplate.TemplateName,
                                                               tagNameTypeEntriesThatMapToSamePrototypeEntry.First().ElementAt(0).IedTagName,
                                                               tagNameTypeEntriesThatMapToSamePrototypeEntry.First().ElementAt(0).IedTagTypeName,
                                                               tagNameTypeEntriesThatMapToSamePrototypeEntry.First().ElementAt(1).IedTagName,
                                                               tagNameTypeEntriesThatMapToSamePrototypeEntry.First().ElementAt(1).IedTagTypeName)
            );

            // Verify the maximum point number for each type is less than the offset.
            // Create an anonymous type of the first tag type's server type, the point number, and isAbsolute
            // Pick non-absolute addresses, sort, group by type, then select the single highest of each type
            // and filter by being equal to or higher than the offset.
            var maxPointNumberHigherThanOffsetByType = IedTagEntryList.Select(iedTagEntry => new
            {
                tagName = iedTagEntry.IedTagNameTypeList.First().IedTagName,
                serverType = rtacTemplate.GetServerTagInfoByDevice(iedTagEntry.IedTagNameTypeList.First().IedTagTypeName).RootServerTagTypeName,
                pointNumber = iedTagEntry.PointNumber,
                isAbsolute = iedTagEntry.PointNumberIsAbsolute
            })
                .Where(tagInfo => tagInfo.isAbsolute == false)
                .OrderByDescending(tagInfo => tagInfo.pointNumber)
                .GroupBy(tagInfo => tagInfo.serverType)
                .Select(groups => new
                {
                    groups.First().tagName,
                    groups.First().serverType,
                    groups.First().pointNumber
                })
                .Where(maxPoint => maxPoint.pointNumber >= Convert.ToDouble(Offsets[maxPoint.serverType]))
                .ToList();

            if (maxPointNumberHigherThanOffsetByType.Count > 0)
                throw new TagGenerationException(string.Format("Tag name \"{0}\" with point number {1} is greater than or equal to the offset for the data type {2} at {3}.",
                    maxPointNumberHigherThanOffsetByType.First().tagName, maxPointNumberHigherThanOffsetByType.First().pointNumber,
                    maxPointNumberHigherThanOffsetByType.First().serverType, Offsets[maxPointNumberHigherThanOffsetByType.First().serverType]));

            // Verify analog limits are defined in pairs
            var analogStatusTagData = IedTagEntryList.Select(x => new { tag = x, tagPrototype = rtacTemplate.GetServerTagPrototypeByDevice(x.IedTagNameTypeList.First().IedTagTypeName) })
                .Where(x => x.tagPrototype.PointType.IsAnalog & x.tagPrototype.PointType.IsStatus)
                .ToList();

            foreach (var tagData in analogStatusTagData)
            {
                var analogLimits = tagData.tag.ScadaColumns.Where(x => x.Key >= tagData.tagPrototype.NominalColumns.lower & x.Key <= tagData.tagPrototype.NominalColumns.upper)
                    .Where(x => !string.IsNullOrWhiteSpace(x.Value))
                    .OrderBy(x => Convert.ToDouble(x.Value))
                    .ToList();

                // Verify even number of limits
                if (analogLimits.Count % 2 != 0)
                    throw new TagGenerationException(string.Format("Tag name \"{0}\" has an odd number of limits. Limits must be in pairs.",
                        tagData.tag.IedTagNameTypeList.First().IedTagName));

                // Verify no duplicates
                if (analogLimits.Count != analogLimits.Distinct().Count())
                    throw new TagGenerationException(string.Format("Tag name \"{0}\" has a duplicate limit. Limits must be nested.",
                        tagData.tag.IedTagNameTypeList.First().IedTagName));
            }

            // Verify binary nominals are either 1, 0, or -1
            var binaryStatusTagsWithInvalidNominalState = IedTagEntryList.Select(x => new { tag = x, tagPrototype = rtacTemplate.GetServerTagPrototypeByDevice(x.IedTagNameTypeList.First().IedTagTypeName) })
                .Where(x => x.tagPrototype.PointType.IsBinary & x.tagPrototype.PointType.IsStatus)
                .Where(x =>
                {
                    if (!x.tag.ScadaColumns.ContainsKey(x.tagPrototype.NominalColumns.lower))
                        throw new TagGenerationException(string.Format("Tag \"{0}\" is missing required column #{1}", x.tag.IedTagNameTypeList.First().IedTagName, x.tagPrototype.NominalColumns.lower));
                    bool parseSuccess = int.TryParse(x.tag.ScadaColumns[x.tagPrototype.NominalColumns.lower], out int parseNumber);

                    // Select invalid
                    return !(parseSuccess && parseNumber >= -1 & parseNumber <= 1);
                })
                .ToList();
            if (binaryStatusTagsWithInvalidNominalState.Count > 0)
                throw new TagGenerationException(string.Format("Tag \"{0}\" has an invalid nominal state of \"{1}\".",
                    binaryStatusTagsWithInvalidNominalState.First().tag.IedTagNameTypeList.First().IedTagName,
                    binaryStatusTagsWithInvalidNominalState.First().tag.ScadaColumns[binaryStatusTagsWithInvalidNominalState.First().tagPrototype.NominalColumns.lower]));

            // Verify filters don't reference device not in the template
            var filtersWithDevicesNotInTemplate = IedTagEntryList.Where(tagEntry => tagEntry.DeviceFilter.DeviceList != null && tagEntry.DeviceFilter.DeviceList.Count > 0)
                .Where(tagEntry => tagEntry.DeviceFilter.DeviceList
                    .Where(deviceName => !IedScadaNames.Any(iedScadaNameTypePair => (iedScadaNameTypePair.IedName ?? "") == (deviceName ?? "")))
                    .Any())
                .Select(tagEntry => new { TagName = tagEntry.IedTagNameTypeList.First().IedTagName, FilterString = tagEntry.DeviceFilter.ToString() })
                .ToList();
            if (filtersWithDevicesNotInTemplate.Count > 0)
                throw new TagGenerationException(string.Format("Tag \"{0}\" has an invalid filter that references a device not in the template.\r\n\r\nFilter: {1}.",
                    filtersWithDevicesNotInTemplate.First().TagName,
                    filtersWithDevicesNotInTemplate.First().FilterString));

            // Verify controls reference another point in the template
            var pointNameInfo = IedTagEntryList.Select(iedTagEntry => new { scadaName = iedTagEntry.ScadaPointName, tagPointType = rtacTemplate.GetServerTagPrototypeByDevice(iedTagEntry.IedTagNameTypeList.First().IedTagTypeName).PointType })
                .Where(x => (x.scadaName ?? "") != "--")
                .ToList();
            var controlsWithNoLink = pointNameInfo.Where(x => x.tagPointType.IsControl)
                .Where(x => !pointNameInfo.Where(y => y.tagPointType.IsStatus)
                    .Where(y => (x.scadaName ?? "") == (y.scadaName ?? ""))
                    .Any())
                .ToList();
            if (controlsWithNoLink.Count > 0)
                throw new TagGenerationException(string.Format("Tag \"{0}\" is a control with no linked status point.",
                    controlsWithNoLink.First().scadaName));
        }

        /// <summary>
        /// Return existing or new tag entry.
        /// </summary>
        /// <param name="iedTagType">Device tag type. Used to look up associated tag types.</param>
        /// <param name="filter">Filter information.</param>
        /// <param name="pointNumber">Device tag address. Used to look up matching existing tags.</param>
        /// <param name="rtacTemplate">RTAC template to look up tag information in.</param>
        /// <returns>New or existing device tag data structure</returns>
        public IedTagEntry GetOrCreateTagEntry(string iedTagType, FilterInfo filter, int pointNumber, RtacTemplate rtacTemplate)
        {
            iedTagType.ThrowIfNull(nameof(iedTagType));
            filter.ThrowIfNull(nameof(filter));
            pointNumber.ThrowIfNull(nameof(pointNumber));
            rtacTemplate.ThrowIfNull(nameof(rtacTemplate));

            // Lookup root tag type name
            string DeviceTagServerTypeName = rtacTemplate.GetServerTagInfoByDevice(iedTagType).RootServerTagTypeName;

            // Search for tags in this template that have matching:
            // - Point numbers
            // - Root tag type names
            var TagArrayQuery = IedTagEntryList.Where(tagEntry => tagEntry.PointNumber == pointNumber)
                .Where(tagEntry => tagEntry.IedTagNameTypeList.Any(iedTagNameTypePair => (rtacTemplate.GetServerTagInfoByDevice(iedTagNameTypePair.IedTagTypeName).RootServerTagTypeName ?? "") == (DeviceTagServerTypeName ?? "")))
                .Where(tagEntry => tagEntry.DeviceFilter == filter)
                .ToList();

            if (TagArrayQuery.Count > 1)
                throw new TagGenerationException("Should not be more than 1 tag with the same point number and type: " + Convert.ToString(pointNumber) + ", " + iedTagType);

            if (TagArrayQuery.Count == 0)
            {
                var t = new IedTagEntry();
                IedTagEntryList.Add(t);
                return t;
            }
            else
                return TagArrayQuery[0];
        }

        /// <summary>
        /// Substitute placeholder with device name in tag.
        /// </summary>
        /// <param name="tagNameToUpdate">Device name with placeholder</param>
        /// <param name="iedName">Device name to substitue</param>
        /// <returns>Device name</returns>
        public static string SubstituteTagName(string tagNameToUpdate, string iedName)
        {
            tagNameToUpdate.ThrowIfNull(nameof(tagNameToUpdate));
            return tagNameToUpdate.Replace(Keywords.IedNameKeyword, iedName);
        }

        /// <summary>
        /// Get the linked tag for a given control point name.
        /// </summary>
        /// <param name="iedName">Name of the device to look up.</param>
        /// <param name="controlPointScadaName">SCADA point name to look up.</param>
        /// <param name="rtacTemplate">RTAC template to use for getting prototypes.</param>
        /// <returns>Device tag data of the linked point.</returns>
        public IedTagEntry GetLinkedStatusPoint(string iedName, string controlPointScadaName, RtacTemplate rtacTemplate)
        {
            iedName.ThrowIfNull(nameof(iedName));
            controlPointScadaName.ThrowIfNull(nameof(controlPointScadaName));
            rtacTemplate.ThrowIfNull(nameof(rtacTemplate));

            var search = IedTagEntryList.Where(iedTagEntry => rtacTemplate.GetServerTagPrototypeByDevice(iedTagEntry.IedTagNameTypeList.First().IedTagTypeName).PointType.IsStatus)
                .Where(iedTagEntry => iedTagEntry.ScadaPointName == controlPointScadaName)
                .Where(iedTagEntry => iedTagEntry.DeviceFilter.ShouldPointBeGenerated(iedName))
                .ToList();

            // This error should not be thrown because this specific state is checked for during validation.
            if (search.Count != 1)
                throw new TagGenerationException(string.Format("Search for linked point for tag \"{0}\" returned something other than exactly 1 result", controlPointScadaName));

            return search.First();
        }
    }

    /// <summary>
    /// Stores IED / SCADA name pair
    /// </summary>
    public class IedScadaNamePair
    {
        /// <summary>Device name.</summary>
        public string IedName { get; set; }
        /// <summary>SCADA name.</summary>
        public string ScadaName { get; set; }
    }

    /// <summary>
    /// Storage device tag name / tag type pair
    /// </summary>
    public class IedTagNameTypePair
    {
        /// <summary>Device name.</summary>
        public string IedTagName { get; set; }
        /// <summary>Tag type name.</summary>
        public string IedTagTypeName { get; set; }
    }

    /// <summary>
    /// Stores filter verb and list of devices.
    /// </summary>
    /// <remarks>Acceptable predicates are ALL, NOT device,list, or device,list</remarks>
    public class FilterInfo
    {
        private const string ALL_PREDICATE = "ALL";
        private const string NOT_PREDICATE = "NOT";
        private const char DELIMITER = ',';

        private readonly string _filterString;

        /// <summary>The filter verb.</summary>
        public FilterPredicateEnum FilterPredicate { get; set; }
        /// <summary>The list of devices to apply the filter to.</summary>
        public List<string> DeviceList { get; private set; }

        /// <summary>
        /// Create a new instance of FilterInfo from the given filter string.
        /// </summary>
        /// <param name="filterString">Text to generate predicate and device list from.</param>
        public FilterInfo(string filterString)
        {
            filterString.ThrowIfNull(nameof(filterString));

            _filterString = filterString;

            if (filterString.Length == 0)
                // Assume all
                FilterPredicate = FilterPredicateEnum.ALL;
            else if (filterString.StartsWith(ALL_PREDICATE))
                FilterPredicate = FilterPredicateEnum.ALL;
            else if (filterString.StartsWith(NOT_PREDICATE))
            {
                FilterPredicate = FilterPredicateEnum.NOT;
                DeviceList = filterString.Remove(0, NOT_PREDICATE.Length).Trim().Split(DELIMITER).Select(x => x.Trim()).ToList();
            }
            else
            {
                // SOME verb is implied by lack of other verbs.
                FilterPredicate = FilterPredicateEnum.SOME;
                DeviceList = filterString.Trim().Split(DELIMITER).Select(x => x.Trim()).ToList();
            }
        }

        /// <summary>
        /// Check a device name against a filter to determine if it should have the point generated.
        /// </summary>
        /// <param name="iedName">Device name to check against filter.</param>
        /// <returns>True if point should be generated for provided device name.</returns>
        public bool ShouldPointBeGenerated(string iedName)
        {
            if (FilterPredicate == FilterPredicateEnum.SOME)
                return DeviceList.Contains(iedName);
            else if (FilterPredicate == FilterPredicateEnum.NOT)
                return !DeviceList.Contains(iedName);
            else if (FilterPredicate == FilterPredicateEnum.ALL)
                return true;
            else
                throw new TagGenerationException($"Invalid Filter: {this.ToString()}");
        }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>The filter creation string.</returns>
        public override string ToString()
        {
            return _filterString;
        }

        /// <summary>
        /// Returns true if the filters have the same predicate and same device list
        /// </summary>
        /// <returns>If the filters are equal</returns>
        public static bool operator ==(FilterInfo a, FilterInfo b)
        {
            a.ThrowIfNull(nameof(a));
            b.ThrowIfNull(nameof(b));

            // Must be same value of null
            if (a.DeviceList == null != (b.DeviceList == null))
                return false;

            // If a is null, return predicate equality only.
            if (a.DeviceList == null)
                return (int)a.FilterPredicate == (int)b.FilterPredicate;

            // Default equality comparer
            return (int)a.FilterPredicate == (int)b.FilterPredicate && a.DeviceList.SequenceEqual(b.DeviceList);
        }

        /// <summary>
        /// Not equal operator for filters
        /// </summary>
        public static bool operator !=(FilterInfo a, FilterInfo b) => !(a == b);

        /// <summary>
        /// Override Equals, generated by VS
        /// </summary>
        public override bool Equals(object obj)
        {
            var info = obj as FilterInfo;
            return info != null &&
                   _filterString == info._filterString &&
                   FilterPredicate == info.FilterPredicate &&
                   EqualityComparer<List<string>>.Default.Equals(DeviceList, info.DeviceList);
        }

        /// <summary>
        /// GetHasCode Override, generated by VS
        /// </summary>
        public override int GetHashCode()
        {
            var hashCode = 582039096;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(_filterString);
            hashCode = hashCode * -1521134295 + FilterPredicate.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<List<string>>.Default.GetHashCode(DeviceList);
            return hashCode;
        }
    }

    /// <summary>
    /// Possible filter predicates.
    /// </summary>
    public enum FilterPredicateEnum
    {
        /// <summary>All devices will have the given entry generated.</summary>
        ALL = 0,
        /// <summary>Specified devices will have the given point generated.</summary>
        SOME = 1,
        /// <summary>All except the specified devices will have the given point generated.</summary>
        NOT = 2
    }

    /// <summary>
    /// Represents a single tag.
    /// </summary>
    public class IedTagEntry
    {
        /// <summary>Filter to conditionally exclude entry for certain devices.</summary>
        public FilterInfo DeviceFilter { get; set; }

        /// <summary>Device (usually) relative address. Unless marked as absolute, added to a running address offset to generate absolute addresses.</summary>
        public int PointNumber { get; set; }
        /// <summary>Treat point number as absolute address, don't add a running address offset to it.</summary>
        public bool PointNumberIsAbsolute { get; set; }

        /// <summary>List of all device tag name / type pairs that share the same address.</summary>
        /// <remarks>All types must resolve to the same root tag type.</remarks>
        public List<IedTagNameTypePair> IedTagNameTypeList { get; } = new List<IedTagNameTypePair>();

        /// <summary>Custom RTAC tag type worksheet column data.</summary>
        public OutputRowEntryDictionary RtacColumns { get; } = new OutputRowEntryDictionary(); // Key: Col #, 1 based, Value: Text

        /// <summary>SCADA points name.</summary>
        /// <remarks>Used for SCADA point names as well as RTAC tag aliases due to their human readability.</remarks>
        public string ScadaPointName { get; set; }
        /// <summary>Custom SCADA worksheet column data.</summary>
        public OutputRowEntryDictionary ScadaColumns { get; } = new OutputRowEntryDictionary(); // Key: Col #, 1 based, Value: Text
    }
}
