Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports OutputList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Stores information used to generate server tags, SCADA tags, and the map between them
''' </summary>
Public Class IedTemplate
    Private _Pointers As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Key: Pointer Name. Value: Cell Reference
    ''' </summary>
    Public ReadOnly Property Pointers As Dictionary(Of String, String)
        Get
            Return _Pointers
        End Get
    End Property

    Private _xlSheet As Excel.Worksheet
    ''' <summary>
    ''' Excel worksheet corresponding to the IED template
    ''' </summary>
    Public ReadOnly Property xlSheet As Excel.Worksheet
        Get
            Return _xlSheet
        End Get
    End Property

    ''' <summary>
    ''' Create a new instance
    ''' </summary>
    ''' <param name="xlSheet">Excel worksheet corresponding to the SCADA template</param>
    Public Sub New(xlSheet As Excel.Worksheet)
        _xlSheet = xlSheet
    End Sub

    Private _Offsets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Stores device address alignment (i.e. 50 status points per device).
    ''' </summary>
    ''' <remarks>Make sure to not use more device addresses than allocated.</remarks>
    Public ReadOnly Property Offsets As Dictionary(Of String, String)
        Get
            Return Me._Offsets
        End Get
    End Property

    Private _AllTags As New List(Of IedTagEntry)
    ''' <summary>
    ''' Contains the SCADA and device data for all points in the template.
    ''' </summary>
    ''' <remarks>Main output of this template</remarks>
    Public ReadOnly Property AllIedPoints As List(Of IedTagEntry)
        Get
            Return Me._AllTags
        End Get
    End Property

    ''' <summary>
    ''' Stores IED / SCADA name pair
    ''' </summary>
    Public Class IedScadaNamePair
        ''' <summary>Device name.</summary>
        Public IedName As String
        ''' <summary>SCADA name.</summary>
        Public ScadaName As String
    End Class

    Private _IedScadaNames As New List(Of IedScadaNamePair)
    ''' <summary>
    ''' List of device and SCADA name pairs to generate point lists for.
    ''' </summary>
    Public ReadOnly Property IedScadaNames As List(Of IedScadaNamePair)
        Get
            Return Me._IedScadaNames
        End Get
    End Property

    ''' <summary>
    ''' Keywords that get replaced with other values.
    ''' </summary>
    Public Class Keywords
        Public Const IED_NAME_KEYWORD = "{IED}"
    End Class

    ''' <summary>
    ''' Represents a single tag.
    ''' </summary>
    Public Class IedTagEntry
        ''' <summary>Filter to conditionally exclude entry for certain devices.</summary>
        Public DeviceFilter As FilterInfo

        ''' <summary>Device (usually) relative address. Unless marked as absolute, added to a running address offset to generate absolute addresses.</summary>
        Public PointNumber As Integer
        ''' <summary>Treat point number as absolute address, don't add a running address offset to it.</summary>
        Public PointNumberIsAbsolute As Boolean

        ''' <summary>List of all device tag name / type pairs that share the same address.</summary>
        ''' <remarks>All types must resolve to the same root tag type.</remarks>
        Public IedTagNameTypeList As New List(Of IedTagNameTypePair)

        ''' <summary>Custom RTAC tag type worksheet column data.</summary>
        Public RtacColumns As New OutputRowEntryDictionary ' Key: Col #, 1 based, Value: Text

        ''' <summary>SCADA points name.</summary>
        ''' <remarks>Used for SCADA point names as well as RTAC tag aliases due to their human readability.</remarks>
        Public ScadaPointName As String
        ''' <summary>Custom SCADA worksheet column data.</summary>
        Public ScadaColumns As New OutputRowEntryDictionary ' Key: Col #, 1 based, Value: Text
    End Class

    ''' <summary>
    ''' Stores filter verb and list of devices.
    ''' </summary>
    ''' <remarks>Acceptable predicates are ALL, NOT device,list, or device,list</remarks>
    Public Class FilterInfo
        Private Const ALL_PREDICATE = "ALL"
        Private Const NOT_PREDICATE = "NOT"
        Private Const DELIMITER = ","c

        Private _filterString As String

        ''' <summary>The filter verb.</summary>
        Public FilterPredicate As FilterPredicateEnum
        ''' <summary>The list of devices to apply the filter to.</summary>
        Public DeviceList As List(Of String)

        ''' <summary>
        ''' Create a new instance of FilterInfo from the given filter string.
        ''' </summary>
        ''' <param name="filterString">Text to generate predicate and device list from.</param>
        Public Sub New(filterString As String)
            Me._filterString = filterString

            If filterString.Length = 0 Then
                ' Assume all
                FilterPredicate = FilterPredicateEnum.ALL
            Else
                If filterString.StartsWith(ALL_PREDICATE) Then
                    FilterPredicate = FilterPredicateEnum.ALL
                ElseIf filterString.StartsWith(NOT_PREDICATE) Then
                    FilterPredicate = FilterPredicateEnum.NOT
                    DeviceList = filterString.Remove(0, NOT_PREDICATE.Length).Trim.Split(DELIMITER).Select(Function(x) x.Trim).ToList
                Else
                    ' SOME verb is implied by lack of other verbs.
                    FilterPredicate = FilterPredicateEnum.SOME
                    DeviceList = filterString.Trim.Split(DELIMITER).Select(Function(x) x.Trim).ToList
                End If
            End If
        End Sub

        ''' <summary>
        ''' Check a device name against a filter to determine if it should have the point generated.
        ''' </summary>
        ''' <param name="iedName">Device name to check against filter.</param>
        ''' <returns>True if point should be generated for provided device name.</returns>
        Public Function ShouldPointBeGenerated(iedName As String) As Boolean
            If Me.FilterPredicate = FilterPredicateEnum.SOME Then
                Return DeviceList.Contains(iedName)
            ElseIf Me.FilterPredicate = FilterPredicateEnum.NOT Then
                Return Not DeviceList.Contains(iedName)
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Returns a string that represents the current object.
        ''' </summary>
        ''' <returns>The filter creation string.</returns>
        Public Overrides Function ToString() As String
            Return Me._filterString
        End Function

        ''' <summary>
        ''' Returns true if the filters have the same predicate and same device list
        ''' </summary>
        ''' <returns>If the filters are equal</returns>
        Public Shared Operator =(a As FilterInfo, b As FilterInfo)
            ' Must be same value of null
            If (a.DeviceList Is Nothing) <> (b.DeviceList Is Nothing) Then Return False

            ' If a is null, return predicate equality only.
            If (a.DeviceList Is Nothing) Then Return a.FilterPredicate = b.FilterPredicate

            ' Default equality comparer
            Return (a.FilterPredicate = b.FilterPredicate) And a.DeviceList.SequenceEqual(b.DeviceList)
        End Operator

        Public Shared Operator <>(a As FilterInfo, b As FilterInfo)
            Return Not a = b
        End Operator

    End Class

    ''' <summary>
    ''' Possible filter predicates.
    ''' </summary>
    Public Enum FilterPredicateEnum
        ''' <summary>All devices will have the given entry generated.</summary>
        ALL = 0
        ''' <summary>Specified devices will have the given point generated.</summary>
        SOME = 1
        ''' <summary>All except the specified devices will have the given point generated.</summary>
        [NOT] = 2
    End Enum

    ''' <summary>
    ''' If invalid data is found throw an error.
    ''' Checks for:
    '''  - Entry has multiple entries that map to the same prototype entry (by index).
    '''  - Maximum point number in a type is less than the offset.
    '''  - Analog limits are present in pairs.
    '''  - Binary nominals are either 1, 0, or -1
    '''  - All filters refer to device contained in the template
    '''  - Controls reference another point in the template
    ''' </summary>
    ''' <param name="rtacTemplate">RTAC template to resolve data types in.</param>
    Public Sub Validate(rtacTemplate As RtacTemplate)
        ' Select list of groups of tags with multiple tag names that map to the same server tag entry.
        ' In each iedTagEntry group the name / type list by the server mapped full type, ie DNPC[2].
        ' There should only be 1 entry, if there's more that means 2 things map to the same MV for example.
        Dim tagNameTypeEntriesThatMapToSamePrototypeEntry = Me._AllTags.SelectMany(
            Function(iedTagEntry) iedTagEntry.IedTagNameTypeList.GroupBy(
                Function(iedTagNameTypePair) rtacTemplate.GetServerTagInfoByDevice(iedTagNameTypePair.IedTagTypeName).FullServerTagTypeName
            ).Where(
                Function(tagGroups) tagGroups.Count > 1
            )
        ).ToList

        If tagNameTypeEntriesThatMapToSamePrototypeEntry.Count > 0 Then
            Throw New Exception(
                String.Format("Template {0} contains a multiple tags of the same type: " & vbCrLf &
                              "{1} of type {2} and " & vbCrLf &
                              "{3} of type {4}.",
                              Me.xlSheet.Name,
                              tagNameTypeEntriesThatMapToSamePrototypeEntry.First.ElementAt(0).IedTagName,
                              tagNameTypeEntriesThatMapToSamePrototypeEntry.First.ElementAt(0).IedTagTypeName,
                              tagNameTypeEntriesThatMapToSamePrototypeEntry.First.ElementAt(1).IedTagName,
                              tagNameTypeEntriesThatMapToSamePrototypeEntry.First.ElementAt(1).IedTagTypeName
                )
            )
        End If

        ' Verify the maximum point number for each type is less than the offset.
        ' Create an anonymous type of the first tag type's server type, the point number, and isAbsolute
        ' Pick non-absolute addresses, sort, group by type, then select the single highest of each type
        ' and filter by being equal to or higher than the offset.
        Dim maxPointNumberHigherThanOffsetByType = Me._AllTags.Select(
            Function(iedTagEntry) New With {
                Key .tagName = iedTagEntry.IedTagNameTypeList.First.IedTagName,
                Key .serverType = rtacTemplate.GetServerTagInfoByDevice(iedTagEntry.IedTagNameTypeList.First.IedTagTypeName).RootServerTagTypeName,
                Key .pointNumber = iedTagEntry.PointNumber,
                Key .isAbsolute = iedTagEntry.PointNumberIsAbsolute
            }
        ).Where(
            Function(tagInfo) tagInfo.isAbsolute = False
        ).OrderByDescending(
            Function(tagInfo) tagInfo.pointNumber
        ).GroupBy(
            Function(tagInfo) tagInfo.serverType
        ).Select(
            Function(groups) New With {
                Key groups.First.tagName,
                Key groups.First.serverType,
                Key groups.First.pointNumber
            }
        ).Where(
            Function(maxPoint) maxPoint.pointNumber >= Me.Offsets(maxPoint.serverType)
        ).ToList

        If maxPointNumberHigherThanOffsetByType.Count > 0 Then
            Throw New Exception(
                String.Format("Tag name ""{0}"" with point number {1} is greater than or equal to the offset for the data type {2} at {3}.",
                    maxPointNumberHigherThanOffsetByType.First.tagName,
                    maxPointNumberHigherThanOffsetByType.First.pointNumber,
                    maxPointNumberHigherThanOffsetByType.First.serverType,
                    Me.Offsets(maxPointNumberHigherThanOffsetByType.First.serverType)
                )
            )
        End If

        ' Verify analog limits are defined in pairs
        Dim analogStatusTagData = Me._AllTags.Select(
            Function(x) New With {Key .tag = x, Key .tagPrototype = rtacTemplate.GetServerTagPrototypeByDevice(x.IedTagNameTypeList.First.IedTagTypeName)}
        ).Where(
            Function(x) x.tagPrototype.PointType.IsAnalog And x.tagPrototype.PointType.IsStatus
        ).ToList

        For Each tagData In analogStatusTagData
            Dim analogLimits = tagData.tag.ScadaColumns.Where(
                Function(x) x.Key >= tagData.tagPrototype.NominalColumns.Item1 And
                    x.Key <= tagData.tagPrototype.NominalColumns.Item2
            ).Where(
                Function(x) Not String.IsNullOrWhiteSpace(x.Value)
            ).OrderBy(
                Function(x) CDbl(x.Value)
            ).ToList

            ' Verify even number of limits
            If analogLimits.Count Mod 2 <> 0 Then
                Throw New Exception(
                    String.Format("Tag name ""{0}"" has an odd number of limits. Limits must be in pairs.",
                                  tagData.tag.IedTagNameTypeList.First.IedTagName
                                  )
                )
            End If

            ' Verify no duplicates
            If analogLimits.Count <> analogLimits.Distinct.Count Then
                Throw New Exception(
                    String.Format("Tag name ""{0}"" has a duplicate limit. Limits must be nested.",
                                  tagData.tag.IedTagNameTypeList.First.IedTagName
                                  )
                )
            End If
        Next

        ' Verify binary nominals are either 1, 0, or -1
        Dim binaryStatusTagsWithInvalidNominalState = Me._AllTags.Select(
            Function(x) New With {Key .tag = x, Key .tagPrototype = rtacTemplate.GetServerTagPrototypeByDevice(x.IedTagNameTypeList.First.IedTagTypeName)}
        ).Where(
            Function(x) x.tagPrototype.PointType.IsBinary And x.tagPrototype.PointType.IsStatus
        ).Where(
            Function(x)
                If Not x.tag.ScadaColumns.ContainsKey(x.tagPrototype.NominalColumns.Item1) Then
                    Throw New Exception(String.Format("Tag ""{0}"" is missing required column #{1}",
                                                      x.tag.IedTagNameTypeList.First.IedTagName,
                                                      x.tagPrototype.NominalColumns.Item1))
                End If

                Dim parseNumber As Integer, parseSuccess As Boolean
                parseSuccess = Integer.TryParse(x.tag.ScadaColumns(x.tagPrototype.NominalColumns.Item1), parseNumber)

                ' Select invalid
                Return Not (parseSuccess AndAlso (parseNumber >= -1 And parseNumber <= 1))
            End Function
        ).ToList
        If binaryStatusTagsWithInvalidNominalState.Count > 0 Then
            Throw New Exception(
                String.Format("Tag ""{0}"" has an invalid nominal state of ""{1}"".",
                    binaryStatusTagsWithInvalidNominalState.First.tag.IedTagNameTypeList.First.IedTagName,
                    binaryStatusTagsWithInvalidNominalState.First.tag.ScadaColumns(binaryStatusTagsWithInvalidNominalState.First.tagPrototype.NominalColumns.Item1)
                )
            )
        End If

        ' Verify filters don't reference device not in the template
        Dim filtersWithDevicesNotInTemplate = Me._AllTags.Where(
            Function(tagEntry) tagEntry.DeviceFilter.DeviceList IsNot Nothing AndAlso tagEntry.DeviceFilter.DeviceList.Count > 0
        ).Where(
            Function(tagEntry) tagEntry.DeviceFilter.DeviceList.Where(
                Function(deviceName) Not Me.IedScadaNames.Any(
                    Function(iedScadaNameTypePair) iedScadaNameTypePair.IedName = deviceName
                )
            ).Count > 0
        ).Select(
            Function(tagEntry) New With {Key .TagName = tagEntry.IedTagNameTypeList.First.IedTagName,
                                         Key .FilterString = tagEntry.DeviceFilter.ToString}
        ).ToList()
        If filtersWithDevicesNotInTemplate.Count > 0 Then
            Throw New Exception(
                String.Format("Tag ""{0}"" has an invalid filter that references a device not in the template." & vbCrLf & vbCrLf & "Filter: {1}.",
                    filtersWithDevicesNotInTemplate.First.TagName,
                    filtersWithDevicesNotInTemplate.First.FilterString
                )
            )
        End If

        ' Verify controls reference another point in the template
        Dim pointNameInfo = Me._AllTags.Select(
            Function(iedTagEntry) New With {Key .scadaName = iedTagEntry.ScadaPointName,
                                            Key .tagPointType = rtacTemplate.GetServerTagPrototypeByDevice(iedTagEntry.IedTagNameTypeList.First.IedTagTypeName).PointType}
        ).Where(
            Function(x) x.scadaName <> "--"
        ).ToList
        Dim controlsWithNoLink = pointNameInfo.Where(
            Function(x) x.tagPointType.IsControl
        ).Where(
            Function(x) pointNameInfo.Where(
                Function(y) y.tagPointType.IsStatus
            ).Where(
                Function(y) x.scadaName = y.scadaName
            ).Count = 0
        ).ToList
        If controlsWithNoLink.Count > 0 Then
            Throw New Exception(
                String.Format("Tag ""{0}"" is a control with no linked status point.",
                    controlsWithNoLink.First.scadaName
                )
            )
        End If
    End Sub

    ''' <summary>
    ''' Storage device tag name / tag type pair
    ''' </summary>
    Public Class IedTagNameTypePair
        ''' <summary>Device name.</summary>
        Public IedTagName As String
        ''' <summary>Tag type name.</summary>
        Public IedTagTypeName As String
    End Class

    ''' <summary>
    ''' Return existing or new tag entry.
    ''' </summary>
    ''' <param name="iedTagType">Device tag type. Used to look up associated tag types.</param>
    ''' <param name="filter">Filter information.</param>
    ''' <param name="pointNumber">Device tag address. Used to look up matching existing tags.</param>
    ''' <param name="rtacTemplate">RTAC template to look up tag information in.</param>
    ''' <returns>New or existing device tag data structure</returns>
    Public Function GetOrCreateTagEntry(iedTagType As String, filter As FilterInfo, pointNumber As Integer, rtacTemplate As RtacTemplate) As IedTagEntry
        ' Lookup root tag type name
        Dim DeviceTagServerTypeName = rtacTemplate.GetServerTagInfoByDevice(iedTagType).RootServerTagTypeName

        ' Search for tags in this template that have matching:
        '  - Point numbers
        '  - Root tag type names
        Dim TagArrayQuery = Me.AllIedPoints.Where(
            Function(tagEntry) tagEntry.PointNumber = pointNumber
        ).Where(
            Function(tagEntry) tagEntry.IedTagNameTypeList.Any(
                Function(iedTagNameTypePair) rtacTemplate.GetServerTagInfoByDevice(iedTagNameTypePair.IedTagTypeName).RootServerTagTypeName = DeviceTagServerTypeName
            )
        ).Where(
            Function(tagEntry) tagEntry.DeviceFilter = filter
        ).ToList

        If TagArrayQuery.Count > 1 Then
            Throw New Exception("Should not be more than 1 tag with the same point number and type: " & pointNumber & ", " & iedTagType)
        End If

        If TagArrayQuery.Count = 0 Then
            Dim t = New IedTagEntry
            Me.AllIedPoints.Add(t)
            Return t
        Else
            Return TagArrayQuery(0)
        End If
    End Function

    ''' <summary>
    ''' Substitute placeholder with device name in tag.
    ''' </summary>
    ''' <param name="tagNameToUpdate">Device name with placeholder</param>
    ''' <param name="iedName">Device name to substitue</param>
    ''' <returns>Device name</returns>
    Public Shared Function SubstituteTagName(tagNameToUpdate As String, iedName As String) As String
        Return tagNameToUpdate.Replace(Keywords.IED_NAME_KEYWORD, iedName)
    End Function

    ''' <summary>
    ''' Get the linked tag for a given control point name.
    ''' </summary>
    ''' <param name="iedName">Name of the device to look up.</param>
    ''' <param name="controlPointScadaName">SCADA point name to look up.</param>
    ''' <param name="rtacTemplate">RTAC template to use for getting prototypes.</param>
    ''' <returns>Device tag data of the linked point.</returns>
    Public Function GetLinkedStatusPoint(iedName As String, controlPointScadaName As String, rtacTemplate As RtacTemplate) As IedTagEntry
        Dim search = Me._AllTags.Where(
            Function(x) rtacTemplate.GetServerTagPrototypeByDevice(x.IedTagNameTypeList.First.IedTagTypeName).PointType.IsStatus
        ).Where(
            Function(x) x.ScadaPointName = controlPointScadaName
        ).Where(
            Function(x) x.DeviceFilter.ShouldPointBeGenerated(iedName)
        ).ToList()

        ' This error should not be thrown because this specific state is checked for during validation.
        If search.Count <> 1 Then
            Throw New Exception(String.Format("Search for linked point for tag ""{0}"" returned something other than exactly 1 result",
                                              controlPointScadaName))
        End If

        Return search.First
    End Function
End Class