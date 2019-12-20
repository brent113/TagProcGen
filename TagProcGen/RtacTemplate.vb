Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports OutputList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Generates RTAC tags. Stores tag prototypes, handles server tag generation.
''' </summary>
Public Class RtacTemplate
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
    ''' Excel worksheet corresponding to the RTAC template
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

    ''' <summary>
    ''' Name of the SCADA server object in the RTAC.
    ''' </summary>
    Public Property RtacServerName As String

    ''' <summary>
    ''' Server tag alias template.
    ''' </summary>
    Public Property AliasNameTemplate As String

    Private _RtacTagPrototypes As New Dictionary(Of String, ServerTagRootPrototype)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Dictionary of server tag prototypes. Key: Server type name, Value: Prototype Root Type
    ''' </summary>
    Public ReadOnly Property RtacTagPrototypes As Dictionary(Of String, ServerTagRootPrototype)
        Get
            Return Me._RtacTagPrototypes
        End Get
    End Property

    Private _TagTypeRunningAddressOffset As New Dictionary(Of String, Integer)
    ''' <summary>
    ''' Starting value of next IED's tags. Incremented by offsets
    ''' </summary>
    Public ReadOnly Property TagTypeRunningAddressOffset As Dictionary(Of String, Integer)
        Get
            Return _TagTypeRunningAddressOffset
        End Get
    End Property

    ''' <summary>
    ''' Class to store server type map info.
    ''' </summary>
    Public Class ServerTagMapInfo
        ''' <summary>Name of the server tag type.</summary>
        Public ServerTagTypeName As String
        ''' <summary>Indicates whether to substitute nominal data when the source tag is bad quality.</summary>
        Public PerformQualityWrapping As Boolean
    End Class

    ''' <remarks>Contains 1:1 map to prototype. Key: Device tag type, Value: Server tag map info.</remarks>
    Private _IedToServerTypeMap As New Dictionary(Of String, ServerTagMapInfo)

    ''' <summary>
    ''' Add a new entry into the device-server tag map.
    ''' </summary>
    ''' <param name="iedTypeName">Device type name.</param>
    ''' <param name="serverTypeName">Server type name.</param>
    ''' <param name="performQualityWrapping">Indicates wheter to substitute nominal data with the source tag quality is bad.</param>
    Public Sub AddIedServerTagMap(iedTypeName As String,
                                  serverTypeName As String,
                                  performQualityWrapping As Boolean)
        Dim tagMapInfo As New ServerTagMapInfo With {.ServerTagTypeName = serverTypeName,
                                                     .PerformQualityWrapping = performQualityWrapping}
        Me._IedToServerTypeMap(iedTypeName) = tagMapInfo
    End Sub

    ''' <summary>
    ''' Get server tag map information for a given device type name.
    ''' </summary>
    ''' <param name="iedTypeName">Device type name.</param>
    ''' <returns>Server tag map information, or nothing if no entry exists.</returns>
    Public Function GetServerTypeByIedType(iedTypeName As String) As ServerTagMapInfo
        If Not Me._IedToServerTypeMap.ContainsKey(iedTypeName) Then Return Nothing
        Return Me._IedToServerTypeMap(iedTypeName)
    End Function

    Private _TagAliasSubstitutes As New Dictionary(Of String, String)
    ''' <summary>
    ''' Placeholders to search for and replace with associated value.
    ''' </summary>
    ''' <remarks>Key: Find, Value: Replace</remarks>
    Public ReadOnly Property TagAliasSubstitutes As Dictionary(Of String, String)
        Get
            Return Me._TagAliasSubstitutes
        End Get
    End Property

    ''' <summary>
    ''' List of all server tags by type. Generated from device templates.
    ''' </summary>
    Private _RtacOutputList As New Dictionary(Of String, OutputList)

    ''' <summary>
    ''' Keywords that get replaced with other values.
    ''' </summary>
    Public Class Keywords
        Public Const NAME_KEYWORD = "{NAME}"
        Public Const ADDRESS_KEYWORD = "{ADDRESS}"
        Public Const ALIAS_KEYWORD = "{ALIAS}"

        Public Const CONTROL_KEYWORD = "{CTRL}"
    End Class

    ''' <summary>
    ''' Server tag prototype root structure
    ''' </summary>
    Public Class ServerTagRootPrototype
        ''' <summary>
        ''' List of child prototype entries. Non-array types have 1 entry.
        ''' </summary>
        ''' <remarks>Array types like DNPC etc will have multiple formats</remarks>
        Public TagPrototypeEntries As New List(Of ServerTagPrototypeEntry)

        ''' <summary>
        ''' Column to sort on
        ''' </summary>
        Public SortingColumn As Integer

        ''' <summary>
        ''' Point type: either binary / analog and status / control.
        ''' </summary>
        Public PointType As PointTypeInfo

        ''' <summary>
        ''' If the type is an analog type with limits this denotes the min and max column range that stores those limits.
        ''' </summary>
        ''' <remarks>
        ''' Used for calculating nominal analog values for quality substitution.
        ''' For binary points, both tuple values are the same.
        ''' </remarks>
        Public NominalColumns As Tuple(Of Integer, Integer)
    End Class

    ''' <summary>
    ''' Server tag prototype child structure
    ''' </summary>
    Public Class ServerTagPrototypeEntry
        ''' <summary>
        ''' Server tag name format with placeholder for address.
        ''' </summary>
        ''' <remarks>Markup supported by String.Format supported</remarks>
        Public ServerTagNameTemplate As String

        ''' <summary>
        ''' Standard data all server tags of this type have
        ''' </summary>
        Public StandardColumns As New OutputRowEntryDictionary
    End Class

    ''' <summary>
    ''' Load a new server tag prototype entry. Creates new prototype or adds information to existing array prototype.
    ''' </summary>
    ''' <param name="tagInfo">Tag name and index information.</param>
    ''' <param name="nameTemplate">Formatting template for generated tags.</param>
    ''' <param name="defaultColumnData">Default data all tags have.</param>
    ''' <param name="sortingColumn">Column to sort alphanumerically on before writing out. Only needs to be specified once per prototype.</param>
    ''' <param name="pointTypeText">Type of the point, either binary / analog and status / control.</param>
    ''' <param name="nominalColumns">String denoting the presence of nominal columns in a format like "23" or "12:25".</param>
    Public Sub AddTagPrototypeEntry(
            tagInfo As ServerTagInfo,
            nameTemplate As String,
            defaultColumnData As String,
            sortingColumn As Integer,
            pointTypeText As String,
            nominalColumns As String
    )

        ' Get existing tag, or create new
        Dim tagRootPrototype As ServerTagRootPrototype
        If Me.RtacTagPrototypes.ContainsKey(tagInfo.RootServerTagTypeName) Then
            tagRootPrototype = Me.RtacTagPrototypes(tagInfo.RootServerTagTypeName)
        Else
            tagRootPrototype = New ServerTagRootPrototype With {.SortingColumn = -1}
            Me.RtacTagPrototypes.Add(tagInfo.RootServerTagTypeName, tagRootPrototype)
        End If

        ' Create entry in TagGenerationAddressBase
        TagTypeRunningAddressOffset(tagInfo.RootServerTagTypeName) = 0

        ' Add sorting column to root prototype if it is valid
        If sortingColumn > -1 Then
            tagRootPrototype.SortingColumn = sortingColumn
        End If

        'Add data direction to root prototype if it is valid
        If pointTypeText.Length > 0 Then
            tagRootPrototype.PointType = New PointTypeInfo(pointTypeText)
        End If

        ' Parse nominal column information
        If nominalColumns.Length > 0 Then
            Dim colonSplit = nominalColumns.Split({"."c, "["c, "]"c}, StringSplitOptions.RemoveEmptyEntries)

            If colonSplit.Length = 1 Then
                tagRootPrototype.NominalColumns = New Tuple(Of Integer, Integer)(colonSplit(0), colonSplit(0))
            ElseIf colonSplit.Length = 2 Then
                tagRootPrototype.NominalColumns = New Tuple(Of Integer, Integer)(colonSplit(0), colonSplit(1))

                ' Check for even number of columns - analog limits come in pairs. 1:10 = 10-1, should be odd
                If (tagRootPrototype.NominalColumns.Item2 - tagRootPrototype.NominalColumns.Item1) Mod 2 = 0 Then
                    Throw New Exception(
                        String.Format("Tag prototype {0} has an odd number of nominal value columns. Only even number of columns allowed.",
                                      tagInfo.RootServerTagTypeName
                        )
                    )
                End If
            Else
                Throw New Exception("Invalid analog limit column range. Expecting format like '10 or [11..20]'")
            End If
        End If

        ' Ensure the array has a placeholder for the incoming index
        For i = tagRootPrototype.TagPrototypeEntries.Count To tagInfo.Index
            tagRootPrototype.TagPrototypeEntries.Add(Nothing) ' Add placeholders
        Next

        ' Store prototype entry
        Dim newTagPrototypeEntry As New ServerTagPrototypeEntry
        With newTagPrototypeEntry
            .ServerTagNameTemplate = nameTemplate
            defaultColumnData.ParseColumnDataPairs(.StandardColumns)
        End With

        ' Store new prototype entry
        tagRootPrototype.TagPrototypeEntries(tagInfo.Index) = newTagPrototypeEntry
    End Sub

    ''' <summary>
    ''' Ensure all loaded tag prototypes have a valid sorting column, point information,
    ''' and status points have a nominal indication column.
    ''' </summary>
    Public Sub ValidateTagPrototypes()
        For Each ta In Me.RtacTagPrototypes
            If ta.Value.SortingColumn < 0 Then Throw New Exception(String.Format("Tag prototype {0} is missing a valid sorting column.", ta.Key))
            If ta.Value.PointType Is Nothing Then Throw New Exception(String.Format("Tag prototype {0} is missing a valid data direction.", ta.Key))
            If ta.Value.PointType.IsStatus AndAlso ta.Value.NominalColumns Is Nothing Then Throw New Exception(String.Format("Tag prototype {0} is a status type but is missing valid nominal columns.", ta.Key))
        Next
    End Sub

    ''' <summary>
    ''' Returns the information of a valid tag, otherwise throws an exception.
    ''' </summary>
    ''' <param name="iedTagName">Device tag name to validate.</param>
    ''' <returns>TagInfo structure of valid tag.</returns>
    Public Function ValidateTag(iedTagName As String) As ServerTagInfo
        Dim tagMapInfo = Me.GetServerTypeByIedType(iedTagName)
        If tagMapInfo Is Nothing Then
            Throw New ArgumentException("Unable to locate tag mapping." & vbCrLf & vbCrLf &
                                "Missing: """ & iedTagName & """ in tag map.")
        End If

        Dim Tag = New ServerTagInfo(tagMapInfo.ServerTagTypeName)
        If Not Me.RtacTagPrototypes.ContainsKey(Tag.RootServerTagTypeName) Then
            Throw New ArgumentException("Unable to locate tag prototype." & vbCrLf & vbCrLf &
                                "Missing: """ & Tag.RootServerTagTypeName & """ in tag prototype.")
        End If

        Return Tag
    End Function

    ''' <summary>
    ''' Return the server tag associated with th given device tag.
    ''' </summary>
    ''' <param name="iedTagType">Name of device tag to get server info for.</param>
    ''' <returns>Server tag information.</returns>
    ''' <remarks>Ex: operSPC-T -> DNPC, in TagInfo container</remarks>
    Public Function GetServerTagInfoByDevice(iedTagType As String) As ServerTagInfo
        Return Me.ValidateTag(iedTagType)
    End Function

    ''' <summary>
    ''' Return the server tag associated with th given device tag.
    ''' </summary>
    ''' <param name="iedTagType">Name of device tag to get server prototype for.</param>
    ''' <returns>Server tag root prototype.</returns>
    ''' <remarks>Ex: operSPC-T -> DNPC</remarks>
    Public Function GetServerTagPrototypeByDevice(iedTagType As String) As ServerTagRootPrototype
        Dim Tag = Me.GetServerTagInfoByDevice(iedTagType)

        Return Me.RtacTagPrototypes(Tag.RootServerTagTypeName)
    End Function

    ''' <summary>
    ''' Returns the specific server prototype entry associated with th given device tag.
    ''' </summary>
    ''' <param name="iedTagType">Name of device tag to get server prototype entry for.</param>
    ''' <returns>Server tag prototype entry.</returns>
    ''' <remarks>Ex: operSPC-T -> DNPC[2]</remarks>
    Public Function GetServerTagEntryByDevice(iedTagType As String) As ServerTagPrototypeEntry
        Dim Tag = Me.GetServerTagInfoByDevice(iedTagType)
        Dim TagInfo = Me.GetServerTagInfoByDevice(iedTagType)

        Return Me.RtacTagPrototypes(Tag.RootServerTagTypeName).TagPrototypeEntries(TagInfo.Index)
    End Function

    ''' <summary>
    ''' Returns the text after the 2nd dot in an array tag. Returns empty string if type is not an array type.
    ''' </summary>
    ''' <param name="tagInfo">Tag information to get array suffix for.</param>
    ''' <returns>String containing the characters after the 2nd dot in a tag format. Ex: Result of input {SERVER}.BO_{0:D5}.operLatchOn is operLatchOn.</returns>
    Public Function GetArraySuffix(tagInfo As ServerTagInfo) As String
        Dim format = RtacTagPrototypes(tagInfo.RootServerTagTypeName).TagPrototypeEntries(tagInfo.Index).ServerTagNameTemplate

        Dim secondDotIndex = format.GetNthIndex("."c, 2)
        If secondDotIndex < 1 Then Return ""

        Return format.Substring(secondDotIndex, format.Length - secondDotIndex)
    End Function

    ''' <summary>
    ''' Generate a server tag name from a given prototype name template and address.
    ''' </summary>
    ''' <param name="tagPrototypeEntry">Tag prototype entry's format to use.</param>
    ''' <param name="address">Address to substitute in.</param>
    ''' <returns>Formatted server tag name.</returns>
    Public Function GenerateServerTagNameByAddress(tagPrototypeEntry As ServerTagPrototypeEntry, address As Integer) As String
        Dim tagName = tagPrototypeEntry.ServerTagNameTemplate.Replace("{SERVER}", Me.RtacServerName)
        Return String.Format(tagName, address)
    End Function

    ''' <summary>
    ''' Increment generation base address by the given amount.
    ''' </summary>
    ''' <param name="rtacTagName">Server tag type name.</param>
    ''' <param name="incrementVal">Value to increment base address by.</param>
    Public Sub IncrementRtacTagBaseAddressByRtacTagType(rtacTagName As String, incrementVal As Integer)
        Me.TagTypeRunningAddressOffset(rtacTagName) += incrementVal
    End Sub

    ''' <summary>
    ''' Replace standard placeholders in columns.
    ''' </summary>
    ''' <param name="rtacDataRow">Row of data to replace placeholders.</param>
    ''' <param name="rtacTagName">Server tag name.</param>
    ''' <param name="tagAddress">Server tag address.</param>
    ''' <param name="tagAlias">Server tag alias.</param>
    Public Sub ReplaceRtacKeywords(rtacDataRow As OutputRowEntryDictionary, rtacTagName As String, tagAddress As String, tagAlias As String)
        Dim replacements = New Dictionary(Of String, String) From {
            {Keywords.NAME_KEYWORD, rtacTagName},
            {Keywords.ADDRESS_KEYWORD, tagAddress},
            {Keywords.ALIAS_KEYWORD, tagAlias}
        }

        ' Replace keywords
        rtacDataRow.ReplaceTagKeywords(replacements)
    End Sub

    ''' <summary>
    ''' Add a tag row to the output type collection.
    ''' </summary>
    ''' <param name="rtacTagTypeName">Type name to add output to.</param>
    ''' <param name="rtacRow">Data to add.</param>
    Public Sub AddRtacTagOutput(rtacTagTypeName As String, rtacRow As OutputRowEntryDictionary)
        If Not Me._RtacOutputList.ContainsKey(rtacTagTypeName) Then
            Me._RtacOutputList(rtacTagTypeName) = New OutputList
        End If

        Me._RtacOutputList(rtacTagTypeName).Add(rtacRow)
    End Sub

    ''' <summary>
    ''' Return the alias of a server tag given a SCADA name and direction.
    ''' </summary>
    ''' <param name="scadaName">SCADA name to process.</param>
    ''' <param name="pointType">Used to determine if the control suffix needs to be appended.</param>
    ''' <returns>Server tag alias.</returns>
    Public Function GetRtacAlias(scadaName As String, pointType As PointTypeInfo) As String
        If pointType.IsControl Then
            scadaName &= Keywords.CONTROL_KEYWORD
        End If

        For Each s In TagAliasSubstitutes
            scadaName = scadaName.Replace(s.Key, s.Value)
        Next

        Return AliasNameTemplate.Replace(Keywords.NAME_KEYWORD, scadaName)
    End Function

    ''' <summary>
    ''' Validate a tag alias. Throws error if invalid.
    ''' </summary>
    ''' <param name="tagAlias">Tag alias to validate.</param>
    Public Sub ValidateTagAlias(tagAlias As String)
        Dim r = Regex.Match(tagAlias, "^[A-Za-z0-9_]+\s*$", RegexOptions.None)
        If Not r.Success Then Throw New ArgumentException("Invalid tag name: " & tagAlias)
    End Sub

    ''' <summary>
    ''' Write all servers tag types to CSV.
    ''' </summary>
    ''' <param name="path">Source filename to append output suffix on.</param>
    Public Sub WriteAllServerTags(path As String)
        For Each tagGroup In Me._RtacOutputList
            Me.WriteServerTagCSV(tagGroup, path)
        Next
    End Sub

    ''' <summary>
    ''' Write the specified server tag type to CSV.
    ''' </summary>
    ''' <param name="type">Tag type to write out.</param>
    ''' <param name="path">Source filename to append output suffix on.</param>
    Private Sub WriteServerTagCSV(type As KeyValuePair(Of String, OutputList), path As String)
        Dim typeName = type.Key
        Dim tagGroup = type.Value

        Dim comparer As New BySortingColumn(Me.RtacTagPrototypes(typeName).SortingColumn)
        tagGroup.Sort(comparer)

        If Not Me.RtacTagPrototypes.ContainsKey(typeName) Then
            Throw New ArgumentException("Unable to locate tag prototype." & vbCrLf & vbCrLf &
                                "Missing: """ & typeName & """ in tag prototype.")
        End If

        Dim csvPath = IO.Path.GetDirectoryName(path) & IO.Path.DirectorySeparatorChar & IO.Path.GetFileNameWithoutExtension(path) & "_RtacServerTags_" & typeName & ".csv"
        Using csvStreamWriter = New IO.StreamWriter(csvPath, False)
            Dim csvWriter = New CsvHelper.CsvWriter(csvStreamWriter)

            ' Remove address hack from earlier in the generation section
            tagGroup.ForEach(Sub(x) x(Me.RtacTagPrototypes(typeName).SortingColumn) = Fix(x(Me.RtacTagPrototypes(typeName).SortingColumn)))

            For Each tag In tagGroup
                For Each s In OutputRowEntryDictionaryToArray(tag)
                    csvWriter.WriteField(s)
                Next
                csvWriter.NextRecord()
            Next
        End Using
    End Sub

    ''' <summary>
    ''' Parse tag information.
    ''' </summary>
    ''' <remarks>Helper class to parse tag info for array-capable tags</remarks>
    Public Class ServerTagInfo
        Private _RootServerTagTypeName As String
        ''' <summary>
        ''' Root tag type name, such as DNPC
        ''' </summary>
        Public ReadOnly Property RootServerTagTypeName As String
            Get
                Return _RootServerTagTypeName
            End Get
        End Property

        Private _FullServerTagTypeName As String
        ''' <summary>
        ''' Full tag type name, such as DNPC[2]
        ''' </summary>
        Public Property FullServerTagTypeName() As String
            Get
                Return _FullServerTagTypeName
            End Get
            Set(value As String)
                Me.ParseServerTagTypeInfo(value)
            End Set
        End Property

        Private _IsArray As Boolean
        ''' <summary>
        ''' Is tag an array type such as DNPC[2]
        ''' </summary>
        Public ReadOnly Property IsArray As Boolean
            Get
                Return _IsArray
            End Get
        End Property

        Private _Index As Integer
        ''' <summary>
        ''' Index of array tag types such as DNPC[2]
        ''' </summary>
        Public ReadOnly Property Index As Integer
            Get
                Return _Index
            End Get
        End Property

        ''' <summary>
        ''' Initialize a new instance of TagInfo with no tag.
        ''' </summary>
        Public Sub New()

        End Sub

        ''' <summary>
        ''' Initialize a new instance of TagInfo with the given tag name.
        ''' </summary>
        ''' <param name="fullServerTagTypeName">Tag type name to parse</param>
        Public Sub New(fullServerTagTypeName As String)
            ParseServerTagTypeInfo(fullServerTagTypeName)
        End Sub

        ''' <summary>
        ''' Parse given tag type name into root type name and index,
        ''' </summary>
        ''' <param name="fullServerTagTypeName">Tag type name to parse</param>
        Private Sub ParseServerTagTypeInfo(fullServerTagTypeName As String)
            ' Note (?: ) is a non capture group
            Dim r = Regex.Match(fullServerTagTypeName, "(\w+)(?:\[(\d+)\])?", RegexOptions.None)
            If Not r.Success Then Throw New ArgumentException("Invalid tag type name: " & fullServerTagTypeName)

            Me._FullServerTagTypeName = fullServerTagTypeName
            Me._RootServerTagTypeName = r.Groups(1).Value
            Me._IsArray = r.Groups(2).Length > 0
            Me._Index = If(Me.IsArray, CInt(r.Groups(2).Value), 0)
        End Sub
    End Class
End Class