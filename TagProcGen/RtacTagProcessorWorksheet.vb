Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports OutputList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Builds the RTAC tag processor map
''' </summary>
Public Class RtacTagProcessorWorksheet
    ''' <summary>
    ''' Keywords that get replaced with other values.
    ''' </summary>
    Public Class Keywords
        Public Const DESTINATION = "{DESTINATION}"
        Public Const DESTINATION_TYPE = "{DESTINATION_TYPE}"
        Public Const SOURCE = "{SOURCE}"
        Public Const SOURCE_TYPE = "{SOURCE_TYPE}"
        Public Const TIME_SOURCE = "{TIME_SOURCE}"
        Public Const QUALITY_SOURCE = "{QUALITY_SOURCE}"
    End Class

    ''' <summary>
    ''' Quality wrapping mode enumeration.
    ''' </summary>
    Public Enum QualityWrapModeEnum
        ''' <summary>Don't wrap tags with quality check.</summary>
        None = 0
        ''' <summary>Check 1 tag's quality and perform data substitution based on that one tag for the entire device.</summary>
        GroupAllByDevice = 1
        ''' <summary>Individually check the first tag of a device, group the rest of the device's tags into a shared quality check.</summary>
        WrapFirstGroupRestByDevice = 2
        ''' <summary>Individually check each tag's quality and substitute data for that tag only.</summary>
        WrapIndividually = 3
    End Enum

    ''' <summary>Output list of tag processor entries.</summary>
    Private _Map As New List(Of TagProcessorMapEntry)

    ''' <summary>Output list of all tag processor columns.</summary>
    Private _TagProcessorOutputRows As New OutputList

    Private _TagProcessorColumnsTemplate As New OutputRowEntryDictionary
    ''' <summary>Tag processor columns template to generate output column order.</summary>
    Public ReadOnly Property TagProcessorColumnsTemplate As OutputRowEntryDictionary
        Get
            Return _TagProcessorColumnsTemplate
        End Get
    End Property

    ''' <summary>
    ''' Add new entry to the tag processor map.
    ''' </summary>
    ''' <param name="scadaTag">SCADA tag name</param>
    ''' <param name="scadaTagDataType">SCADA tag types</param>
    ''' <param name="iedTagName">Device tag name</param>
    ''' <param name="iedDataType">Device tag type</param>
    ''' <param name="pointType">Is the point status or control and analog or binary.</param>
    ''' <param name="scadaRow">SCADA row entry. Used for calculating nominal values.</param>
    ''' <param name="performQualityWrapping">Indicates wheter to substitute nominal data with the source tag quality is bad.</param>
    ''' <param name="nominalValueColumns">Which columns in the SCADA data should be used to generate nominal values.</param>
    Public Sub AddEntry(
            scadaTag As String, scadaTagDataType As String,
            iedTagName As String, iedDataType As String,
            pointType As PointTypeInfo, scadaRow As OutputRowEntryDictionary,
            performQualityWrapping As Boolean, nominalValueColumns As Tuple(Of Integer, Integer)
    )
        Dim destTag, destType, sourceTag, sourceType As String
        If pointType.IsStatus Then
            ' Flow: IED -> SCADA
            destTag = scadaTag
            destType = scadaTagDataType
            sourceTag = iedTagName
            sourceType = iedDataType
        ElseIf pointType.IsControl Then
            ' Flow: SCADA -> IED
            destTag = iedTagName
            destType = iedDataType
            sourceTag = scadaTag
            sourceType = scadaTagDataType
        Else
            Throw New ArgumentException("Invalid Direction")
        End If

        Me._Map.Add(
            New TagProcessorMapEntry(
                destTag, destType,
                sourceTag, sourceType,
                pointType, scadaRow,
                performQualityWrapping, nominalValueColumns
            )
        )
    End Sub

    ''' <summary>
    ''' The tag processor at RTAC startup sends bad quality data that results in nuisance alarms.
    ''' Wrap tag processor devices in IEC 61131-3 logic to substitute bad quality data with
    ''' placeholder data that will not generate nuisance alarms.
    ''' 
    ''' Data determined as:
    '''  - For status points: SCADA Normal State.
    '''  - For analog points: If alarm limits exist, something that satisfies high / low limits.
    ''' 
    ''' Can write out in multiple ways: By device, by first tag then device group, or by tag.
    ''' By first tag then device group seems to be the best tradeoff betwen granularity and length.
    ''' </summary>
    Private Sub WrapDevicesWithQualitySubstitutions(wrapMode As QualityWrapModeEnum)
        If wrapMode <> QualityWrapModeEnum.None Then
            Dim qualityWrappedMap = New List(Of TagProcessorMapEntry)

            ' Group tag processor into devices
            Dim qualityWrappingGroups = Me._Map.Where(
                Function(mapEntry) mapEntry.PerformQualityWrapping
            ).GroupBy(
                Function(mapEntry) mapEntry.ParsedDeviceName
            ).ToList

            Dim NonQualityWrappingTags = Me._Map.Where(
                Function(mapEntry) Not mapEntry.PerformQualityWrapping
            ).ToList

            If wrapMode = QualityWrapModeEnum.GroupAllByDevice Then
                ' Basic - wrap every device in 1 quality group. Sometimes generates racked in/out alarms.
                ' Seems like the RTAC initializes some points earlier than other points
                For Each deviceGroup In qualityWrappingGroups
                    Dim deviceGroupList = deviceGroup.ToList
                    Dim tagWrapper = New TagQualityWrapGenerator(deviceGroupList)

                    qualityWrappedMap.AddRange(tagWrapper.Generate)
                Next

            ElseIf wrapMode = QualityWrapModeEnum.WrapFirstGroupRestByDevice Then
                ' Write the first tag, then write the remainder as a group. This means at least 2 tags are checked.
                ' Seems to be the best tradeoff between granularity and length in the tag processor.
                For Each deviceGroup In qualityWrappingGroups
                    Dim deviceGroupList = deviceGroup.ToList

                    Dim first As New List(Of TagProcessorMapEntry)
                    first.Add(deviceGroupList.First)
                    Dim tagWrapper = New TagQualityWrapGenerator(first)
                    qualityWrappedMap.AddRange(tagWrapper.Generate)

                    Dim rest = deviceGroupList.Skip(1).ToList
                    If rest.Count > 0 Then
                        tagWrapper = New TagQualityWrapGenerator(rest)
                        qualityWrappedMap.AddRange(tagWrapper.Generate)
                    End If
                Next

            ElseIf wrapMode = QualityWrapModeEnum.WrapIndividually Then
                ' Write every point in its own quality tag. Guaranteed to work at the expense of tag processor length.
                For Each deviceGroup In qualityWrappingGroups
                    Dim deviceGroupList = deviceGroup.ToList

                    For Each tagEntry In deviceGroupList.ToList
                        Dim listOfOne As New List(Of TagProcessorMapEntry)
                        listOfOne.Add(tagEntry)
                        Dim tagWrapper = New TagQualityWrapGenerator(listOfOne)

                        qualityWrappedMap.AddRange(tagWrapper.Generate)
                    Next
                Next
            End If

            Me._Map = qualityWrappedMap ' Update old map with quality wrapped map
            Me._Map.AddRange(NonQualityWrappingTags)
        End If
    End Sub

    ''' <summary>
    ''' Transformer Me._Map into output rows.
    ''' </summary>
    Private Sub GenerateOutputList()
        For Each entry In Me._Map
            Dim outputRowEntry As New OutputRowEntryDictionary(Me.TagProcessorColumnsTemplate)
            ReplaceRtacTagProcessorKeywords(
                    outputRowEntry,
                    entry.DestinationTagName, entry.DestinationTagDataType,
                    entry.SourceExpression, entry.SourceExpressionDataType,
                    entry.TimeSourceTagName, entry.QualitySourceTagName
                )
            Me._TagProcessorOutputRows.Add(outputRowEntry)
        Next
    End Sub

    ''' <summary>
    ''' Replace standard placeholders in columns.
    ''' </summary>
    ''' <param name="rtacTagProcessorRow">Output row to substitute placeholders in.</param>
    ''' <param name="destination">Destination tag.</param>
    ''' <param name="destinationType">Destination tag type.</param>
    ''' <param name="source">Source tag.</param>
    ''' <param name="sourceType">Source tag type.</param>
    ''' <param name="timeSource">Time source tag.</param>
    ''' <param name="qualitySource">Quality source tag.</param>
    Private Sub ReplaceRtacTagProcessorKeywords(
            rtacTagProcessorRow As OutputRowEntryDictionary,
            destination As String,
            destinationType As String,
            source As String,
            sourceType As String,
            timeSource As String,
            qualitySource As String
    )
        Dim replacements = New Dictionary(Of String, String) From {
            {Keywords.DESTINATION, destination},
            {Keywords.DESTINATION_TYPE, destinationType},
            {Keywords.SOURCE, source},
            {Keywords.SOURCE_TYPE, sourceType},
            {Keywords.TIME_SOURCE, timeSource},
            {Keywords.QUALITY_SOURCE, qualitySource}
        }

        ' Replace keywords
        rtacTagProcessorRow.ReplaceTagKeywords(replacements)
    End Sub

    ''' <summary>
    ''' Write the tag processor map out to CSV
    ''' </summary>
    ''' <param name="path">Source filename to append output suffix on.</param>
    ''' <param name="wrapMode">0-3, From no tag wrapping to wrap every tag individually.</param>
    Public Sub WriteCsv(path As String, wrapMode As QualityWrapModeEnum)
        Me.WrapDevicesWithQualitySubstitutions(wrapMode)

        Me.GenerateOutputList()

        Dim csvPath = IO.Path.GetDirectoryName(path) & IO.Path.DirectorySeparatorChar & IO.Path.GetFileNameWithoutExtension(path) & "_TagProcessor.csv"
        Using csvStreamWriter = New IO.StreamWriter(csvPath, False)
            Dim csvWriter As New CsvHelper.CsvWriter(csvStreamWriter)

            For Each c In Me._TagProcessorOutputRows
                For Each s In OutputRowEntryDictionaryToArray(c)
                    csvWriter.WriteField(s)
                Next
                csvWriter.NextRecord()
            Next
        End Using
    End Sub

    ''' <summary>
    ''' Intermediate storage format to generate quality wrapped tag processor entries.
    ''' </summary>
    Public Class TagQualityWrapGenerator
        ''' <summary>Defines the conditional to be used to select bad quality points. Replace {TAG} with tag name.</summary>
        Public Const QUALITY_CONDITIONAL_TEMPLATE = "IF ({TAG}.q.validity <> good) THEN"
        ''' <summary>Else format.</summary>
        Public Const ELSE_TEMPLATE = "ELSE"
        ''' <summary>End if format.</summary>
        Public Const END_IF_TEMPLATE = "END_IF"
        ''' <summary>Time source template. Replace {TAG} with tag name.</summary>
        Public Const TIME_SOURCE_TEMPLATE = "{TAG}.t"
        ''' <summary>Quality source template. Replace {TAG} with tag name.</summary>
        Public Const QUALITY_SOURCE_TEMPLATE = "{TAG}.q"
        ''' <summary>Tag keyword to substitute in the templates.</summary>
        Public Const TAG_KEYWORD = "{TAG}"

        Private _TagsToWrap As List(Of TagProcessorMapEntry)
        ''' <summary>
        ''' List of tags to wrap with quality substitution.
        ''' </summary>
        Public ReadOnly Property TagsToWrap As List(Of TagProcessorMapEntry)
            Get
                Return _TagsToWrap
            End Get
        End Property

        ''' <summary>
        ''' Initialize a new instance of TagQualityWrapGenerator.
        ''' </summary>
        Public Sub New()
            Me._TagsToWrap = New List(Of TagProcessorMapEntry)
        End Sub

        ''' <summary>
        ''' Initialze a new instance of TagQualityWrapGenerator.
        ''' </summary>
        ''' <param name="tagsToWrap">List of all tags to wrap.</param>
        Public Sub New(tagsToWrap As List(Of TagProcessorMapEntry))
            Me._TagsToWrap = New List(Of TagProcessorMapEntry)(tagsToWrap)
        End Sub

        ''' <summary>
        ''' Generate a new list of TagProcessorMapEntry classes that include a quality wrap with nominal values for bad quality tags.
        ''' </summary>
        ''' <returns>List of TagProcessorMapEntry classes with output row information.</returns>
        ''' <remarks>
        ''' Output format:
        '''   If (tag.qual != good) then
        '''     dest = nominal value
        '''   Else
        '''     dest = sourceExpr
        '''   End_if
        ''' </remarks>
        Public Function Generate() As List(Of TagProcessorMapEntry)
            Dim outputTags As New List(Of TagProcessorMapEntry)

            ' Add first conditional for bad quality
            Dim firstTagName = Me._TagsToWrap.First.ParsedTagName
            outputTags.Add(GenerateConditionalTagEntry(QUALITY_CONDITIONAL_TEMPLATE, firstTagName))

            ' Add nominal data
            For Each tag In Me._TagsToWrap
                Dim nominalValue = GetNominalValue(tag.PointType, tag.ScadaRow, tag.NominalValueColumns)
                Dim timeSourceTag = TIME_SOURCE_TEMPLATE.Replace(TAG_KEYWORD, tag.ParsedTagName)
                Dim qualitySourceTag = QUALITY_SOURCE_TEMPLATE.Replace(TAG_KEYWORD, tag.ParsedTagName)

                Dim nominalTag = New TagProcessorMapEntry() With {
                    .DestinationTagName = tag.DestinationTagName,
                    .DestinationTagDataType = tag.DestinationTagDataType,
                    .SourceExpression = nominalValue,
                    .SourceExpressionDataType = "",
                    .TimeSourceTagName = timeSourceTag,
                    .QualitySourceTagName = qualitySourceTag
                }

                outputTags.Add(nominalTag)
            Next

            ' Add else
            outputTags.Add(GenerateConditionalTagEntry(ELSE_TEMPLATE))

            ' Add original data mapping
            outputTags.AddRange(Me._TagsToWrap)

            ' Add end_if
            outputTags.Add(GenerateConditionalTagEntry(END_IF_TEMPLATE))

            Return outputTags
        End Function

        ''' <summary>
        ''' Generate a tag processor entry for a conditional with the given text and optional tag name.
        ''' </summary>
        ''' <param name="conditionalText">Conditional text to put into the source expression.</param>
        ''' <param name="qualityTag">Optional placeholder to substitute.</param>
        ''' <returns>Tag map entry with the given data.</returns>
        Public Shared Function GenerateConditionalTagEntry(conditionalText As String, Optional qualityTag As String = Nothing) As TagProcessorMapEntry
            Dim pointType = New PointTypeInfo(True, True) ' Status binary as a placeholder
            Dim mapEntry As New TagProcessorMapEntry() With {
                .SourceExpression = If(qualityTag = Nothing,
                                       conditionalText,
                                       conditionalText.Replace(TAG_KEYWORD, qualityTag)),
                .PointType = pointType
            }

            Return mapEntry
        End Function

        ''' <summary>
        ''' Get the nominal value of binary or analog status points from provided SCADA data.
        ''' </summary>
        ''' <param name="pointType">Point type information.</param>
        ''' <param name="scadaColumns">SCADA data to derive nominal values from.</param>
        ''' <param name="nominalValueColumns">Where to look in the SCADA data for the nominal values.</param>
        ''' <returns>String that is the normal state for binaries or average of the two median analog alarms limits.</returns>
        Public Shared Function GetNominalValue(
                pointType As PointTypeInfo,
                scadaColumns As OutputRowEntryDictionary,
                nominalValueColumns As Tuple(Of Integer, Integer)
        ) As String
            If pointType.IsBinary Then ' Get binary nominal value
                If scadaColumns.ContainsKey(nominalValueColumns.Item1) Then
                    ' Convert Boolean to IEC 61131-3 TRUE or FALSE
                    Return CBool(scadaColumns(nominalValueColumns.Item1)).ToString.ToUpper
                Else ' Catch all
                    Return "False"
                End If
            Else ' Get analog nominal value
                Dim analogLimits = scadaColumns.Where(
                    Function(x) x.Key >= nominalValueColumns.Item1 And
                        x.Key <= nominalValueColumns.Item2
                ).Where(
                    Function(x) Not String.IsNullOrWhiteSpace(x.Value)
                ).OrderBy(
                    Function(x) CDbl(x.Value)
                ).ToList
                Dim middleStart = CInt(Math.Floor(analogLimits.Count / 2)) - 1

                ' Must have conditional again in case there are no limits defined.
                If middleStart >= 0 Then
                    Dim average = (CDbl(analogLimits(middleStart).Value) + CDbl(analogLimits(middleStart + 1).Value)) / 2
                    Return CStr(average)
                Else ' Catch all
                    Return "0"
                End If
            End If
        End Function
    End Class

    ''' <summary>
    ''' Stores data for each tag processor map entry
    ''' </summary>
    Public Class TagProcessorMapEntry
        ''' <summary>Destination tag name.</summary>
        Public Property DestinationTagName As String
        ''' <summary>Destination tag datat type.</summary>
        Public Property DestinationTagDataType As String
        Private _SourceExpression As String
        ''' <summary>Source expression.</summary>
        Public Property SourceExpression As String
            Get
                Return Me._SourceExpression
            End Get
            Set(value As String)
                Me._SourceExpression = value
                ParseSourceExpression()
            End Set
        End Property
        ''' <summary>Source expression tag type.</summary>
        Public Property SourceExpressionDataType As String
        ''' <summary>Point type (i.e. status, analog).</summary>
        Public Property PointType As PointTypeInfo
        ''' <summary>SCADA data associated with the source tag.</summary>
        ''' <remarks>Used for generating SCADA nominal values that will not generate alarms.</remarks>
        Public Property ScadaRow As OutputRowEntryDictionary
        ''' <summary>Indicates whether to substitute nominal data when the source tag quality is bad.</summary>
        Public Property PerformQualityWrapping As Boolean
        ''' <summary>Column numbers in the SCADA row data that contains the nominal data information.</summary>
        Public Property NominalValueColumns As Tuple(Of Integer, Integer)

        ''' <summary>Time source tag to use when using tag quality substitution.</summary>
        Public Property TimeSourceTagName As String
        ''' <summary>Quality source tag to use when using tag quality substitution.</summary>
        Public Property QualitySourceTagName As String

        Private _ParsedDeviceName As String
        ''' <summary>Device name parsed from the source expression.</summary>
        Public ReadOnly Property ParsedDeviceName As String
            Get
                Return _ParsedDeviceName
            End Get
        End Property

        Private _ParsedTagName As String
        ''' <summary>Tag name parsed from the source expression.</summary>
        Public ReadOnly Property ParsedTagName As String
            Get
                Return _ParsedTagName
            End Get
        End Property

        Private Const RegexMatch As String = ".*?(\p{L}\w*)\.(\p{L}\w*).*"
        Private Const RegexTagName As String = "$1.$2"
        Private Const RegexDeviceName As String = "$1"
        ''' <summary>
        ''' Locates the device name and point name from a source expression.
        ''' </summary>
        Private Sub ParseSourceExpression()
            Me._ParsedDeviceName = Regex.Replace(SourceExpression, RegexMatch, RegexDeviceName)
            Me._ParsedTagName = Regex.Replace(SourceExpression, RegexMatch, RegexTagName)
        End Sub

        ''' <summary>Initialize a new instance of the TagProcessorMapEntry class.</summary>
        Public Sub New()

        End Sub

        ''' <summary>
        ''' Initialize a new instance of the TagProcessorMapEntry class.
        ''' </summary>
        ''' <param name="destinationTagName">Destination tag name.</param>
        ''' <param name="destinationTagDataType">Destination tag data type.</param>
        ''' <param name="sourceExpression">Source expression.</param>
        ''' <param name="sourceExpressionDataType">Source expression datat type.</param>
        ''' <param name="pointType">Point type (i.e. status, analog). Used for filtering entries in the tag processor.</param>
        ''' <param name="scadaRow">SCADA data associtated with source tag.</param>
        ''' <param name="performQualityWrapping">Indicates wheter to substitute nominal data with the source tag quality is bad.</param>
        ''' <param name="nominalValueColumns">Which columns in the SCADA data should be used to generate nominal values.</param>
        Public Sub New(destinationTagName As String, destinationTagDataType As String,
                       sourceExpression As String, sourceExpressionDataType As String,
                       pointType As PointTypeInfo, scadaRow As OutputRowEntryDictionary,
                       performQualityWrapping As Boolean, nominalValueColumns As Tuple(Of Integer, Integer))
            Me.DestinationTagName = destinationTagName
            Me.DestinationTagDataType = destinationTagDataType
            Me.SourceExpression = sourceExpression
            Me.SourceExpressionDataType = sourceExpressionDataType
            Me.PointType = pointType
            Me.ScadaRow = scadaRow
            Me.PerformQualityWrapping = performQualityWrapping
            Me.NominalValueColumns = nominalValueColumns

            Me.ParseSourceExpression()
        End Sub
    End Class
End Class