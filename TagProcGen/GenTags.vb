Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports OutputList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Main class that orchestrates and does the tag generation.
''' </summary>
Public Module GenTags

    ''' <summary>
    ''' List of global template reference lookup pairs
    ''' </summary>
    Public GlobalPointers As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    ' Worksheet References.
    ''' <summary>Global Excel application reference.</summary>
    Public xlApp As Excel.Application
    ''' <summary>Global Excel Workbook reference.</summary>
    Public xlWorkbook As Excel.Workbook
    ''' <summary>Global template reference to definitions worksheet.</summary>
    Public xlDef As Excel.Worksheet

    ' Data templates for storing data.
    ''' <summary>RTAC server tags template.</summary>
    Public TPL_Rtac As RtacTemplate
    ''' <summary>Every loaded device template.</summary>
    Public IedTemplates As List(Of IedTemplate)

    ' Output worksheet generators.
    ''' <summary>RTAC tag processor worksheet template</summary>
    Public TPL_TagProcessor As RtacTagProcessorWorksheet
    ''' <summary>SCADA worksheet template.</summary>
    Public TPL_Scada As ScadaWorksheet

    ''' <summary>
    ''' Master function that orchestrates the generation process. Calls each responsible function in turn.
    ''' </summary>
    ''' <param name="path">Path to the Excel workbook containing the configuration.</param>
    Public Sub Generate(path As String)
        Dim Processing As String = ""
        Try
            Processing = "Initializing" : InitExcel(path)
            Processing = "Loading Templates" : LocateTemplates()
            Processing = "Loading Pointers" : LoadPointers()
            Processing = "Reading RTAC Template" : ReadRtac()
            Processing = "Reading SCADA" : ReadScada()

            For Each t In IedTemplates
                Processing = "Reading Template " & t.xlSheet.Name
                ReadTemplate(t)
                t.Validate(TPL_Rtac)
            Next

            For Each t In IedTemplates
                Processing = "Processing Template " & t.xlSheet.Name
                GenIEDTagProcMap(t)
            Next

            Processing = "Writing Tag Map"
            TPL_TagProcessor.WriteCsv(path, TPL_Rtac.Pointers(Constants.TPL_RTAC_TAG_PROC_WRAP_MODE))

            Processing = "Writing SCADA Tags"
            TPL_Scada.WriteAllSCADATags(path)

            Processing = "Writing RTAC Tags"
            TPL_Rtac.WriteAllServerTags(path)
        Catch ex As Exception
            MsgBox("Could not successfully generate tag map. Error text:" &
                   vbCrLf & vbCrLf & ex.Message &
                   vbCrLf & vbCrLf & "Occured while: " & Processing, MsgBoxStyle.Critical, "Error")
            Return
        Finally
            xlWorkbook.Close(False)
        End Try

        MsgBox("Successfully generated tag processor map." & vbCrLf & vbCrLf &
               "Longest SCADA tag name: " & TPL_Scada.MaxValidatedTag & " at " & TPL_Scada.MaxValidatedTagLength & " characters.",
               MsgBoxStyle.Information, "Success")
    End Sub

    ''' <summary>
    ''' Initialize an instance of Excel and load the workbook specified.
    ''' </summary>
    ''' <param name="Path">Path to the Excel workbook containing the configuration.</param>
    Public Sub InitExcel(Path As String)
        xlApp = New Excel.Application
        xlWorkbook = xlApp.Workbooks.Open(Path)
    End Sub

    ''' <summary>
    ''' Locate and load the worksheet templates in the workbook that need processing.
    ''' </summary>
    Public Sub LocateTemplates()
        xlDef = xlWorkbook.Sheets(Constants.TPL_DEF_SHEET)

        TPL_TagProcessor = New RtacTagProcessorWorksheet

        TPL_Rtac = New RtacTemplate(xlWorkbook.Sheets(Constants.TPL_RTAC_SHEET))

        TPL_Scada = New ScadaWorksheet(xlWorkbook.Sheets(Constants.TPL_SCADA_SHEET))

        IedTemplates = New List(Of IedTemplate)
        Dim specialSheets = {Constants.TPL_DEF_SHEET, Constants.TPL_RTAC_SHEET, Constants.TPL_SCADA_SHEET}
        For Each sht As Excel.Worksheet In xlWorkbook.Sheets
            If sht.Name.StartsWith(Constants.TPL_SHEET_PREFIX) And Not specialSheets.Contains(sht.Name) Then
                IedTemplates.Add(New IedTemplate(sht))
            End If
        Next
    End Sub

    ''' <summary>
    ''' Read the worksheet pointers from each template.
    ''' </summary>
    Public Sub LoadPointers()
        ' Definition sheet pointers
        ReadPairRange(xlDef.Range(Constants.TPL_DEF), GlobalPointers,
                      Constants.TPL_RTAC_DEF,
                      Constants.TPL_SCADA_DEF,
                      Constants.TPL_IED_DEF)

        ' RTAC sheet pointers
        ReadPairRange(TPL_Rtac.xlSheet.Range(GlobalPointers(Constants.TPL_RTAC_DEF)), TPL_Rtac.Pointers,
                      Constants.TPL_RTAC_MAP_NAME,
                      Constants.TPL_RTAC_TAG_PROTO,
                      Constants.TPL_RTAC_TAG_MAP,
                      Constants.TPL_RTAC_ALIAS_SUB,
                      Constants.TPL_RTAC_TAG_PROC_COLS,
                      Constants.TPL_RTAC_TAG_PROC_WRAP_MODE)

        ' SCADA sheet pointers
        ReadPairRange(TPL_Scada.xlSheet.Range(GlobalPointers(Constants.TPL_SCADA_DEF)), TPL_Scada.Pointers,
                      Constants.TPL_SCADA_NAME_FORMAT,
                      Constants.TPL_SCADA_MAX_NAME_LENGTH,
                      Constants.TPL_SCADA_TAG_PROTO,
                      Constants.TPL_SCADA_ADDRESS_OFFSET)

        ' IED pointers
        For Each t In IedTemplates
            ' Pointers
            ReadPairRange(t.xlSheet.Range(GlobalPointers(Constants.TPL_IED_DEF)), t.Pointers,
                          Constants.TPL_DATA,
                          Constants.TPL_IED_NAMES,
                          Constants.TPL_OFFSETS)
        Next
    End Sub

    ''' <summary>
    ''' Read the RTAC template data.
    ''' </summary>
    Public Sub ReadRtac()
        Dim c As Excel.Range
        With TPL_Rtac
            ' Read name
            .RtacServerName = .xlSheet.Range(.Pointers(Constants.TPL_RTAC_MAP_NAME)).Text
            .AliasNameTemplate = .xlSheet.Range(.Pointers(Constants.TPL_RTAC_MAP_NAME)).Offset(0, 1).Text

            ' Read tag prototypes, splitting when necessary
            c = .xlSheet.Range(.Pointers(Constants.TPL_RTAC_TAG_PROTO))
            While Not String.IsNullOrEmpty(c.Text)
                Dim tag As New RtacTemplate.ServerTagInfo(c.Text)
                Dim prototypeFormat As String = c.Offset(0, 1).Text
                Dim colDataPairString As String = c.Offset(0, 2).Text
                Dim sortingColumnRaw As String = c.Offset(0, 3).Text
                Dim pointTypeInfoText As String = c.Offset(0, 4).Text
                Dim analogLimitColumnRange As String = c.Offset(0, 5).Text

                Dim sortingColumn = -1
                If Not Integer.TryParse(sortingColumnRaw, sortingColumn) Then sortingColumn = -1

                .AddTagPrototypeEntry(
                    tag,
                    prototypeFormat,
                    colDataPairString,
                    sortingColumn,
                    pointTypeInfoText,
                    analogLimitColumnRange
                )

                c = c.Offset(1, 0)
            End While
            TPL_Rtac.ValidateTagPrototypes()

            ' Read tag type map
            c = .xlSheet.Range(.Pointers(Constants.TPL_RTAC_TAG_MAP))
            While Not String.IsNullOrEmpty(c.Text)
                Dim iedType As String = c.Text
                Dim rtacType As String = c.Offset(0, 1).Text
                Dim performQualityMappingRaw As String = c.Offset(0, 2).Text

                Dim performQualityMapping As Boolean
                Dim parseSuccess = Boolean.TryParse(performQualityMappingRaw, performQualityMapping)
                If Not parseSuccess Then
                    Throw New Exception(String.Format("Invalid quality wrapping flag for IED Type map entry {0}", iedType))
                End If

                .AddIedServerTagMap(iedType, rtacType, performQualityMapping)

                c = c.Offset(1, 0)
            End While

            ' Read Tag Alias Substitutions
            ReadPairRange(.xlSheet.Range(.Pointers(Constants.TPL_RTAC_ALIAS_SUB)), .TagAliasSubstitutes)

            ' Read Tag processor Columns
            DirectCast(.xlSheet.Range(.Pointers(Constants.TPL_RTAC_TAG_PROC_COLS)).Text, String).ParseColumnDataPairs(TPL_TagProcessor.TagProcessorColumnsTemplate)
        End With
    End Sub

    ''' <summary>
    ''' Read the SCADA template data.
    ''' </summary>
    Public Sub ReadScada()
        With TPL_Scada
            ' Read name format
            Dim c = .xlSheet.Range(TPL_Scada.Pointers(Constants.TPL_SCADA_NAME_FORMAT))
            .ScadaNameTemplate = c.Text

            ' Read SCADA prototypes
            c = .xlSheet.Range(.Pointers(Constants.TPL_SCADA_TAG_PROTO))
            While Not String.IsNullOrEmpty(c.Text)
                Dim pointTypeName As String = c.Text
                Dim defaultColumnData As String = c.Offset(0, 1).Text
                Dim keyFormat As String = c.Offset(0, 2).Text
                Dim csvHeader As String = c.Offset(0, 3).Text
                Dim csvRowDefaults As String = c.Offset(0, 4).Text
                Dim sortingColumnRaw As String = c.Offset(0, 5).Text

                Dim sortingColumn = -1
                If Not Integer.TryParse(sortingColumnRaw, sortingColumn) Then sortingColumn = -1

                If sortingColumn < 0 Then Throw New Exception(String.Format("SCADA prototype {0} is missing a valid sorting column.", pointTypeName))

                .AddTagPrototypeEntry(pointTypeName,
                                      defaultColumnData,
                                      keyFormat,
                                      csvHeader,
                                      csvRowDefaults,
                                      sortingColumn)

                c = c.Offset(1, 0)
            End While
        End With
    End Sub

    ''' <summary>
    ''' Read the specified device template data.
    ''' </summary>
    ''' <param name="t">Device template to read.</param>
    Public Sub ReadTemplate(t As IedTemplate)
        Dim c As Excel.Range

        ' Read offsets
        ReadPairRange(t.xlSheet.Range(t.Pointers(Constants.TPL_OFFSETS)), t.Offsets)

        ' Read IED and SCADA names
        c = t.xlSheet.Range(t.Pointers(Constants.TPL_IED_NAMES))
        While Not String.IsNullOrEmpty(c.Text)
            t.IedScadaNames.Add(New IedTemplate.IedScadaNamePair With
                                {
                                    .IedName = c.Text,
                                    .ScadaName = c.Offset(0, 1).Text
                                })

            c = c.Offset(1, 0)
        End While

        ' Read tag data
        c = t.xlSheet.Range(t.Pointers(Constants.TPL_DATA))
        ' for speed locate the last row, then do 1 large read
        While Not String.IsNullOrEmpty(c.Offset(10, 0).Text) ' read by 10s
            c = c.Offset(10, 0)
        End While
        While Not String.IsNullOrEmpty(c.Offset(1, 0).Text) ' read by 1s
            c = c.Offset(1, 0)
        End While
        Dim dataTable = t.xlSheet.Range(t.xlSheet.Range(t.Pointers(Constants.TPL_DATA)).Address & ":" &
                                        c.Offset(0, 7).Address).Value2

        For i As Integer = 1 To dataTable.GetLength(0)
            Dim Process = dataTable(i, 1).ToString.ToUpper = "TRUE"
            If Process Then
                Dim filterRaw As String = If(dataTable(i, 2) IsNot Nothing, dataTable(i, 2).ToString, "")
                Dim pointNumberRaw As String = If(dataTable(i, 3) IsNot Nothing, dataTable(i, 3).ToString, "")
                Dim iedTagName As String = If(dataTable(i, 4) IsNot Nothing, dataTable(i, 4).ToString, "")
                Dim iedTagType As String = If(dataTable(i, 5) IsNot Nothing, dataTable(i, 5).ToString, "")
                Dim rtacColumns As String = If(dataTable(i, 6) IsNot Nothing, dataTable(i, 6).ToString, "")
                Dim scadaPointName As String = If(dataTable(i, 7) IsNot Nothing, dataTable(i, 7).ToString, "")
                Dim scadaColumns As String = If(dataTable(i, 8) IsNot Nothing, dataTable(i, 8).ToString, "")

                Dim pointNumber As Integer
                Dim pointNumberIsAbsolute As Boolean
                If pointNumberRaw.Length = 0 Then
                    Throw New Exception("Point number missing")
                End If
                pointNumberIsAbsolute = pointNumberRaw.Substring(0, 1) = "="
                pointNumber = CInt(If(pointNumberIsAbsolute, pointNumberRaw.Substring(1), pointNumberRaw))

                Dim filter = New IedTemplate.FilterInfo(filterRaw)
                Dim dataEntry = t.GetOrCreateTagEntry(iedTagType, filter, pointNumber, TPL_Rtac)
                With dataEntry
                    .DeviceFilter = New IedTemplate.FilterInfo(filterRaw)
                    .PointNumber = pointNumber
                    .PointNumberIsAbsolute = pointNumberIsAbsolute
                    .IedTagNameTypeList.Add(New IedTemplate.IedTagNameTypePair With {
                                        .IedTagName = iedTagName,
                                        .IedTagTypeName = iedTagType})

                    If rtacColumns.Length > 0 Then rtacColumns.ParseColumnDataPairs(.RtacColumns)
                    If scadaPointName.Length > 0 Then .ScadaPointName = scadaPointName
                    If scadaColumns.Length > 0 Then scadaColumns.ParseColumnDataPairs(.ScadaColumns)
                End With
            End If
        Next
    End Sub

    ''' <summary>
    ''' This function does a few things:
    '''   Generate SCADA output rows
    '''   Generate RTAC output rows
    '''   Generate Tag Map
    ''' </summary>
    ''' <param name="t">Template to generate data for.</param>
    Public Sub GenIEDTagProcMap(t As IedTemplate)
        For Each iedScadaNamePair In t.IedScadaNames
            ' Generate Data tag map and server tags from IEDs
            For Each tag In t.AllIedPoints
                ' Skip tag generation if this tag is filtered out
                If Not tag.DeviceFilter.ShouldPointBeGenerated(iedScadaNamePair.IedName) Then Continue For

                ' Begin calc in advance
                ' the address and alias. Calc the name for each format entry in the loop for each format

                ' Lookup RTAC tag info and prototype for later
                Dim rtacTagInfoRootName = TPL_Rtac.GetServerTagInfoByDevice(tag.IedTagNameTypeList.First.IedTagTypeName).RootServerTagTypeName
                Dim newTagRootPrototype = TPL_Rtac.RtacTagPrototypes(rtacTagInfoRootName)
                Dim addressBase = TPL_Rtac.TagTypeRunningAddressOffset(rtacTagInfoRootName)

                ' Calc in advance some basic info
                Dim address As Integer = IIf(tag.PointNumberIsAbsolute, tag.PointNumber, addressBase + tag.PointNumber)
                Dim ProcessScada = tag.ScadaPointName <> "--"

                Dim scadaFullName = "", rtacAlias = ""
                If ProcessScada Then
                    scadaFullName = TPL_Scada.ScadaNameGenerator(iedScadaNamePair.ScadaName, tag.ScadaPointName)

                    TPL_Scada.ValidateTagName(scadaFullName)

                    rtacAlias = TPL_Rtac.GetRtacAlias(scadaFullName, newTagRootPrototype.PointType)
                    TPL_Rtac.ValidateTagAlias(rtacAlias)
                End If
                ' End calc in advance

                Dim scadaTagPrototype = TPL_Scada.ScadaTagPrototypes(newTagRootPrototype.PointType.ToString)
                Dim scadaColumns As New OutputRowEntryDictionary(scadaTagPrototype.StandardColumns) ' Default SCADA columns
                If ProcessScada Then
                    ' Begin SCADA column processing
                    Try
                        tag.ScadaColumns.ToList.ForEach(Sub(c) scadaColumns.Add(c.Key, c.Value)) ' Custom SCADA columns
                    Catch e As Exception
                        Throw New Exception("Invalid SCADA column definitions - duplicate columns present.")
                    End Try

                    ' todo: replace with linked address if control
                    If newTagRootPrototype.PointType.IsStatus Then
                        ' Replace keywords and add SCADA columns to output
                        TPL_Scada.ReplaceScadaKeywords(scadaColumns, scadaFullName, address, scadaTagPrototype.KeyFormat)
                    Else
                        ' Replace keywords. Specify separate key link address based on linked status point
                        Dim linkedStatusPoint = t.GetLinkedStatusPoint(iedScadaNamePair.IedName, tag.ScadaPointName, TPL_Rtac)
                        Dim linkedTagRootPrototype = TPL_Rtac.GetServerTagInfoByDevice(linkedStatusPoint.IedTagNameTypeList.First.IedTagTypeName).RootServerTagTypeName
                        Dim linkedAddressBase = TPL_Rtac.TagTypeRunningAddressOffset(linkedTagRootPrototype)
                        Dim linkedAddress As Integer = IIf(linkedStatusPoint.PointNumberIsAbsolute, linkedStatusPoint.PointNumber, linkedAddressBase + linkedStatusPoint.PointNumber)

                        TPL_Scada.ReplaceScadaKeywords(scadaColumns, scadaFullName, address, scadaTagPrototype.KeyFormat, linkedAddress)
                    End If


                    TPL_Scada.AddScadaTagOutput(newTagRootPrototype.PointType.ToString, scadaColumns)
                    ' SCADA columns done
                End If

                ' Begin RTAC column processing
                For index = 0 To newTagRootPrototype.TagPrototypeEntries.Count - 1
                    Dim rtacColumns As New OutputRowEntryDictionary(newTagRootPrototype.TagPrototypeEntries(index).StandardColumns) ' Default RTAC Columns

                    Try
                        tag.RtacColumns.ToList.ForEach(Sub(c) rtacColumns.Add(c.Key, c.Value)) ' Custom RTAC columns
                    Catch e As Exception
                        Throw New Exception("Invalid RTAC column definitions - duplicate columns present.")
                    End Try

                    ' Point name from format
                    Dim tagName = TPL_Rtac.GenerateServerTagNameByAddress(newTagRootPrototype.TagPrototypeEntries(index), address)

                    If ProcessScada Then
                        ' Begin tag map processing
                        ' Check if there's an IED tag that maps to the current tag prototype
                        Dim idx = index ' Required because iteration variables cannot be used in queries
                        Dim iedTag = (From ied In tag.IedTagNameTypeList
                                      Where (TPL_Rtac.GetServerTagInfoByDevice(ied.IedTagTypeName).Index = idx)
                                      ).ToList

                        If iedTag.Count > 1 Then Throw New Exception("Too many tags that map to " & rtacTagInfoRootName & ". Tag = " & iedTag.First.IedTagName)

                        If iedTag.Count = 1 Then
                            Dim iedTagName = IedTemplate.SubstituteTagName(iedTag(0).IedTagName, iedScadaNamePair.IedName)
                            Dim iedTagTypeName = iedTag(0).IedTagTypeName

                            Dim rtacTagInfo = TPL_Rtac.GetServerTagInfoByDevice(iedTagTypeName)
                            Dim rtacTagSuffix = TPL_Rtac.GetArraySuffix(rtacTagInfo)

                            Dim rtacServerTagTypeMap = TPL_Rtac.GetServerTypeByIedType(iedTagTypeName)

                            Dim rtacServerTagName = "Tags." & rtacAlias & rtacTagSuffix
                            Dim rtacServerTagType = rtacServerTagTypeMap.ServerTagTypeName

                            TPL_TagProcessor.AddEntry(
                                rtacServerTagName, rtacServerTagType,
                                iedTagName, iedTagTypeName,
                                newTagRootPrototype.PointType, scadaColumns,
                                rtacServerTagTypeMap.PerformQualityWrapping, newTagRootPrototype.NominalColumns
                            )
                        End If
                        ' Tag map processing done

                        ' Calculate address fractional addition below to maintain sort order later 
                        ' for when potentially duplicate addresses get sorted, ie: array type
                        Dim fractionalAddress = CDbl(index) / CDbl(newTagRootPrototype.TagPrototypeEntries.Count)

                        TPL_Rtac.ReplaceRtacKeywords(rtacColumns, tagName, CDbl(address) + fractionalAddress, rtacAlias)
                        TPL_Rtac.AddRtacTagOutput(rtacTagInfoRootName, rtacColumns)
                    End If
                Next
                ' RTAC columns done
            Next

            ' Increment Server tag starting value by type offsets
            For Each offset In t.Offsets
                TPL_Rtac.IncrementRtacTagBaseAddressByRtacTagType(offset.Key, CInt(offset.Value))
            Next
        Next
    End Sub
End Module