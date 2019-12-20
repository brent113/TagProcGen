Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports OutputList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Builds the SCADA worksheet and handles tag name formatting and merging
''' </summary>
Public Class ScadaWorksheet
    Private _Pointers As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Key: Pointer Name. Value: Cell Reference
    ''' </summary>
    Public ReadOnly Property Pointers As Dictionary(Of String, String)
        Get
            Return Me._Pointers
        End Get
    End Property

    Private _xlSheet As Excel.Worksheet
    ''' <summary>
    ''' Excel worksheet corresponding to the SCADA template
    ''' </summary>
    Public ReadOnly Property xlSheet As Excel.Worksheet
        Get
            Return Me._xlSheet
        End Get
    End Property

    ''' <summary>
    ''' Create a new instance
    ''' </summary>
    ''' <param name="xlSheet">Excel worksheet corresponding to the SCADA template</param>
    Public Sub New(xlSheet As Excel.Worksheet)
        Me._xlSheet = xlSheet
    End Sub

    ''' <summary>
    ''' Template format to join the IED and Point names into the SCADA name.
    ''' </summary>
    Public Property ScadaNameTemplate As String

    ''' <summary>
    ''' Join IED and Point names to form final SCADA name
    ''' </summary>
    ''' <param name="iedName">SCADA device name</param>
    ''' <param name="pointName">SCADA point name</param>
    ''' <returns>SCADA Name</returns>
    Public Function ScadaNameGenerator(iedName As String, pointName As String) As String
        Return Me.ScadaNameTemplate.Replace(Keywords.IED_NAME_KEYWORD, iedName).Replace(Keywords.POINT_NAME_KEYWORD, pointName)
    End Function

    Private _ScadaTagPrototypes As New Dictionary(Of String, ScadaTagPrototype)()
    ''' <summary>
    ''' Dictionary of all SCADA tag prototypes. Key: SCADA type name, Value: Prototype
    ''' </summary>
    Public ReadOnly Property ScadaTagPrototypes As Dictionary(Of String, ScadaTagPrototype)
        Get
            Return Me._ScadaTagPrototypes
        End Get
    End Property

    ''' <summary>
    ''' SCADA tag prototype containing type-specific data.
    ''' </summary>
    Public Class ScadaTagPrototype
        ''' <summary>
        ''' Standard data all SCADA tags of this type have.
        ''' </summary>
        Public StandardColumns As New OutputRowEntryDictionary

        ''' <summary>
        ''' Format to generate key from address.
        ''' </summary>
        Public KeyFormat As String

        ''' <summary>
        ''' Header row of output CSV.
        ''' </summary>
        Public CsvHeader As String

        ''' <summary>
        ''' Default data equivalent to a new blank record from DataExplorer to merge custom data into.
        ''' </summary>
        Public CsvRowDefaults As String

        ''' <summary>
        ''' Column to sort on
        ''' </summary>
        Public SortingColumn As Integer
    End Class

    ''' <summary>
    ''' Add a new SCADA prototype entry from the given data.
    ''' </summary>
    ''' <param name="pointTypeName">Point type name to add a prototype for.</param>
    ''' <param name="defaultColumnData">Column data all SCADA points of this type have.</param>
    ''' <param name="keyFormat">Format to generate key from address.</param>
    ''' <param name="csvHeader">Header row of output CSV.</param>
    ''' <param name="csvRowDefaults">Default values to use if column data is not specified.</param>
    ''' <param name="sortingColumn">Column to sort output by.</param>
    Public Sub AddTagPrototypeEntry(
            pointTypeName As String,
            defaultColumnData As String,
            keyFormat As String,
            csvHeader As String,
            csvRowDefaults As String,
            sortingColumn As Integer
    )
        Dim pointTypeInfo = New PointTypeInfo(pointTypeName)
        Dim scadaTagPrototype = New ScadaWorksheet.ScadaTagPrototype
        defaultColumnData.ParseColumnDataPairs(scadaTagPrototype.StandardColumns)

        scadaTagPrototype.KeyFormat = keyFormat
        scadaTagPrototype.CsvHeader = csvHeader
        scadaTagPrototype.CsvRowDefaults = csvRowDefaults
        scadaTagPrototype.SortingColumn = sortingColumn

        Me.ScadaTagPrototypes.Add(pointTypeInfo.ToString, scadaTagPrototype)
    End Sub

    ''' <summary>
    ''' Rows of entries used to build the SCADA worksheet. Each row is a column dictionary.
    ''' </summary>
    ''' <remarks>Main output of this template</remarks>
    Private _ScadaOutputList As New Dictionary(Of String, OutputList)

    Private _MaxValidatedTagLength As Integer = 0
    ''' <summary>
    ''' Length of the longest tag name.
    ''' </summary>
    Public ReadOnly Property MaxValidatedTagLength As Integer
        Get
            Return Me._MaxValidatedTagLength
        End Get
    End Property

    Private _MaxValidatedTag As String
    ''' <summary>
    ''' Name of the longest tag.
    ''' </summary>
    Public ReadOnly Property MaxValidatedTag As String
        Get
            Return Me._MaxValidatedTag
        End Get
    End Property

    ''' <summary>
    ''' Keywords that get replaced with other values.
    ''' </summary>
    Public Class Keywords
        Public Const FULL_NAME_KEYWORD = "{NAME}"
        Public Const IED_NAME_KEYWORD = "{DEVICENAME}"
        Public Const POINT_NAME_KEYWORD = "{POINTNAME}"
        Public Const ADDRESS_KEYWORD = "{ADDRESS}"
        Public Const KEY_KEYWORD = "{KEY}"
        Public Const RECORD_KEYWORD = "{RECORD}"
    End Class

    ''' <summary>
    ''' Throws error is tag name is invalid. Letters, numbers and space only; no symbols.
    ''' </summary>
    ''' <param name="tagName">Tag name to validate.</param>
    Public Sub ValidateTagName(tagName As String)
        Dim r = Regex.Match(tagName, "^[A-Za-z0-9 ]+$", RegexOptions.None)
        If Not r.Success Then Throw New ArgumentException("Invalid tag name: " & tagName)

        If tagName.Length > CInt(Me.Pointers(Constants.TPL_SCADA_MAX_NAME_LENGTH)) Then Throw New ArgumentException("Tag name too long: " & tagName)
        If tagName.Length > Me._MaxValidatedTagLength Then
            Me._MaxValidatedTagLength = tagName.Length
            Me._MaxValidatedTag = tagName
        End If
    End Sub

    ''' <summary>
    ''' Substitute SCADA point name and address placeholders with specified values.
    ''' </summary>
    ''' <param name="scadaRowEntry">SCADA entry to find and replace</param>
    ''' <param name="name">Name to substitute into placeholder</param>
    ''' <param name="address">Address to substitute into placeholder. Handles offset here.</param>
    ''' <param name="keyFormat">Format to generate key from address with.</param>
    ''' <param name="keyAddress">Optional parameter to use a different address when generating a key. Do not apply any offset.</param>
    Public Sub ReplaceScadaKeywords(scadaRowEntry As OutputRowEntryDictionary, name As String, address As Integer, keyFormat As String, Optional keyAddress As Integer = -1)
        Dim adjustedAddress = address + CInt(Me.Pointers(Constants.TPL_SCADA_ADDRESS_OFFSET))
        keyAddress = If(keyAddress > 0, keyAddress + CInt(Me.Pointers(Constants.TPL_SCADA_ADDRESS_OFFSET)), adjustedAddress)
        Dim c = String.Format(keyFormat, adjustedAddress)
        Dim replacements = New Dictionary(Of String, String) From {
            {Keywords.FULL_NAME_KEYWORD, name},
            {Keywords.ADDRESS_KEYWORD, adjustedAddress},
            {Keywords.KEY_KEYWORD, String.Format(keyFormat, keyAddress)}
        }

        scadaRowEntry.ReplaceTagKeywords(replacements)
    End Sub

    ''' <summary>
    ''' Add row entry to output.
    ''' </summary>
    ''' <param name="pointTypeInfoName">Name of the point type. </param>
    ''' <param name="scadaRowEntry">Row to add to output.</param>
    Public Sub AddScadaTagOutput(pointTypeInfoName As String, scadaRowEntry As OutputRowEntryDictionary)
        If Not Me._ScadaOutputList.ContainsKey(pointTypeInfoName) Then
            Me._ScadaOutputList(pointTypeInfoName) = New OutputList
        End If

        Me._ScadaOutputList(pointTypeInfoName).Add(scadaRowEntry)
    End Sub

    ''' <summary>
    ''' Write all SCADA tag types to CSV.
    ''' </summary>
    ''' <param name="path">Source filename to append output suffix on.</param>
    Public Sub WriteAllSCADATags(path As String)
        For Each tagGroup In Me._ScadaOutputList
            Me.WriteScadaTagCSV(tagGroup, path)
        Next
    End Sub

    ''' <summary>
    ''' Write the scada worksheet out to CSV.
    ''' </summary>
    ''' <param name="type">Tag type to write out.</param>
    ''' <param name="path">Source filename to append output suffix on.</param>
    Private Sub WriteScadaTagCSV(type As KeyValuePair(Of String, OutputList), path As String)
        Dim typeName = type.Key
        Dim tagGroup = type.Value

        Dim comparer As New BySortingColumn(Me.ScadaTagPrototypes(typeName).SortingColumn)
        tagGroup.Sort(comparer)

        If Not Me.ScadaTagPrototypes.ContainsKey(typeName) Then
            Throw New ArgumentException("Unable to locate tag prototype." & vbCrLf & vbCrLf &
                                "Missing: """ & typeName & """ in tag prototype.")
        End If

        Dim csvPath = IO.Path.GetDirectoryName(path) & IO.Path.DirectorySeparatorChar & IO.Path.GetFileNameWithoutExtension(path) & "_ScadaTags_" & typeName & ".csv"
        Using csvStreamWriter = New IO.StreamWriter(csvPath, False)
            Dim csvWriter As New CsvHelper.CsvWriter(csvStreamWriter)

            ' Write header
            Me.ScadaTagPrototypes(typeName).CsvHeader.Split(","c).ToList.ForEach(Sub(x) csvWriter.WriteField(x))
            csvWriter.NextRecord()

            ' Parse default columns and types to substitute data into
            Dim newRow = Me.ScadaTagPrototypes(typeName).CsvRowDefaults.Split(","c).ToList.Select(
                Function(s)
                    Dim i As Integer
                    Dim r = New With {.Value = s, .isString = Not Integer.TryParse(s, i)}
                    If r.isString Then
                        r.Value = r.Value.Replace(""""c, "")
                    End If
                    Return r
                End Function
            ).ToList

            ' Write out to CSV
            Dim record = 1
            For Each c In tagGroup
                For i = 1 To newRow.Count
                    If c.ContainsKey(i) AndAlso Not String.IsNullOrWhiteSpace(c(i)) Then ' If custom data exists write it.
                        If c(i) = Keywords.RECORD_KEYWORD Then c(i) = record
                        csvWriter.WriteField(c(i), newRow(i - 1).isString)
                    Else ' Otherwise write default data.
                        csvWriter.WriteField(newRow(i - 1).Value, newRow(i - 1).isString)
                    End If
                Next
                csvWriter.NextRecord()
                record += 1
            Next
        End Using
    End Sub
End Class

