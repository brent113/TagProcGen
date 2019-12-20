Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports PointList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

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

    Private _StandardColumns As New OutputRowEntryDictionary
    ''' <summary>
    ''' Base columns of data every SCADA worksheet has.
    ''' </summary>
    Public ReadOnly Property StandardColumns As OutputRowEntryDictionary
        Get
            Return Me._StandardColumns
        End Get
    End Property

    ''' <summary>
    ''' Rows of entries used to build the SCADA worksheet. Each row is a column dictionary.
    ''' </summary>
    ''' <remarks>Main output of this template</remarks>
    Private _ScadaTags As New PointList

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
        Public Const IED_NAME_KEYWORD = "{IED}"
        Public Const POINT_NAME_KEYWORD = "{POINT}"
        Public Const ADDRESS_KEYWORD = "{ADDRESS}"
    End Class

    ''' <summary>
    ''' Throws error is tag name is invalid. Letters, numbers and space only; no symbols.
    ''' </summary>
    ''' <param name="tagName">Tag name to validate.</param>
    Public Sub ValidateTagName(tagName As String)
        Dim r = Regex.Match(tagName, "^[A-Za-z0-9 ]+$", RegexOptions.None)
        If Not r.Success Then Throw New ArgumentException("Invalid tag name: " & tagName)

        If tagName.Length > CInt(Me.Pointers(Constants.TL_SCADA_MAX_NAME_LENGTH)) Then Throw New ArgumentException("Tag name too long: " & tagName)
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
    ''' <param name="address">Address to substitute into placeholder</param>
    Public Shared Sub ReplaceScadaKeywords(scadaRowEntry As OutputRowEntryDictionary, name As String, address As String)
        Dim replacements = New Dictionary(Of String, String) From {
            {Keywords.FULL_NAME_KEYWORD, name},
            {Keywords.ADDRESS_KEYWORD, address}
        }

        scadaRowEntry.ReplaceTagKeywords(replacements)
    End Sub

    ''' <summary>
    ''' Add row entry to output.
    ''' </summary>
    ''' <param name="scadaRowEntry">Row to add to output.</param>
    Public Sub AddScadaTagOutput(scadaRowEntry As OutputRowEntryDictionary)
        Me._ScadaTags.Add(scadaRowEntry)
    End Sub

    ''' <summary>
    ''' Merge input and output tags with the same point name into a single worksheet row, combining column data
    ''' </summary>
    Private Sub MergeScadaTags()
        Dim MergingCol = CInt(Me.Pointers(Constants.TL_SCADA_MERGE_COMPARE_COLUMN))
        Dim DataInCol = CInt(Me.Pointers(Constants.TL_SCADA_MERGE_DATA_IN_COLUMN))
        Dim DataOutCol = CInt(Me.Pointers(Constants.TL_SCADA_MERGE_DATA_OUT_COLUMN))
        Dim MergedScadaTags As New PointList

        ' Loop through each tag entry without an enumerator because
        ' we will be updating the list by removing rows.
        Dim i = 0
        Do While i < Me._ScadaTags.Count
            Dim CurrentScadaRowEntry = Me._ScadaTags(i)

            ' Select tags:
            ' - Add row index into new anonymous type
            ' - Where (typically) searched row's name = current row's name
            Dim TagNameQuery = Me._ScadaTags.
                Select(Function(rowEntry, idx) New With {.Tag = rowEntry, .Index = idx}).
                Where(Function(indexedRowEntry) indexedRowEntry.Tag(MergingCol) = CurrentScadaRowEntry(MergingCol))

            ' Do not allow multiple tags in any given direction
            If TagNameQuery.Count > 2 Then
                ' More than 2 tags implies 2 tags in same direction (only 2 directions possible)
                Throw New Exception(String.Format("Duplicate tag names: {0}", TagNameQuery.First.Tag(MergingCol)))
            ElseIf TagNameQuery.Count > 1 Then
                Dim Item1 = TagNameQuery(0)
                Dim Item2 = TagNameQuery(1)

                ' Get data directions
                Dim Item1DirQuery = {DataInCol, DataOutCol}.ToList.Where(Function(x) Item1.Tag.ContainsKey(x))
                If Item1DirQuery.Count <> 1 Then Throw New Exception("Invalid number of items returned while merging.")
                Dim Item1DirIn = Item1DirQuery(0) = DataInCol

                Dim Item2DirQuery = {DataInCol, DataOutCol}.ToList.Where(Function(x) Item2.Tag.ContainsKey(x))
                If Item2DirQuery.Count <> 1 Then Throw New Exception("Invalid number of items returned while merging.")
                Dim Item2DirIn = Item2DirQuery(0) = DataInCol

                ' Same direction? throw error
                If Item1DirIn = Item2DirIn Then Throw New Exception(String.Format("Duplicate item names with the same data direction: {0}", Item1.Tag(MergingCol)))

                ' Copy both items into merged columns
                Dim Merged As New OutputRowEntryDictionary(Item1.Tag)
                For Each kv In Item2.Tag
                    Merged(kv.Key) = kv.Value
                Next
                MergedScadaTags.Add(Merged)

                ' Remove highest index first because it doesn't affect low index
                Me._ScadaTags.RemoveAt(Math.Max(Item1.Index, Item2.Index))
                Me._ScadaTags.RemoveAt(Math.Min(Item1.Index, Item2.Index))
                ' Don't add 1 to iterator
            Else
                ' Copy to output
                MergedScadaTags.Add(CurrentScadaRowEntry)
                i += 1
            End If
        Loop

        Me._ScadaTags = MergedScadaTags
    End Sub

    ''' <summary>
    ''' Write the scada worksheet out to CSV
    ''' </summary>
    ''' <param name="Path">Source filename to append output suffix on.</param>
    Public Sub WriteCsv(Path As String)
        Me.MergeScadaTags()

        Dim csvPath = IO.Path.GetDirectoryName(Path) & IO.Path.DirectorySeparatorChar & IO.Path.GetFileNameWithoutExtension(Path) & "_SCADA_Worksheet.csv"
        Using csvStreamWriter = New IO.StreamWriter(csvPath, False)
            Dim CsvWriter As New CsvHelper.CsvWriter(csvStreamWriter)

            For Each c In Me._ScadaTags
                For Each s In OutputRowEntryDictionaryToArray(c)
                    CsvWriter.WriteField(s)
                Next
                CsvWriter.NextRecord()
            Next
        End Using
    End Sub
End Class

