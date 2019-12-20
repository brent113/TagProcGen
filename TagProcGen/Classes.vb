Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports ColumnDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports PointList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

Public Class RtacTagMap
    Public Pointers As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Public xlSheet As Excel.Worksheet
    Public Map As New List(Of CustomMapEntry)

    Public Class CustomMapEntry
        Public DestinationTagName As String
        Public DTDataType As String
        Public SourceExpression As String
        Public SEDataType As String
    End Class
End Class

Public Class BySortingColumn
    Implements IComparer(Of ColumnDictionary)
    Private m_sortingColumn As Integer

    Public Sub New(sortingColumn As Integer)
        m_sortingColumn = sortingColumn
    End Sub

    Public Function Compare(x As ColumnDictionary, y As ColumnDictionary) As Integer Implements IComparer(Of ColumnDictionary).Compare
        Dim xVal = CDbl(x(m_sortingColumn))
        Dim YVal = CDbl(y(m_sortingColumn))

        Return xVal.CompareTo(YVal)
    End Function
End Class

Public Module ExtensionMethods
    <System.Runtime.CompilerServices.Extension>
    Public Sub ParseColumnDataPairs(columnDataPairString As String, ByRef columnDataDict As ColumnDictionary)
        ' Split col / data pairs - example format: [1, True];[2, {NAME}]
        For Each colPair In DirectCast(columnDataPairString, String).Split(";"c)
            If colPair.Length = 0 Then Throw New Exception("Malformed column / data pair: " & columnDataPairString)
            ' strip [ and ]
            If colPair(0) <> "["c Or colPair(colPair.Length - 1) <> "]"c Then Throw New Exception("Malformed column / data pair: " & colPair)
            Dim t = colPair.Substring(1, colPair.Length - 2).Split(","c)

            Dim colIndex As Integer
            If Not Integer.TryParse(t(0).Trim, colIndex) Then
                Throw New Exception("Invalid Column Index: unable to convert """ & t(0).Trim & """ to an integer")
            End If

            Dim colData = t(1).Trim

            columnDataDict.Add(colIndex, colData)
        Next
    End Sub

    ''' <summary>
    ''' Apply replacements to column keywords like {NAME} and {ADDRESS}
    ''' </summary>
    ''' <param name="Columns">Column data pair dictionary to update</param>
    ''' <param name="Replacements">Dictionary of keywords (like {NAME}) and their replacement</param>
    <System.Runtime.CompilerServices.Extension>
    Public Sub ReplaceTagKeywords(ByRef columns As ColumnDictionary, ByRef replacements As Dictionary(Of String, String))
        Dim copy = New ColumnDictionary(columns)
        For Each col In columns
            For Each rep In replacements
                copy(col.Key) = copy(col.Key).Replace(rep.Key, rep.Value)
            Next
        Next
        columns = copy
    End Sub

    <System.Runtime.CompilerServices.Extension>
    Public Function GetNthIndex(s As String, t As Char, n As Integer) As Integer
        Dim count = 0
        For i = 0 To s.Length - 1
            If s(i) = t Then
                count += 1
                If count = n Then
                    Return i
                End If
            End If
        Next
        Return -1
    End Function
End Module