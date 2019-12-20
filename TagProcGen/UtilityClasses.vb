Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports OutputList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Global constants.
''' </summary>
Public Class Constants
    ' Initial Pointer
    Public Const TPL_DEF = "A3"

    ' Important Excel worksheet names
    Public Const TPL_SHEET_PREFIX = "TPL_"
    Public Const TPL_DEF_SHEET = "TPL_Def"
    Public Const TPL_RTAC_SHEET = "TPL_RTAC"
    Public Const TPL_SCADA_SHEET = "TPL_SCADA"

    ' Global Template Reference Constants
    Public Const TPL_RTAC_DEF = "TPL_RTAC_DEF"
    Public Const TPL_SCADA_DEF = "TPL_SCADA_DEF"
    Public Const TPL_IED_DEF = "TPL_IED_DEF"

    ' Custom Map Constants
    Public Const TPL_TAG_MAP = "TPL_TAG_MAP"
    Public Const TPL_RTAC_TAGS = "TPL_RTAC_TAGS"

    ' RTAC Template Reference Constants
    Public Const TPL_RTAC_TAG_PROTO = "TPL_RTAC_TAG_PROTO"
    Public Const TPL_RTAC_MAP_NAME = "TPL_RTAC_MAP_NAME"
    Public Const TPL_RTAC_TAG_MAP = "TPL_RTAC_TAG_MAP"
    Public Const TPL_RTAC_ALIAS_SUB = "TPL_RTAC_ALIAS_SUB"
    Public Const TPL_RTAC_TAG_PROC_COLS = "TPL_RTAC_TAG_PROC_COLS"
    Public Const TPL_RTAC_TAG_PROC_WRAP_MODE = "TPL_RTAC_TAG_PROC_WRAP_MODE"

    ' SCADA Template Reference Constants
    Public Const TPL_SCADA_NAME_FORMAT = "TPL_SCADA_NAME_FORMAT"
    Public Const TPL_SCADA_MAX_NAME_LENGTH = "TPL_SCADA_MAX_NAME_LENGTH"
    Public Const TPL_SCADA_TAG_PROTO = "TPL_SCADA_TAG_PROTO"
    Public Const TPL_SCADA_ADDRESS_OFFSET = "TPL_SCADA_ADDRESS_OFFSET"

    ' IED Template Reference Constants
    Public Const TPL_DATA = "TPL_DATA"
    Public Const TPL_IED_NAMES = "TPL_IED_NAMES"
    Public Const TPL_OFFSETS = "TPL_OFFSETS"

    ' Point Type Constants
    Public Const STATUS_BINARY = "STATUSBINARY"
    Public Const STATUS_ANALOG = "STATUSANALOG"
    Public Const CONTROL_BINARY = "CONTROLBINARY"
    Public Const CONTROL_ANALOG = "CONTROLANALOG"
End Class

''' <summary>
''' Represents points that are analog or binary, both status and control.
''' </summary>
Public Class PointTypeInfo
    Private _pointTypeName As String

    Private _IsStatus As Boolean
    ''' <summary>Returns True if the point is a status type.</summary>
    Public ReadOnly Property IsStatus As Boolean
        Get
            Return _IsStatus
        End Get
    End Property
    ''' <summary>Returns True if the point is a control type.</summary>
    Public ReadOnly Property IsControl As Boolean
        Get
            Return Not _IsStatus
        End Get
    End Property

    Private _IsBinary As Boolean
    ''' <summary>Returns True if the point is a binary type.</summary>
    Public ReadOnly Property IsBinary As Boolean
        Get
            Return _IsBinary
        End Get
    End Property
    ''' <summary>Returns True if the point is an analog type.</summary>
    Public ReadOnly Property IsAnalog As Boolean
        Get
            Return Not _IsBinary
        End Get
    End Property

    ''' <summary>
    ''' Initialize a new instance of PointTypeInfo from text.
    ''' </summary>
    ''' <param name="pointTypeText">Text to parse point type information for.</param>
    ''' <remarks>Text should be like: StatusAnalog, or ControlBinary</remarks>
    Public Sub New(pointTypeText As String)
        Me._pointTypeName = pointTypeText.ToUpper

        Dim AllTypes = {
                Constants.STATUS_BINARY, Constants.STATUS_ANALOG,
                Constants.CONTROL_BINARY, Constants.CONTROL_ANALOG
        }
        Dim StatusTypes = {Constants.STATUS_BINARY, Constants.STATUS_ANALOG}
        Dim BinaryTypes = {Constants.STATUS_BINARY, Constants.CONTROL_BINARY}

        If Not AllTypes.Contains(_pointTypeName) Then
            Throw New Exception(String.Format("Point type {0} is not a valid point type.", _pointTypeName))
        End If

        Me._IsStatus = StatusTypes.Contains(_pointTypeName)
        Me._IsBinary = BinaryTypes.Contains(_pointTypeName)
    End Sub

    ''' <summary>
    ''' Initialize a new instance of PointTypeInfo from values.
    ''' </summary>
    ''' <param name="isStatus">Indicates the point is a status point.</param>
    ''' <param name="isBinary">Indicates the point is a binary point.</param>
    Public Sub New(isStatus As Boolean, isBinary As Boolean)
        Me._IsStatus = isStatus
        Me._IsBinary = isBinary
    End Sub

    ''' <summary>
    ''' Returns a string that represents the current object.
    ''' </summary>
    ''' <returns>The point type name as a string.</returns>
    Public Overrides Function ToString() As String
        Return Me._pointTypeName
    End Function
End Class

''' <summary>
''' Custom comparer that sorts a list of output rows by the given sorting column alphanumerically.
''' </summary>
Public Class BySortingColumn
    Implements IComparer(Of OutputRowEntryDictionary)
    Private m_sortingColumn As Integer

    ''' <summary>
    ''' Initialize a new instance of BySortingColumn.
    ''' </summary>
    ''' <param name="sortingColumn">Column number to sort by.</param>
    Public Sub New(sortingColumn As Integer)
        m_sortingColumn = sortingColumn
    End Sub

    ''' <summary>
    ''' Compare two values.
    ''' </summary>
    ''' <param name="x">Value1 to compare.</param>
    ''' <param name="y">Value2 to compare.</param>
    ''' <returns>X less than Y: Less than 0. X=Y: 0. X greater than Y: Greater than 0.</returns>
    Public Function Compare(x As OutputRowEntryDictionary, y As OutputRowEntryDictionary) As Integer Implements IComparer(Of OutputRowEntryDictionary).Compare
        Dim xVal = CDbl(x(m_sortingColumn))
        Dim YVal = CDbl(y(m_sortingColumn))

        Return xVal.CompareTo(YVal)
    End Function
End Class

''' <summary>
''' Utilities that are used by various function throughout the program.
''' </summary>
Public Module SharedUtils
    ''' <summary>
    ''' Read 2 columns of data into a Key: Value structure. Optional list of parameters to verify were successfully read in.
    ''' </summary>
    ''' <param name="start">Excel range to begin reading data pairs at.</param>
    ''' <param name="dict">Dictionary to store data pairs in.</param>
    ''' <param name="ExpectedParameters">List of parameters that must be in the dictionary or an error will be thrown.</param>
    Public Sub ReadPairRange(start As Excel.Range, dict As Dictionary(Of String, String), ParamArray ExpectedParameters() As String)
        While Not String.IsNullOrEmpty(start.Value)
            dict(start.Value) = start.Offset(0, 1).Text
            start = start.Offset(1, 0)
        End While

        For Each p In ExpectedParameters
            If Not dict.ContainsKey(p) Then
                Throw New Exception("Unable to locate pointer." & vbCrLf & vbCrLf &
                                    "Missing: """ & p & """ from " & start.Parent.Name)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Convert an output row dictionary into a sparse string array where the 1-based output row column indices are converted to a 0-based string array index.
    ''' </summary>
    ''' <param name="outputRow">Row data to convert into a string array.</param>
    ''' <returns>Sparsely populated string array.</returns>
    Public Function OutputRowEntryDictionaryToArray(outputRow As OutputRowEntryDictionary) As String()
        ' Create string array (0 based) from max column index (1 based)
        Dim arrayUBound = outputRow.Max(Function(kv) kv.Key) - 1
        Dim s(arrayUBound) As String

        ' Store column values in string array with 1 base to 0 base index conversion
        For Each c In outputRow
            s(c.Key - 1) = c.Value
        Next

        Return s
    End Function
End Module

''' <summary>
''' Contains extension methods
''' </summary>
Public Module ExtensionMethods
    ''' <summary>
    ''' Parse a string containing formatted column / data pairs into the output row dictionary format.
    ''' </summary>
    ''' <param name="columnDataPairString">String to parse. Ex: [2, {NAME}];[3, {ADDRESS}];[5, {ALIAS}]</param>
    ''' <param name="columnDataDict">Output dictionary to store parsed data in.</param>
    <System.Runtime.CompilerServices.Extension>
    Public Sub ParseColumnDataPairs(columnDataPairString As String, columnDataDict As OutputRowEntryDictionary)
        If columnDataPairString.Length = 0 Then Return
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
    ''' <param name="columns">Column data pair dictionary to update</param>
    ''' <param name="replacements">Dictionary of keywords (like {NAME}) and their replacement</param>
    <System.Runtime.CompilerServices.Extension>
    Public Sub ReplaceTagKeywords(columns As OutputRowEntryDictionary, replacements As Dictionary(Of String, String))
        Dim keys = New List(Of Integer)(columns.Keys.ToList)
        For Each columnKey In keys
            For Each rep In replacements
                columns(columnKey) = columns(columnKey).Replace(rep.Key, rep.Value)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Search a string for a character and return the Nth character.
    ''' </summary>
    ''' <param name="s">String to search.</param>
    ''' <param name="t">Character to search for.</param>
    ''' <param name="n">Instance number of the character.</param>
    ''' <returns>Returns the Nth index as an Integer.</returns>
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