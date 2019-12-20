Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports OutputRowEntryDictionary = System.Collections.Generic.Dictionary(Of Integer, String)
Imports PointList = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of Integer, String))

''' <summary>
''' Builds the RTAC tag processor map
''' </summary>
Public Class RtacTagProcessorWorksheet
    ''' <summary>
    ''' Output list of tag processor entries
    ''' </summary>
    Private _Map As New List(Of TagProcessorMapEntry)

    ''' <summary>
    ''' Add new entry to the tag processor map.
    ''' </summary>
    ''' <param name="ScadaTag">SCADA tag name</param>
    ''' <param name="ScadaTagDataType">SCADA tag types</param>
    ''' <param name="Ied">Device tag name</param>
    ''' <param name="IedDataType">Device tag type</param>
    ''' <param name="Direction">Direction data flows. Determines if SCADA or a device is the destination.</param>
    Public Sub AddEntry(ScadaTag As String, ScadaTagDataType As String,
                        Ied As String, IedDataType As String,
                        Direction As String)
        If Direction = RtacTemplate.Keywords.DATA_IN Then
            ' Flow: IED -> SCADA
            Me._Map.Add(New TagProcessorMapEntry(ScadaTag, ScadaTagDataType,
                                                Ied, IedDataType,
                                                Direction)
                                            )
        ElseIf Direction = RtacTemplate.Keywords.DATA_OUT Then
            ' Flow: SCADA -> IED
            Me._Map.Add(
                New TagProcessorMapEntry(
                    ScadaTag, ScadaTagDataType,
                    Ied, IedDataType,
                    Direction
                )
            )
        Else
            Throw New ArgumentException("Invalid Direction")
        End If
    End Sub

    ''' <summary>
    ''' Write the tag processor map out to CSV
    ''' </summary>
    ''' <param name="path">Source filename to append output suffix on.</param>
    Public Sub WriteCsv(path As String)
        Dim csvPath = IO.Path.GetDirectoryName(path) & IO.Path.DirectorySeparatorChar & IO.Path.GetFileNameWithoutExtension(path) & "_TagProcessor.csv"
        Using csvStreamWriter = New IO.StreamWriter(csvPath, False)
            Dim csvWriter As New CsvHelper.CsvWriter(csvStreamWriter)

            ' Write header
            csvWriter.WriteField("Destination Tag Name")
            csvWriter.WriteField("DT Data Type")
            csvWriter.WriteField("Source Expression")
            csvWriter.WriteField("SE Data Type")
            csvWriter.WriteField("Time Source")
            csvWriter.WriteField("Quality Source")
            csvWriter.NextRecord()

            ' Write IED map
            For Each Tag In Me._Map
                csvWriter.WriteField(Tag.DestinationTagName)
                csvWriter.WriteField(Tag.DestinationTagDataType)
                csvWriter.WriteField(Tag.SourceExpression)
                csvWriter.WriteField(Tag.SourceExpressionDataType)
                csvWriter.NextRecord()
            Next
        End Using
    End Sub

    ''' <summary>
    ''' Stores data for each tag processor map entry
    ''' </summary>
    Private Class TagProcessorMapEntry
        Private _DestinationTagName As String
        Public ReadOnly Property DestinationTagName As String
            Get
                Return Me._DestinationTagName
            End Get
        End Property

        Private _DestinationTagDataType As String
        Public ReadOnly Property DestinationTagDataType As String
            Get
                Return Me._DestinationTagDataType
            End Get
        End Property

        Private _SourceExpression As String
        Public ReadOnly Property SourceExpression As String
            Get
                Return Me._SourceExpression
            End Get
        End Property

        Private _SourceExpressionDataType As String
        Public ReadOnly Property SourceExpressionDataType As String
            Get
                Return Me._SourceExpressionDataType
            End Get
        End Property

        Private _ParsedDeviceName As String
        Private _ParsedTagName As String
        Private Const RegexMatch As String = ".*?(\p{L}\w*\.\p{L}\w*).*"
        Private Const RegexTagName As String = "$1.$2"
        Private Const RegexDeviceName As String = "$1"
        ''' <summary>
        ''' Locates the device name and point name from a source expression.
        ''' </summary>
        Private Sub ParseSourceExpression()
            Me._ParsedDeviceName = Regex.Replace(SourceExpression, RegexMatch, RegexDeviceName)
            Me._ParsedTagName = Regex.Replace(SourceExpression, RegexMatch, RegexTagName)
        End Sub

        Private _Direction As String
        Public ReadOnly Property Direction As String
            Get
                Return Me._Direction
            End Get
        End Property

        Public Sub New(DestinationTag As String, DestinationTagDataType As String,
                       SourceExpression As String, SourceExpressionDataType As String,
                       Direction As String)
            Me._DestinationTagName = DestinationTag
            Me._DestinationTagDataType = DestinationTagDataType
            Me._SourceExpression = SourceExpression
            Me._SourceExpressionDataType = SourceExpressionDataType
            Me._Direction = Direction

            Me.ParseSourceExpression()
        End Sub
    End Class
End Class