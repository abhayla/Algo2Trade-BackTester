Imports NLog
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Namespace DAL
    Public Class CSVHelper
        Implements IDisposable
#Region "Logging and Status Progress"
        Public Shared logger As Logger = LogManager.GetCurrentClassLogger
#End Region

#Region "Events"
        Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        Public Event Heartbeat(ByVal msg As String)
        'The below functions are needed to allow the derived classes to raise the above two events
        Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
            RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
        End Sub
        Protected Overridable Sub OnHeartbeat(ByVal msg As String)
            RaiseEvent Heartbeat(msg)
        End Sub
#End Region

#Region "Private Attributes"
        Private _CSVFilePath As String
        Private _separator As String
        Protected _canceller As CancellationTokenSource
#End Region

#Region "Constructors"
        Public Sub New(ByVal CSVFilePath As String, ByVal separator As String, ByVal canceller As CancellationTokenSource)
            _CSVFilePath = CSVFilePath
            _separator = separator
            _canceller = canceller
        End Sub
#End Region

#Region "Public Attributes"
#End Region

#Region "Private Methods"
#End Region

#Region "Public Methods"
        Public Function GetDataTableFromCSV(ByVal headerLineNumber As Integer) As DataTable
            Dim ret As New DataTable
            Dim headerScanned As Boolean = False
            If IO.File.Exists(_CSVFilePath) Then
                Dim totalLinesCtr As Long = File.ReadLines(_CSVFilePath).Count()
                _canceller.Token.ThrowIfCancellationRequested()
                Using sr As New IO.StreamReader(_CSVFilePath)
                    Dim runningLineNumber As Integer = 0

                    While Not sr.EndOfStream
                        runningLineNumber += 1
                        _canceller.Token.ThrowIfCancellationRequested()
                        OnHeartbeat(String.Format("Reading CSV file...({0}/{1})",
                                                  runningLineNumber,
                                                  totalLinesCtr))
                        Dim data() As String = sr.ReadLine.Split(_separator)
                        If Not headerScanned And ((headerLineNumber = 0 And runningLineNumber = 1) Or (headerLineNumber > 0 And runningLineNumber = headerLineNumber)) Then
                            headerScanned = True
                            Dim colCtr As Integer = 0
                            For Each col In data
                                _canceller.Token.ThrowIfCancellationRequested()
                                colCtr += 1
                                ret.Columns.Add(New DataColumn(String.Format("{0}_{1}", col, colCtr), GetType(String)))
                            Next
                        End If
                        If headerScanned Then
                            ret.Rows.Add(data.ToArray)
                        End If
                    End While
                End Using
            Else
                Throw New ApplicationException(String.Format("CSV file was not found, fileName:{0}", _CSVFilePath))
            End If
            Return ret
        End Function

        Public Sub GetCSVFromDataTable(ByVal dt As DataTable)
            Dim sb As New StringBuilder
            For i As Integer = 0 To dt.Columns.Count - 1
                sb.Append(dt.Columns(i))
                If i < (dt.Columns.Count - 1) Then
                    sb.Append(_separator)
                End If
            Next
            sb.AppendLine()
            For Each dr As DataRow In dt.Rows
                For i As Integer = 0 To dt.Columns.Count - 1
                    sb.Append(dr(i).ToString())
                    If i < (dt.Columns.Count - 1) Then
                        sb.Append(_separator)
                    End If
                Next
                sb.AppendLine()
            Next
            File.WriteAllText(_CSVFilePath, sb.ToString())
        End Sub

        Public Sub GetCSVFromDataGrid(ByVal dg As DataGridView)
            Try
                Dim csvFileWriter As StreamWriter = New StreamWriter(_CSVFilePath, False)
                Dim columnHeaderText As String = ""
                Dim countColumn As Integer = dg.ColumnCount - 1
                If countColumn >= 0 Then
                    columnHeaderText = dg.Columns(0).HeaderText
                End If
                For i As Integer = 1 To countColumn
                    columnHeaderText = columnHeaderText + _separator + dg.Columns(i).HeaderText
                Next
                csvFileWriter.WriteLine(columnHeaderText)
                For Each dataRowObject As DataGridViewRow In dg.Rows
                    If Not dataRowObject.IsNewRow Then
                        Dim dataFromGrid As String = ""
                        dataFromGrid = dataRowObject.Cells(0).Value.ToString()
                        For i As Integer = 1 To countColumn
                            dataFromGrid = dataFromGrid + _separator + dataRowObject.Cells(i).Value.ToString()
                        Next
                        csvFileWriter.WriteLine(dataFromGrid)
                    End If
                Next
                csvFileWriter.Flush()
                csvFileWriter.Close()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
        End Sub
#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    _CSVFilePath = Nothing
                    _separator = Nothing
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace
