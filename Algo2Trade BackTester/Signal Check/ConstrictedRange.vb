Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading
Imports MySql.Data.MySqlClient
Public Class ConstrictedRange
#Region "Events/Event handlers"
    Public Event DocumentDownloadComplete()
    Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
    Public Event Heartbeat(ByVal msg As String)
    Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
    'The below functions are needed to allow the derived classes to raise the above two events
    Protected Overridable Sub OnDocumentDownloadComplete()
        RaiseEvent DocumentDownloadComplete()
    End Sub
    Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        RaiseEvent DocumentRetryStatus(currentTry, totalTries)
    End Sub
    Protected Overridable Sub OnHeartbeat(ByVal msg As String)
        RaiseEvent Heartbeat(msg)
    End Sub
    Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
    End Sub
#End Region

    Private _canceller As CancellationTokenSource
    Dim cmn As Common = New Common(_canceller)
    Public Sub New(ByVal cts As CancellationTokenSource)
        _canceller = cts
    End Sub
    Public Async Function TestDataAsync(ByVal startDate As Date, ByVal endDate As Date, ByVal signalTimeFrame As Integer) As Task(Of DataTable)
        Await Task.Delay(1).ConfigureAwait(False)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim outputdt As DataTable = Nothing
        outputdt = New DataTable
        outputdt.Columns.Add("Date")
        outputdt.Columns.Add("Instrument")

        Dim chkDate As Date = startDate
        While chkDate <= endDate
            Dim stockList As List(Of String) = Nothing
            stockList = GetStockList(chkDate)
            'stockList = New List(Of String)
            'stockList.Add("RELINFRA")

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                For Each stock In stockList
                    Dim firstCandle As Boolean = True
                    Dim currentDayPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim indicatorPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim indicatorHAPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim oneMinutePayload As Dictionary(Of Date, Payload) = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, stock, chkDate.AddDays(-5), chkDate)
                    If oneMinutePayload IsNot Nothing AndAlso oneMinutePayload.Count > 0 Then
                        If signalTimeFrame > 1 Then
                            indicatorPayload = cmn.ConvertPayloadsToXMinutes(oneMinutePayload, signalTimeFrame)
                        Else
                            indicatorPayload = oneMinutePayload
                        End If
                        If indicatorPayload IsNot Nothing AndAlso indicatorPayload.Count > 0 Then
                            Indicator.HeikenAshi.ConvertToHeikenAshi(indicatorPayload, indicatorHAPayload)
                            For Each runningPayload In indicatorPayload.Keys
                                If chkDate.Date = runningPayload.Date Then
                                    If currentDayPayload Is Nothing Then currentDayPayload = New Dictionary(Of Date, Payload)
                                    currentDayPayload.Add(runningPayload, indicatorPayload(runningPayload))
                                End If
                            Next
                        End If
                    End If
                    If currentDayPayload IsNot Nothing AndAlso currentDayPayload.Count > 0 Then
                        Dim outputPayload As Dictionary(Of Date, String) = CalculateConstrictedRange(chkDate, indicatorHAPayload)
                        If outputPayload IsNot Nothing AndAlso outputPayload.Count > 0 Then
                            For Each runningPayload In outputPayload.Keys
                                Dim row As DataRow = outputdt.NewRow
                                row("Date") = runningPayload
                                row("Instrument") = outputPayload(runningPayload)
                                outputdt.Rows.Add(row)
                            Next
                        End If
                    End If
                Next
            End If
            chkDate = chkDate.AddDays(1)
        End While
        Return outputdt
    End Function
    Private Function GetStockList(ByVal tradeDate As Date) As List(Of String)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim stockList As List(Of String) = Nothing
        Dim dt As DataTable = Nothing
        Dim conn As MySqlConnection = cmn.OpenDBConnection

        If conn.State = ConnectionState.Open Then
            OnHeartbeat("Fetching All Stock Data")
            Dim cmd As New MySqlCommand("GET_STOCK_CASH_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", tradeDate)
            cmd.Parameters.AddWithValue("@endDate", tradeDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@minClose", 100)
            cmd.Parameters.AddWithValue("@maxClose", 1500)
            cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
            cmd.Parameters.AddWithValue("@potentialAmount", 1000000)

            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 3000
            dt = New DataTable
            adapter.Fill(dt)

        End If

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            stockList = New List(Of String)
            For i As Integer = 0 To dt.Rows.Count - 1
                stockList.Add(dt.Rows(i).Item(1))
            Next
        End If

        Return stockList
    End Function
    Private Function CalculateBuffer(ByVal price As Double, ByVal floorOrCeiling As RoundOfType) As Double
        Dim bufferPrice As Double = Nothing
        'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
        bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, 0.05, floorOrCeiling)
        Return bufferPrice
    End Function
    Private Function CalculateConstrictedRange(ByVal tradeDate As Date, ByVal inputPayload As Dictionary(Of Date, Payload)) As Dictionary(Of Date, String)
        Dim ret As Dictionary(Of Date, String) = Nothing
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
            OnHeartbeat(String.Format("Calculating Fractals for {0} On {1}", inputPayload.Values.LastOrDefault.TradingSymbol, tradeDate.ToShortDateString))
            Indicator.Fractals.CalculateFractal(5, inputPayload, fractalHighPayload, fractalLowPayload)
            OnHeartbeat(String.Format("Checking Signal for {0} On {1}", inputPayload.Values.LastOrDefault.TradingSymbol, tradeDate.ToShortDateString))
            Dim lastFractalHigh As Decimal = Decimal.MinValue
            Dim previousLastFractalHigh As Decimal = Decimal.MinValue
            Dim currentFractalHigh As Decimal = Decimal.MinValue
            Dim lastFractalLow As Decimal = Decimal.MaxValue
            Dim previousLastFractalLow As Decimal = Decimal.MaxValue
            Dim currentFractalLow As Decimal = Decimal.MaxValue
            Dim lastSignalTime As Date = "14:30:00"
            lastSignalTime = New Date(tradeDate.Year, tradeDate.Month, tradeDate.Day, lastSignalTime.Hour, lastSignalTime.Minute, lastSignalTime.Second)
            For Each runningPayload In inputPayload.Keys
                If runningPayload.Date = tradeDate.Date AndAlso runningPayload < lastSignalTime Then
                    If fractalHighPayload(runningPayload) <> currentFractalHigh Then
                        previousLastFractalHigh = lastFractalHigh
                        lastFractalHigh = currentFractalHigh
                        currentFractalHigh = fractalHighPayload(runningPayload)
                    End If
                    If fractalLowPayload(runningPayload) <> currentFractalLow Then
                        previousLastFractalLow = lastFractalLow
                        lastFractalLow = currentFractalLow
                        currentFractalLow = fractalLowPayload(runningPayload)
                    End If
                    If fractalHighPayload(runningPayload) < lastFractalHigh AndAlso lastFractalHigh < previousLastFractalHigh AndAlso
                        fractalLowPayload(runningPayload) > lastFractalLow AndAlso lastFractalLow > previousLastFractalLow Then
                        If fractalHighPayload(runningPayload) <> fractalHighPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) OrElse
                            fractalLowPayload(runningPayload) <> fractalLowPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                            If ret Is Nothing Then ret = New Dictionary(Of Date, String)
                            ret.Add(runningPayload, inputPayload(runningPayload).TradingSymbol)
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class
