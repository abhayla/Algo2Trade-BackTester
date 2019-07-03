Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading
Imports MySql.Data.MySqlClient
Public Class VWALLRule

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
        outputdt.Columns.Add("Signal Direction")

        Dim chkDate As Date = startDate
        While chkDate <= endDate
            Dim stockList As List(Of String) = Nothing
            stockList = GetStockList(chkDate)

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                For Each stock In stockList
                    Dim firstCandle As Boolean = True
                    Dim currentDayPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim indicatorPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
                    Dim XMinuteHAPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim oneMinutePayload As Dictionary(Of Date, Payload) = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, stock, chkDate.AddDays(-7), chkDate)
                    If oneMinutePayload IsNot Nothing AndAlso oneMinutePayload.Count > 0 Then
                        If signalTimeFrame > 1 Then
                            indicatorPayload = cmn.ConvertPayloadsToXMinutes(oneMinutePayload, signalTimeFrame)
                        Else
                            indicatorPayload = oneMinutePayload
                        End If
                        If indicatorPayload IsNot Nothing AndAlso indicatorPayload.Count > 0 Then
                            Indicator.HeikenAshi.ConvertToHeikenAshi(indicatorPayload, XMinuteHAPayload)
                            Indicator.VWAP.CalculateVWAP(XMinuteHAPayload, VWAPPayload)
                            For Each runningPayload In XMinuteHAPayload.Keys
                                If chkDate.Date = runningPayload.Date Then
                                    If currentDayPayload Is Nothing Then currentDayPayload = New Dictionary(Of Date, Payload)
                                    currentDayPayload.Add(runningPayload, XMinuteHAPayload(runningPayload))
                                End If
                            Next
                        End If
                    End If
                    If currentDayPayload IsNot Nothing AndAlso currentDayPayload.Count > 0 Then
                        Dim noSignal As Boolean = True
                        For Each runningPayload In currentDayPayload.Keys
                            Dim potentialEntryPrice As Double = Nothing
                            If Not firstCandle AndAlso runningPayload.Hour < 15 Then
                                If currentDayPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bullish AndAlso
                                    currentDayPayload(runningPayload).PreviousCandlePayload.High > VWAPPayload(currentDayPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                    currentDayPayload(runningPayload).PreviousCandlePayload.Low < VWAPPayload(currentDayPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                    potentialEntryPrice = currentDayPayload(runningPayload).PreviousCandlePayload.Low
                                    potentialEntryPrice -= CalculateBuffer(potentialEntryPrice, RoundOfType.Celing)
                                    If currentDayPayload(runningPayload).Low < potentialEntryPrice Then
                                        noSignal = False
                                        Dim row As DataRow = outputdt.NewRow
                                        row("Date") = runningPayload
                                        row("Instrument") = currentDayPayload(runningPayload).TradingSymbol
                                        row("Signal Direction") = "Sell"
                                        outputdt.Rows.Add(row)
                                    End If
                                ElseIf currentDayPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bearish AndAlso
                                        currentDayPayload(runningPayload).PreviousCandlePayload.High > VWAPPayload(currentDayPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                        currentDayPayload(runningPayload).PreviousCandlePayload.Low < VWAPPayload(currentDayPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                    potentialEntryPrice = currentDayPayload(runningPayload).PreviousCandlePayload.High
                                    potentialEntryPrice += CalculateBuffer(potentialEntryPrice, RoundOfType.Celing)
                                    If currentDayPayload(runningPayload).High > potentialEntryPrice Then
                                        noSignal = False
                                        Dim row As DataRow = outputdt.NewRow
                                        row("Date") = runningPayload
                                        row("Instrument") = currentDayPayload(runningPayload).TradingSymbol
                                        row("Signal Direction") = "Buy"
                                        outputdt.Rows.Add(row)
                                    End If
                                End If
                            End If
                            firstCandle = False
                        Next
                        If noSignal Then
                            Dim row As DataRow = outputdt.NewRow
                            row("Date") = chkDate
                            row("Instrument") = stock
                            row("Signal Direction") = "No Signal"
                            outputdt.Rows.Add(row)
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
            OnHeartbeat("Fetching Pre Market Volume Spike Data")
            Dim cmd As New MySqlCommand("GET_PRE_MARKET_VOLUME_SPIKE_PRE_MARKET_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", tradeDate)
            cmd.Parameters.AddWithValue("@endDate", tradeDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@spikeChangePercentage", 100)
            cmd.Parameters.AddWithValue("@minClose", 100)
            cmd.Parameters.AddWithValue("@maxClose", 1500)
            cmd.Parameters.AddWithValue("@atrPercentage", 4)
            cmd.Parameters.AddWithValue("@potentialAmount", 100000)
            cmd.Parameters.AddWithValue("@sortColumn", "QuantityXPrice")

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
    Public Function CalculateBuffer(ByVal price As Double, ByVal floorOrCeiling As RoundOfType) As Double
        Dim bufferPrice As Double = Nothing
        'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
        bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, 0.05, floorOrCeiling)
        Return bufferPrice
    End Function
End Class
