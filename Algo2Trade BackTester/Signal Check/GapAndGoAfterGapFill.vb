Imports Algo2TradeBLL
Imports System.Threading
Imports MySql.Data.MySqlClient
Public Class GapAndGoAfterGapFill
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
        outputdt.Columns.Add("Change %")
        outputdt.Columns.Add("Lot Size")

        Dim chkDate As Date = startDate
        While chkDate <= endDate
            Dim stockList As Dictionary(Of String, Double()) = Nothing
            stockList = GetModifiedStockData(Common.DataBaseTable.Intraday_Cash, chkDate)

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                For Each stock In stockList.Keys
                    Dim row As DataRow = outputdt.NewRow
                    row("Date") = chkDate.Date
                    row("Instrument") = stock
                    row("Change %") = stockList(stock)(2)
                    row("Lot Size") = stockList(stock)(1)
                    outputdt.Rows.Add(row)
                Next
            End If
            chkDate = chkDate.AddDays(1)
        End While
        Return outputdt
    End Function
    Private Function GetStockData(databaseTableType As Common.DataBaseTable, tradingDate As Date) As Dictionary(Of String, Double())
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        Dim dts As DataSet = Nothing
        Dim conn As MySqlConnection = cmn.OpenDBConnection
        Dim outputPayload As Dictionary(Of String, Double()) = Nothing

        If conn.State = ConnectionState.Open Then
            OnHeartbeat(String.Format("Fetching Pre Market Data for {0}", tradingDate.ToShortDateString))
            Dim cmd As New MySqlCommand("GET_PRE_MARKET_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", tradingDate)
            cmd.Parameters.AddWithValue("@endDate", tradingDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@minClose", 80)
            cmd.Parameters.AddWithValue("@maxClose", Double.MaxValue)
            cmd.Parameters.AddWithValue("@atrPercentage", 0)
            cmd.Parameters.AddWithValue("@potentialAmount", 0)
            cmd.Parameters.AddWithValue("@sortColumn", "ChangePer")

            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 3000
            dts = New DataSet
            adapter.Fill(dts)
        End If
        If dts IsNot Nothing AndAlso dts.Tables.Count > 0 Then
            Dim totalTables As Integer = dts.Tables.Count
            Dim dt As DataTable = Nothing
            Dim count As Integer = 0
            While Not count > totalTables - 1
                Dim temp_dt As New DataTable
                temp_dt = dts.Tables(count)
                If temp_dt.Rows.Count > 0 Then
                    Dim tempCol As DataColumn = temp_dt.Columns.Add("Abs", GetType(Double), "Iif(ChangePer <0 ,ChangePer*(-1),ChangePer)")
                    temp_dt = temp_dt.Select("Abs>=0.8", "Abs DESC").Cast(Of DataRow).Take(100).CopyToDataTable
                End If
                If dt Is Nothing Then dt = New DataTable
                dt.Merge(temp_dt)
                count += 1
            End While
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i).Item(7) > 1000 Then
                        Dim signalDirection As Integer = If(dt.Rows(i).Item(5) >= 0, 1, -1)
                        If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Double())
                        'TODO:
                        'Change lot size
                        outputPayload.Add(dt.Rows(i).Item(1), {signalDirection, dt.Rows(i).Item(11), dt.Rows(i).Item(5), dt.Rows(i).Item(3)})
                    End If
                Next
            End If
        End If
        Return outputPayload
    End Function
    Private Function GetModifiedStockData(databaseTableType As Common.DataBaseTable, tradingDate As Date) As Dictionary(Of String, Double())
        Dim ret As Dictionary(Of String, Double()) = Nothing
        Dim previousTradingDate As Date = cmn.GetPreviousTradingDay(databaseTableType, tradingDate)
        Dim tempStocklist As Dictionary(Of String, Double()) = GetStockData(databaseTableType, tradingDate)
        Dim signalTime As TimeSpan = TimeSpan.Parse("09:20:00")
        Dim secondSignalTime As TimeSpan = TimeSpan.Parse("09:30:00")
        Dim potentialCandleSignalTime As Date = New Date(tradingDate.Year, tradingDate.Month, tradingDate.Day, signalTime.Hours, signalTime.Minutes, signalTime.Seconds)
        Dim potentialSecondSignalTime As Date = New Date(tradingDate.Year, tradingDate.Month, tradingDate.Day, secondSignalTime.Hours, secondSignalTime.Minutes, secondSignalTime.Seconds)
        If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
            For Each stock In tempStocklist.Keys
                Dim eodPayload As Dictionary(Of Date, Payload) = Nothing
                Select Case databaseTableType
                    Case Common.DataBaseTable.Intraday_Cash
                        eodPayload = cmn.GetRawPayload(Common.DataBaseTable.EOD_Cash, stock, previousTradingDate, previousTradingDate)
                    Case Common.DataBaseTable.EOD_Commodity
                        eodPayload = cmn.GetRawPayload(Common.DataBaseTable.EOD_Commodity, stock, previousTradingDate, previousTradingDate)
                    Case Common.DataBaseTable.Intraday_Currency
                        eodPayload = cmn.GetRawPayload(Common.DataBaseTable.EOD_Currency, stock, previousTradingDate, previousTradingDate)
                    Case Common.DataBaseTable.Intraday_Futures
                        eodPayload = cmn.GetRawPayload(Common.DataBaseTable.EOD_Futures, stock, previousTradingDate, previousTradingDate)
                    Case Else
                        Throw New ApplicationException("Wrong Database Table Type Entry")
                End Select
                Dim intradayPayload As Dictionary(Of Date, Payload) = cmn.GetRawPayload(databaseTableType, stock, tradingDate, tradingDate)
                If intradayPayload IsNot Nothing AndAlso intradayPayload.Count > 0 Then

                    Dim firstCandle As Payload = intradayPayload.Values.FirstOrDefault

                    Dim previousHigh As Decimal = intradayPayload.Max(Function(x)
                                                                          If x.Key < potentialCandleSignalTime Then
                                                                              Return x.Value.High
                                                                          Else
                                                                              Return Decimal.MinValue
                                                                          End If
                                                                      End Function)

                    Dim previousLow As Decimal = intradayPayload.Min(Function(x)
                                                                         If x.Key < potentialCandleSignalTime Then
                                                                             Return x.Value.Low
                                                                         Else
                                                                             Return Decimal.MaxValue
                                                                         End If
                                                                     End Function)

                    Dim secondHigh As Decimal = intradayPayload.Max(Function(x)
                                                                        If x.Key < potentialSecondSignalTime Then
                                                                            Return x.Value.High
                                                                        Else
                                                                            Return Decimal.MinValue
                                                                        End If
                                                                    End Function)

                    Dim secondLow As Decimal = intradayPayload.Min(Function(x)
                                                                       If x.Key < potentialSecondSignalTime Then
                                                                           Return x.Value.Low
                                                                       Else
                                                                           Return Decimal.MaxValue
                                                                       End If
                                                                   End Function)

                    Dim chkValue As Decimal = Nothing
                    If firstCandle.Open > eodPayload.LastOrDefault.Value.High Then
                        chkValue = eodPayload.LastOrDefault.Value.High
                    ElseIf firstCandle.Open < eodPayload.LastOrDefault.Value.Low Then
                        chkValue = eodPayload.LastOrDefault.Value.Low
                    Else
                        chkValue = eodPayload.LastOrDefault.Value.Close
                    End If

                    If tempStocklist(stock)(0) = 1 Then
                        If previousLow <= chkValue AndAlso previousHigh < secondHigh Then
                            If ret Is Nothing Then ret = New Dictionary(Of String, Double())
                            ret.Add(stock, tempStocklist(stock))
                        End If
                    ElseIf tempStocklist(stock)(0) = -1 Then
                        If previousHigh >= chkValue AndAlso previousLow > secondLow Then
                            If ret Is Nothing Then ret = New Dictionary(Of String, Double())
                            ret.Add(stock, tempStocklist(stock))
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class
