Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports MySql.Data.MySqlClient
Public Class InTheTrend
    Inherits Strategy
    Implements IDisposable

    Private Property _SignalTimeFrame As Integer
    Private Property _ATRShift As Integer
    Private Property _ATRPeriod As Integer
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan,
                   ByVal signalTimeFrame As Integer,
                   ByVal atrShift As Integer,
                   ByVal atrPeriod As Integer)
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
        Me._SignalTimeFrame = signalTimeFrame
        Me._ATRShift = atrShift
        Me._ATRPeriod = atrPeriod
    End Sub
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(1).ConfigureAwait(False)
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        'TODO:
        'ChangeType
        Dim tradeStockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
        Dim databaseTable As Common.DataBaseTable = Common.DataBaseTable.Intraday_Cash

        Dim tradeCheckingDate As Date = startDate
        While tradeCheckingDate <= endDate
            Dim stockList As Dictionary(Of String, Double()) = Nothing
            Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
            'TODO:
            'Change StockList
            'tempStocklist = New Dictionary(Of String, Double()) From {{"INDIGO", {600}}}
            'tempStocklist = New Dictionary(Of String, Double()) From {{"HEXAWARE", {1500}}, {"BHARATFIN", {500}}, {"KOTAKBANK", {800}}, {"JETAIRWAYS", {2200}}, {"BHARTIARTL", {1700}}, {"KAJARIACER", {1300}}}
            'tempStocklist = New Dictionary(Of String, Double()) From {{"DHFL", {1500}}, {"JUSTDIAL", {1400}}, {"JETAIRWAYS", {2200}}, {"CANBK", {2000}}, {"EQUITAS", {4000}}, {"RAMCOCEM", {800}}}
            tempStocklist = GetStockData(databaseTable, tradeCheckingDate)

            If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
                For Each tradingSymbol In tempStocklist.Keys
                    If stockList Is Nothing Then stockList = New Dictionary(Of String, Double())
                    stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0)})
                Next
            End If

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim VWAPStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim HighEntryStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim LowEntryStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SignalStocksPayload As Dictionary(Of String, Dictionary(Of Date, Integer)) = Nothing

                'First lets build the payload for all the stocks
                For Each stock In stockList.Keys
                    Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim XMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    'Get the last 5 days payload
                    Dim oneMinutePayload As Dictionary(Of Date, Payload) = Cmn.GetRawPayload(databaseTable, stock, tradeCheckingDate.AddDays(-5), tradeCheckingDate)
                    'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date isa valid date)
                    If oneMinutePayload IsNot Nothing AndAlso oneMinutePayload.Count > 0 Then
                        For Each tradingDate In oneMinutePayload.Keys
                            If tradingDate.Date = tradeCheckingDate.Date Then
                                If currentDayOneMinutePayload Is Nothing Then currentDayOneMinutePayload = New Dictionary(Of Date, Payload)
                                currentDayOneMinutePayload.Add(tradingDate, oneMinutePayload(tradingDate))
                            End If
                        Next
                    End If

                    If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                        OnHeartbeat(String.Format("Processing for {0}", tradeCheckingDate.ToShortDateString))
                        'Convert the one-minute payload of the last 5 days to the corresponding timeframe payload that this strategy needs for signalling
                        If _SignalTimeFrame > 1 Then
                            XMinutePayload = Cmn.ConvertPayloadsToXMinutes(oneMinutePayload, _SignalTimeFrame)
                        Else
                            XMinutePayload = oneMinutePayload
                        End If

                        'Now get the rule specific to this strategy by passing the XMinute payload
                        Dim ruleXMinuteHighEntryPayload As Dictionary(Of Date, Decimal) = Nothing
                        Dim ruleXMinuteLowEntryPayload As Dictionary(Of Date, Decimal) = Nothing
                        Dim ruleXMinuteVWAPPayload As Dictionary(Of Date, Decimal) = Nothing
                        Dim ruleXMinuteSignalPayload As Dictionary(Of Date, Integer) = Nothing
                        Dim ruleXMinutePearsingCandleSignalPayload As Dictionary(Of Date, Integer) = Nothing

                        StrategyRules.InTheTrendRule.CalculateInTheTrendRule(_ATRShift, _ATRPeriod, XMinutePayload, ruleXMinuteVWAPPayload, ruleXMinuteHighEntryPayload, ruleXMinuteLowEntryPayload, ruleXMinuteSignalPayload, ruleXMinutePearsingCandleSignalPayload)


                        'Add all these payloads into the stock collections
                        If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                        If XMinuteStocksPayload Is Nothing Then XMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        XMinuteStocksPayload.Add(stock, XMinutePayload)
                        If HighEntryStocksPayload Is Nothing Then HighEntryStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                        HighEntryStocksPayload.Add(stock, ruleXMinuteHighEntryPayload)
                        If LowEntryStocksPayload Is Nothing Then LowEntryStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                        LowEntryStocksPayload.Add(stock, ruleXMinuteLowEntryPayload)
                        If VWAPStocksPayload Is Nothing Then VWAPStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                        VWAPStocksPayload.Add(stock, ruleXMinuteVWAPPayload)
                        If SignalStocksPayload Is Nothing Then SignalStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Integer))
                        SignalStocksPayload.Add(stock, ruleXMinuteSignalPayload)
                    End If
                Next

                '------------------------------------------------------------------------------------------------------------------------------------------------

                If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                    OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                    Dim startMinute As TimeSpan = ExchangeStartTime
                    Dim endMinute As TimeSpan = ExchangeEndTime
                    While startMinute < endMinute
                        If startMinute = ExchangeStartTime Then
                            Dim oneMinute As TimeSpan = TimeSpan.Parse("00:01:00")
                            startMinute = startMinute.Add(oneMinute)
                        End If

                        Dim startSecond As TimeSpan = startMinute
                        Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromSeconds(59))
                        Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                        Dim potentialTickSignalTime As Date = Nothing
                        Dim currentMinuteCandlePayload As Payload = Nothing

                        While startSecond < endSecond
                            potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                            For Each stockName In stockList.Keys
                                Dim runningTrade As Trade = Nothing
                                Dim signalCandleTime As Date = Nothing
                                Dim currentStockPLAfterBrokerage As Double = StockPLAfterBrokerage(tradeCheckingDate, stockName)
                                'Get the current minute candle from the stock collection for this stock for that day
                                Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                    currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialCandleSignalTime)
                                End If
                                'Now get the ticks for this minute and second

                                If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.Ticks IsNot Nothing AndAlso
                                    Not IsStockAnyPLReachedForTheDay(currentMinuteCandlePayload, Trade.TradeType.MIS) Then
                                    currentSecondTickPayload = currentMinuteCandlePayload.Ticks.FindAll(Function(x)
                                                                                                            Return x.PayloadDate = potentialTickSignalTime
                                                                                                        End Function)

                                    If _SignalTimeFrame > 1 Then
                                        signalCandleTime = GetPreviousXMinuteCandleTime(currentMinuteCandlePayload.PayloadDate, XMinuteStocksPayload(stockName), _SignalTimeFrame)
                                    Else
                                        signalCandleTime = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate
                                    End If
                                End If

                                'Main Strategy Logic
                                If currentSecondTickPayload IsNot Nothing AndAlso currentSecondTickPayload.Count > 0 Then
                                    Dim finalEntryPrice As Double = Nothing
                                    Dim finalTargetPrice As Double = Nothing
                                    Dim finalTargetRemark As String = Nothing
                                    Dim finalStoplossPrice As Double = Nothing
                                    Dim finalStoplossRemark As String = Nothing
                                    Dim potentialTargetOnMaxPLForStock As Double = Nothing
                                    Dim potentialStoploss As Double = Nothing
                                    Dim quantity As Integer = stockList(stockName)(0)
                                    Dim entryBuffer As Decimal = Nothing
                                    Dim stoplossBuffer As Decimal = Nothing
                                    Dim highEntryPrice As Decimal = Nothing
                                    Dim lowEntryPrice As Decimal = Nothing

                                    For Each tick In currentSecondTickPayload
                                        SetCurrentLTPForStock(currentMinuteCandlePayload, tick, Trade.TradeType.MIS)

                                        Dim openTradesToModifyOrCancel As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)

                                        If Not IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                                            ((HighEntryStocksPayload(stockName)(signalCandleTime) + LowEntryStocksPayload(stockName)(signalCandleTime)) / 2) < tick.Open Then

                                            finalEntryPrice = highEntryPrice
                                            entryBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Celing)
                                            finalEntryPrice += entryBuffer

                                            potentialTargetOnMaxPLForStock = finalEntryPrice + finalEntryPrice * PerTradeMaxProfitPercentage / 100
                                            finalTargetPrice = potentialTargetOnMaxPLForStock
                                            finalTargetRemark = "Max Profit% For Stock"

                                            potentialStoploss = lowEntryPrice
                                            finalStoplossPrice = potentialStoploss
                                            finalStoplossRemark = "Low ATR Band Stoploss Below VWAP"
                                            stoplossBuffer = CalculateBuffer(finalStoplossPrice, RoundOfType.Floor)
                                            finalStoplossPrice -= stoplossBuffer
                                            'Ensure that the loss is not more than per trade max loss percentage
                                            If ((LowEntryStocksPayload(stockName)(signalCandleTime) / finalEntryPrice) - 1) * 100 * -1 < -1 * PerTradeMaxLossPercentage Then
                                                runningTrade = New Trade(Me,
                                                                     currentMinuteCandlePayload.TradingSymbol,
                                                                     tradeStockType,
                                                                     currentMinuteCandlePayload.PayloadDate,
                                                                     Trade.TradeExecutionDirection.Buy,
                                                                     finalEntryPrice,
                                                                     entryBuffer,
                                                                     Trade.TradeType.MIS,
                                                                     Trade.TradeEntryCondition.Original,
                                                                     "High ATR Band Above VWAP",
                                                                     quantity,
                                                                     finalTargetPrice,
                                                                     finalTargetRemark,
                                                                     finalStoplossPrice,
                                                                     stoplossBuffer,
                                                                     finalStoplossRemark,
                                                                     currentMinuteCandlePayload)

                                                If openTradesToModifyOrCancel IsNot Nothing AndAlso openTradesToModifyOrCancel.Count > 0 Then
                                                    For Each openTrade In openTradesToModifyOrCancel
                                                        If openTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                            PlaceOrModifyOrder(openTrade, runningTrade)
                                                        ElseIf openTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                            CancelTrade(openTrade, currentMinuteCandlePayload, "Opposite Direction Signal Triggered")
                                                            PlaceOrModifyOrder(runningTrade, Nothing)
                                                        End If
                                                    Next
                                                Else
                                                    If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                                End If
                                            End If
                                        ElseIf Not IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                                                ((HighEntryStocksPayload(stockName)(signalCandleTime) + LowEntryStocksPayload(stockName)(signalCandleTime)) / 2) >= tick.Open Then

                                            finalEntryPrice = lowEntryPrice
                                            entryBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)
                                            finalEntryPrice -= entryBuffer

                                            potentialTargetOnMaxPLForStock = finalEntryPrice - finalEntryPrice * PerTradeMaxProfitPercentage / 100
                                            finalTargetPrice = potentialTargetOnMaxPLForStock
                                            finalTargetRemark = "Max Profit% For Stock"

                                            potentialStoploss = highEntryPrice
                                            finalStoplossPrice = potentialStoploss
                                            finalStoplossRemark = "High ATR Band Stoploss Above VWAP"
                                            stoplossBuffer = CalculateBuffer(finalStoplossPrice, RoundOfType.Floor)
                                            finalStoplossPrice += stoplossBuffer
                                            'Ensure that the loss is not more than per trade max loss percentage
                                            If ((HighEntryStocksPayload(stockName)(signalCandleTime) / finalEntryPrice) - 1) * 100 < -1 * PerTradeMaxLossPercentage Then
                                                runningTrade = New Trade(Me,
                                                                     currentMinuteCandlePayload.TradingSymbol,
                                                                     tradeStockType,
                                                                     currentMinuteCandlePayload.PayloadDate,
                                                                     Trade.TradeExecutionDirection.Sell,
                                                                     finalEntryPrice,
                                                                     entryBuffer,
                                                                     Trade.TradeType.MIS,
                                                                     Trade.TradeEntryCondition.Original,
                                                                     "Low ATR Band Stoploss Below VWAP",
                                                                     quantity,
                                                                     finalTargetPrice,
                                                                     finalTargetRemark,
                                                                     finalStoplossPrice,
                                                                     stoplossBuffer,
                                                                     finalStoplossRemark,
                                                                     currentMinuteCandlePayload)

                                                If openTradesToModifyOrCancel IsNot Nothing AndAlso openTradesToModifyOrCancel.Count > 0 Then
                                                    For Each openTrade In openTradesToModifyOrCancel
                                                        If openTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                            PlaceOrModifyOrder(openTrade, runningTrade)
                                                        ElseIf openTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                            CancelTrade(openTrade, currentMinuteCandlePayload, "Opposite Direction Signal Triggered")
                                                            PlaceOrModifyOrder(runningTrade, Nothing)
                                                        End If
                                                    Next
                                                Else
                                                    If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                                End If
                                            End If
                                        End If

                                        Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                            For Each potentialExitTrade In potentialExitTrades
                                                ExitTradeIfPossible(potentialExitTrade, tick, False)
                                            Next
                                        End If

                                        Dim potentialSLMoveTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialSLMoveTrades IsNot Nothing AndAlso potentialSLMoveTrades.Count > 0 Then
                                            For Each potentialSLMoveTrade In potentialSLMoveTrades
                                                If potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                    potentialStoploss = lowEntryPrice
                                                    finalStoplossPrice = Math.Max(potentialStoploss, potentialSLMoveTrade.PotentialStopLoss + CalculateBuffer(potentialSLMoveTrade.PotentialStopLoss, RoundOfType.Floor))
                                                    finalStoplossRemark = If(finalStoplossPrice = potentialStoploss, "Low ATR Band Stoploss Below VWAP", potentialSLMoveTrade.SLRemark)
                                                    stoplossBuffer = CalculateBuffer(finalStoplossPrice, RoundOfType.Floor)
                                                    finalStoplossPrice -= stoplossBuffer
                                                    potentialSLMoveTrade.UpdateTrade(StoplossBuffer:=stoplossBuffer)    'Buffer
                                                    MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, finalStoplossPrice, finalStoplossRemark)
                                                ElseIf potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                    potentialStoploss = highEntryPrice
                                                    finalStoplossPrice = Math.Min(potentialStoploss, potentialSLMoveTrade.PotentialStopLoss - CalculateBuffer(potentialSLMoveTrade.PotentialStopLoss, RoundOfType.Celing))
                                                    finalStoplossRemark = If(finalStoplossPrice = potentialStoploss, "High ATR Band Stoploss Above VWAP", potentialSLMoveTrade.SLRemark)
                                                    stoplossBuffer = CalculateBuffer(finalStoplossPrice, RoundOfType.Celing)
                                                    finalStoplossPrice += stoplossBuffer
                                                    potentialSLMoveTrade.UpdateTrade(StoplossBuffer:=stoplossBuffer)    'Buffer
                                                    MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, finalStoplossPrice, finalStoplossRemark)
                                                End If
                                            Next
                                        End If

                                        'Don't take Trade if PL for the day has been reached for the stock
                                        If IsStockAnyPLReachedForTheDay(currentMinuteCandlePayload, Trade.TradeType.MIS) Then
                                            ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Force Exit All Trades As Stock Max PL Reached")
                                            GoTo label1
                                        End If
                                        'Exit Form all stocks If Days PL Reached
                                        Dim allStockPLAfterBrokerage As Double = AllPLAfterBrokerage(tradeCheckingDate)
                                        If allStockPLAfterBrokerage > OverAllProfitPerDay OrElse allStockPLAfterBrokerage < OverAllLossPerDay OrElse startMinute >= EODExitTime Then
                                            Dim currentExitRemark As String = "Force Exit All Trades As Per Day Max PL Reached"
                                            If startMinute >= EODExitTime Then
                                                currentExitRemark = "EOD Force Exit"
                                            End If
                                            ExitAllTradeByForce(tick.PayloadDate, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, currentExitRemark)
                                            GoTo label2
                                        End If

                                        Dim lastTrade As Trade = GetLastSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Close)
                                        Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 AndAlso
                                            (lastTrade Is Nothing OrElse (lastTrade IsNot Nothing AndAlso lastTrade.ExitCondition <> Trade.TradeExitCondition.Target)) Then
                                            For Each potentialEntryTrade In potentialEntryTrades
                                                EnterTradeIfPossible(potentialEntryTrade, tick, False)
                                            Next
                                        End If
                                    Next
                                End If
label1:                     Next
                            startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                        End While   'Second Loop

                        startMinute = startMinute.Add(TimeSpan.FromMinutes(1))
                    End While   'Minute Loop
                End If
            End If
label2:     tradeCheckingDate = tradeCheckingDate.AddDays(1)
        End While   'Date Loop

        'Excel Printing
        Dim curDate As DateTime = System.DateTime.Now
        Dim filename As String = String.Format("In The Trend Strategy for {0} {1}-{2}-{3}_{4}-{5}-{6}.xlsx", "All Stock",
                                               curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
        PrintArrayToExcel(filename)
    End Function

    Private Function GetStockData(databaseTableType As Common.DataBaseTable, tradingDate As Date) As Dictionary(Of String, Double())
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        Dim dts As DataSet = Nothing
        Dim conn As MySqlConnection = cmn.OpenDBConnection
        Dim outputPayload As Dictionary(Of String, Double()) = Nothing

        If conn.State = ConnectionState.Open Then
            OnHeartbeat("Fetching Pre Market Volume Spike Data")
            Dim cmd As New MySqlCommand("GET_PRE_MARKET_VOLUME_SPIKE_PRE_MARKET_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", tradingDate)
            cmd.Parameters.AddWithValue("@endDate", tradingDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", NumberOfTradeableStockPerDay)
            cmd.Parameters.AddWithValue("@spikeChangePercentage", 100)
            cmd.Parameters.AddWithValue("@minClose", 100)
            cmd.Parameters.AddWithValue("@maxClose", 1500)
            cmd.Parameters.AddWithValue("@atrPercentage", 4)
            cmd.Parameters.AddWithValue("@potentialAmount", 1000000)
            cmd.Parameters.AddWithValue("@sortColumn", "QuantityXPrice")

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
                If dt Is Nothing Then dt = New DataTable
                dt.Merge(temp_dt)
                count += 1
            End While
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Double())
                    outputPayload.Add(dt.Rows(i).Item(1), {dt.Rows(i).Item(12)})
                Next
            End If
        End If
        Return outputPayload
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
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
