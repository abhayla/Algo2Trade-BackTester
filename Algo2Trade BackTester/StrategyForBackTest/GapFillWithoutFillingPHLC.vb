Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports MySql.Data.MySqlClient
Public Class GapFillWithoutFillingPHLC
    Inherits Strategy
    Implements IDisposable

    Dim capitalPerStock As Double = Nothing
    Dim totalPL As Double = 0
    Private Property _BreakevenPoint As Decimal
    Private Property _UseRevisedSlabSL As Boolean
    Private Property _Slab As Decimal
    Private Property _UsePreviousHighLow As Boolean
    Public Property OverAllCapital As Double
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan,
                   ByVal breakevenPoint As Decimal,
                   ByVal slab As Decimal,
                   ByVal useRevisedSlabSL As Boolean,
                   ByVal usePreviousHighLow As Boolean)
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
        _BreakevenPoint = breakevenPoint
        _Slab = slab
        _UseRevisedSlabSL = useRevisedSlabSL
        _UsePreviousHighLow = usePreviousHighLow
    End Sub
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(1).ConfigureAwait(False)
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        'TODO:
        'ChangeType
        Dim tradeStockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
        Dim databaseTable As Common.DataBaseTable = Common.DataBaseTable.Intraday_Futures

        Dim tradeCheckingDate As Date = startDate
        While tradeCheckingDate <= endDate
            Dim stockList As Dictionary(Of String, Double()) = Nothing
            Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
            'TODO:
            'Change StockList
            'tempStocklist = New Dictionary(Of String, Double()) From {{"JETAIRWAYS", {1, 1}}, {"YESBANK", {1, 1}}, {"ORIENTBANK", {1, 1}},{"CHOLAFIN", {1, 1}}, {"CANBK", {1, 1}}, {"BANKBARODA", {1, 1}}, {"CAPF", {1, 1}}, {"DLF", {1, 1}}}
            'tempStocklist = New Dictionary(Of String, Double()) From {{"MGL", {-1, 1}}, {"GAIL", {-1, 1}}}
            tempStocklist = GetRevisedStockData(databaseTable, tradeCheckingDate)

            If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
                For Each tradingSymbol In tempStocklist.Keys
                    If stockList Is Nothing Then stockList = New Dictionary(Of String, Double())
                    stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0), tempStocklist(tradingSymbol)(1)})
                Next
            End If

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                'capitalPerStock = OverAllCapital / stockList.Count
                Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing

                'First lets build the payload for all the stocks
                For Each stock In stockList.Keys
                    Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    'Get the currentDay payload
                    currentDayOneMinutePayload = Cmn.GetRawPayload(databaseTable, stock, tradeCheckingDate, tradeCheckingDate)
                    'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date isa valid date)

                    If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                        OnHeartbeat(String.Format("Processing for {0}", tradeCheckingDate.ToShortDateString))
                        'Add all these payloads into the stock collections
                        If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                    End If
                Next

                '------------------------------------------------------------------------------------------------------------------------------------------------

                If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                    OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                    Dim startMinute As TimeSpan = ExchangeStartTime
                    Dim endMinute As TimeSpan = ExchangeEndTime
                    While startMinute < endMinute
                        Dim onceInADay As Boolean = False
                        If startMinute = ExchangeStartTime Then
                            Dim addMinute As TimeSpan = TimeSpan.Parse("00:05:00")
                            startMinute = startMinute.Add(addMinute)
                            onceInADay = True
                        End If

                        Dim startSecond As TimeSpan = startMinute
                        Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromSeconds(59))
                        Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                        Dim potentialTickSignalTime As Date = Nothing
                        Dim currentMinuteCandlePayload As Payload = Nothing

                        While startSecond < endSecond
                            potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                            For Each stockName In stockList.Keys
                                Dim oncePerStock As Boolean = True
                                Dim runningTrade As Trade = Nothing
                                Dim previousHigh As Decimal = Nothing
                                Dim previousLow As Decimal = Nothing
                                'Get the current minute candle from the stock collection for this stock for that day
                                Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                    currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialCandleSignalTime)
                                End If

                                If onceInADay AndAlso oncePerStock AndAlso currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentMinuteCandlePayload IsNot Nothing Then
                                    oncePerStock = False
                                    previousHigh = currentDayOneMinuteStocksPayload(stockName).Max(Function(x)
                                                                                                       If x.Key < currentMinuteCandlePayload.PayloadDate Then
                                                                                                           Return x.Value.High
                                                                                                       Else
                                                                                                           Return Decimal.MinValue
                                                                                                       End If
                                                                                                   End Function)

                                    previousLow = currentDayOneMinuteStocksPayload(stockName).Min(Function(x)
                                                                                                      If x.Key < currentMinuteCandlePayload.PayloadDate Then
                                                                                                          Return x.Value.Low
                                                                                                      Else
                                                                                                          Return Decimal.MaxValue
                                                                                                      End If
                                                                                                  End Function)
                                End If

                                Dim lastTrade As Trade = GetLastSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Close)

                                'Now get the ticks for this minute and second
                                If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.Ticks IsNot Nothing AndAlso
                                    Not IsStockAnyPLReachedForTheDay(currentMinuteCandlePayload, Trade.TradeType.MIS) AndAlso
                                    (lastTrade Is Nothing OrElse lastTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Close) Then
                                    currentSecondTickPayload = currentMinuteCandlePayload.Ticks.FindAll(Function(x)
                                                                                                            Return x.PayloadDate = potentialTickSignalTime
                                                                                                        End Function)
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
                                    Dim quantity As Integer = stockList(stockName)(1)
                                    Dim entryBuffer As Decimal = Nothing
                                    Dim tempStoplossBuffer As Decimal = Nothing
                                    Dim stoplossBuffer As Decimal = Nothing

                                    If Not IsTradeOpen(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                                        Not IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                                            stockList(stockName)(0) = 1 Then

                                        finalEntryPrice = currentMinuteCandlePayload.Open
                                        entryBuffer = 0
                                        finalEntryPrice += entryBuffer

                                        potentialTargetOnMaxPLForStock = finalEntryPrice + finalEntryPrice * PerTradeMaxProfitPercentage / 100
                                        finalTargetPrice = potentialTargetOnMaxPLForStock
                                        finalTargetRemark = "Max Profit% For Stock"

                                        If _UsePreviousHighLow Then
                                            tempStoplossBuffer = CalculateBuffer(previousLow, NumberManipulation.RoundOfType.Floor)
                                            If previousLow - tempStoplossBuffer > finalEntryPrice Then
                                                potentialStoploss = finalEntryPrice + finalEntryPrice * PerTradeMaxLossPercentage / 100
                                                finalStoplossRemark = "Max Loss% For Stock"
                                                tempStoplossBuffer = 0
                                            Else
                                                potentialStoploss = previousLow - tempStoplossBuffer
                                                finalStoplossRemark = "Candle Low"
                                            End If
                                        Else
                                            potentialStoploss = finalEntryPrice + finalEntryPrice * PerTradeMaxLossPercentage / 100
                                            finalStoplossRemark = "Max Loss% For Stock"
                                            tempStoplossBuffer = 0
                                        End If
                                        finalStoplossPrice = potentialStoploss
                                        stoplossBuffer = tempStoplossBuffer

                                        'quantity = CalculateQuantityFromInvestment(quantity, capitalPerStock, finalEntryPrice, tradeStockType)

                                        runningTrade = New Trade(Me,
                                                                 currentMinuteCandlePayload.TradingSymbol,
                                                                 tradeStockType,
                                                                 currentMinuteCandlePayload.PayloadDate,
                                                                 Trade.TradeExecutionDirection.Buy,
                                                                 finalEntryPrice,
                                                                 entryBuffer,
                                                                 Trade.TradeType.MIS,
                                                                 Trade.TradeEntryCondition.Original,
                                                                 "Signal Candle Open Price",
                                                                 quantity,
                                                                 finalTargetPrice,
                                                                 finalTargetRemark,
                                                                 finalStoplossPrice,
                                                                 stoplossBuffer,
                                                                 finalStoplossRemark,
                                                                 currentMinuteCandlePayload)

                                        If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                    ElseIf Not IsTradeOpen(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                                        Not IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                                                stockList(stockName)(0) = -1 Then

                                        finalEntryPrice = currentMinuteCandlePayload.Open
                                        entryBuffer = 0
                                        finalEntryPrice += entryBuffer

                                        potentialTargetOnMaxPLForStock = finalEntryPrice - finalEntryPrice * PerTradeMaxProfitPercentage / 100
                                        finalTargetPrice = potentialTargetOnMaxPLForStock
                                        finalTargetRemark = "Max Profit% For Stock"

                                        If _UsePreviousHighLow Then
                                            tempStoplossBuffer = CalculateBuffer(previousHigh, NumberManipulation.RoundOfType.Celing)
                                            If previousHigh + tempStoplossBuffer < finalEntryPrice Then
                                                potentialStoploss = finalEntryPrice - finalEntryPrice * PerTradeMaxLossPercentage / 100
                                                finalStoplossRemark = "Max Loss% For Stock"
                                                tempStoplossBuffer = 0
                                            Else
                                                potentialStoploss = previousHigh + tempStoplossBuffer
                                                finalStoplossRemark = "Candle High"
                                            End If
                                        Else
                                            potentialStoploss = finalEntryPrice - finalEntryPrice * PerTradeMaxLossPercentage / 100
                                            finalStoplossRemark = "Max Loss% For Stock"
                                            tempStoplossBuffer = 0
                                        End If
                                        finalStoplossPrice = potentialStoploss
                                        stoplossBuffer = tempStoplossBuffer

                                        'quantity = CalculateQuantityFromInvestment(quantity, capitalPerStock, finalEntryPrice, tradeStockType)

                                        runningTrade = New Trade(Me,
                                                                currentMinuteCandlePayload.TradingSymbol,
                                                                tradeStockType,
                                                                currentMinuteCandlePayload.PayloadDate,
                                                                Trade.TradeExecutionDirection.Sell,
                                                                finalEntryPrice,
                                                                entryBuffer,
                                                                Trade.TradeType.MIS,
                                                                Trade.TradeEntryCondition.Original,
                                                                "Signal Candle Open Price",
                                                                quantity,
                                                                finalTargetPrice,
                                                                finalTargetRemark,
                                                                finalStoplossPrice,
                                                                stoplossBuffer,
                                                                finalStoplossRemark,
                                                                currentMinuteCandlePayload)

                                        If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                    End If

                                    For Each tick In currentSecondTickPayload
                                        SetCurrentLTPForStock(currentMinuteCandlePayload, tick, Trade.TradeType.MIS)

                                        Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                            For Each potentialExitTrade In potentialExitTrades
                                                If ExitTradeIfPossible(potentialExitTrade, tick) Then
                                                    Console.WriteLine("")
                                                End If
                                            Next
                                        End If

                                        Dim potentialSLMoveTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialSLMoveTrades IsNot Nothing AndAlso potentialSLMoveTrades.Count > 0 Then
                                            For Each potentialSLMoveTrade In potentialSLMoveTrades
                                                If _UseRevisedSlabSL Then
                                                    potentialStoploss = CalculateRevisedMathematicalTrailingSL(potentialSLMoveTrade.EntryPrice, potentialSLMoveTrade.CurrentLTP, potentialSLMoveTrade.EntryDirection, _Slab, _BreakevenPoint, finalStoplossRemark)
                                                Else
                                                    potentialStoploss = CalculateMathematicalTrailingSL(potentialSLMoveTrade.EntryPrice, potentialSLMoveTrade.CurrentLTP, potentialSLMoveTrade.EntryDirection, _Slab, _BreakevenPoint, finalStoplossRemark)
                                                End If

                                                If potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                                                    potentialSLMoveTrade.PotentialStopLoss < potentialStoploss Then
                                                    finalStoplossPrice = potentialStoploss
                                                    stoplossBuffer = 0
                                                    finalStoplossPrice -= stoplossBuffer
                                                    potentialSLMoveTrade.UpdateTrade(StoplossBuffer:=stoplossBuffer)    'Buffer
                                                    MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, finalStoplossPrice, finalStoplossRemark)
                                                ElseIf potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                                                    potentialSLMoveTrade.PotentialStopLoss > potentialStoploss Then
                                                    finalStoplossPrice = potentialStoploss
                                                    stoplossBuffer = 0
                                                    finalStoplossPrice -= stoplossBuffer
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

                                        Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                            For Each potentialEntryTrade In potentialEntryTrades
                                                If EnterTradeIfPossible(potentialEntryTrade, tick) Then
                                                    Console.WriteLine("")
                                                End If
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
            totalPL += AllPLAfterBrokerage(tradeCheckingDate)
            tradeCheckingDate = tradeCheckingDate.AddDays(1)
        End While   'Date Loop

        'Excel Printing
        Dim curDate As DateTime = System.DateTime.Now
        Dim filename As String = String.Format("Gap Fill Without Filling PHLC Strategy PL-{0} Stocks-{1} SL-{2} {3}-{4}-{5}_{6}-{7}-{8}.xlsx",
                                               Math.Round(totalPL, 0), NumberOfTradeableStockPerDay, If(_UseRevisedSlabSL, "Continuous", "Slab"),
                                               curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
        PrintArrayToExcel(filename)
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
                        Dim signalDirection As Integer = If(dt.Rows(i).Item(5) >= 0, -1, 1)
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
        Dim signalTime As TimeSpan = TimeSpan.Parse("00:05:00")
        signalTime = ExchangeStartTime.Add(signalTime)
        Dim potentialCandleSignalTime As Date = New Date(tradingDate.Year, tradingDate.Month, tradingDate.Day, signalTime.Hours, signalTime.Minutes, signalTime.Seconds)
        If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
            Dim count As Integer = 0
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

                    Dim chkValue As Decimal = Nothing
                    If firstCandle.Open > eodPayload.LastOrDefault.Value.High Then
                        chkValue = eodPayload.LastOrDefault.Value.High
                    ElseIf firstCandle.Open < eodPayload.LastOrDefault.Value.Low Then
                        chkValue = eodPayload.LastOrDefault.Value.Low
                    Else
                        chkValue = eodPayload.LastOrDefault.Value.Close
                    End If

                    If tempStocklist(stock)(0) = 1 Then
                        If previousHigh < chkValue Then
                            If ret Is Nothing Then ret = New Dictionary(Of String, Double())
                            ret.Add(stock, tempStocklist(stock))
                            count += 1
                        End If
                    ElseIf tempStocklist(stock)(0) = -1 Then
                        If previousLow > chkValue Then
                            If ret Is Nothing Then ret = New Dictionary(Of String, Double())
                            ret.Add(stock, tempStocklist(stock))
                            count += 1
                        End If
                    End If
                End If
                If count = NumberOfTradeableStockPerDay Then Exit For
            Next
        End If
        Return ret
    End Function
    Private Function GetRevisedStockData(databaseTableType As Common.DataBaseTable, tradingDate As Date) As Dictionary(Of String, Double())
        Dim ret As Dictionary(Of String, Double()) = Nothing
        Dim tempStocklist As Dictionary(Of String, Double()) = GetModifiedStockData(databaseTableType, tradingDate)
        Dim tempSortedStockCapital As SortedList(Of Double, String) = Nothing
        If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
            For Each stock In tempStocklist.Keys
                Dim requiredCapital As Double = Math.Round((tempStocklist(stock)(1) * tempStocklist(stock)(3)) / MarginMultiplier, 2)
                If tempSortedStockCapital Is Nothing Then tempSortedStockCapital = New SortedList(Of Double, String)
                tempSortedStockCapital.Add(requiredCapital, stock)
            Next

            Dim stockFutureQuantity As Dictionary(Of String, Integer) = Nothing
            Dim capitalLeft As Double = OverAllCapital
            While True
                Dim atleastOneFound As Boolean = False
                For sortedCapitalIndex As Integer = 0 To tempSortedStockCapital.Count - 1
                    If capitalLeft > tempSortedStockCapital.ElementAt(sortedCapitalIndex).Key Then
                        atleastOneFound = True

                        If stockFutureQuantity Is Nothing Then stockFutureQuantity = New Dictionary(Of String, Integer)
                        If stockFutureQuantity.ContainsKey(tempSortedStockCapital.ElementAt(sortedCapitalIndex).Value) Then
                            stockFutureQuantity(tempSortedStockCapital.ElementAt(sortedCapitalIndex).Value) += 1
                        Else
                            stockFutureQuantity.Add(tempSortedStockCapital.ElementAt(sortedCapitalIndex).Value, 1)
                        End If
                        capitalLeft -= tempSortedStockCapital.ElementAt(sortedCapitalIndex).Key
                    End If
                Next
                If Not atleastOneFound = True Then Exit While
            End While

            For Each stock In stockFutureQuantity.Keys
                If ret Is Nothing Then ret = New Dictionary(Of String, Double())
                ret.Add(stock, {tempStocklist(stock)(0), tempStocklist(stock)(1) * stockFutureQuantity(stock), tempStocklist(stock)(2)})
            Next
        End If
        Return ret
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
