Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports MySql.Data.MySqlClient
Public Class FractalMACandleRetrace
    Inherits Strategy
    Implements IDisposable

    Dim cmn As Common = New Common(Canceller)
    Protected Property SMAPeriod As Integer = 50
    Public Property SpikeChangePercentageOfStockForPreMarketVolumeSpikeScreener As Double
    Public Property NumberOfDaysForNRxScreener As Integer
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan,
                   ByVal smaPeriod As Integer)
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
        Me.SMAPeriod = smaPeriod
    End Sub
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(1).ConfigureAwait(False)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        'TODO:
        'ChangeType
        Dim tradeStockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
        Dim databaseTable As Common.DataBaseTable = Common.DataBaseTable.Intraday_Cash

        Dim slPercentage As Double = 1 / 100
        Dim slabPercentage As Double = 1 / 100
        Dim breakevenPercentage As Double = slabPercentage / 10

        Dim tradeCheckingDate As Date = startDate
        While tradeCheckingDate <= endDate
            Dim stockList As Dictionary(Of String, Double()) = Nothing
            Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
            'TODO:
            'Change StockList
            'tempStocklist = New Dictionary(Of String, Double()) From {{"SUNTV", {1000}}, {"GSFC", {4500}}}
            'tempStocklist = GetPreMarketVolumeSpikePreMarketData(tradeCheckingDate)
            tempStocklist = GetNarrowRangeStockData(databaseTable, tradeCheckingDate)
            If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
                For Each tradingSymbol In tempStocklist.Keys
                    If stockList Is Nothing Then stockList = New Dictionary(Of String, Double())
                    stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0)})
                Next
            End If

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim OneMinuteHAPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim FractalWithSMASignalPayload As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim FractalWithSMAEntryPricePayload As Dictionary(Of String, Dictionary(Of Date, Double)) = Nothing
                Dim FractalWithSMAStoplossPricePayload As Dictionary(Of String, Dictionary(Of Date, Double)) = Nothing
                Dim FractalRetracementEntrySignalPayload As Dictionary(Of String, Dictionary(Of Date, Integer)) = Nothing
                Dim FractalRetracementEntryPricePayload As Dictionary(Of String, Dictionary(Of Date, Double)) = Nothing

                For Each stock In stockList.Keys
                    Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempOneMinutePayload As Dictionary(Of Date, Payload) = cmn.GetRawPayload(databaseTable, stock, tradeCheckingDate.AddDays(-5), tradeCheckingDate)
                    If tempOneMinutePayload IsNot Nothing AndAlso tempOneMinutePayload.Count > 0 Then
                        For Each tradingDate In tempOneMinutePayload.Keys
                            If tradingDate.Date = tradeCheckingDate.Date Then
                                If currentOneMinutePayload Is Nothing Then currentOneMinutePayload = New Dictionary(Of Date, Payload)
                                currentOneMinutePayload.Add(tradingDate, tempOneMinutePayload(tradingDate))
                            End If
                        Next
                    End If
                    If currentOneMinutePayload IsNot Nothing AndAlso currentOneMinutePayload.Count > 0 Then
                        OnHeartbeat(String.Format("Calculating Indicators for {0}", tradeCheckingDate.ToShortDateString))
                        Dim tempOneMinuteHAPayload As Dictionary(Of Date, Payload) = Nothing
                        Indicator.HeikenAshi.ConvertToHeikenAshi(tempOneMinutePayload, tempOneMinuteHAPayload)
                        Dim tempFractalWithSMASignalPayload As Dictionary(Of Date, Color) = Nothing
                        Dim tempFractalWithSMAEntryPayload As Dictionary(Of Date, Double) = Nothing
                        Dim tempFractalWithSMAStoplossPayload As Dictionary(Of Date, Double) = Nothing
                        Indicator.FractalWithSMA.CalculateFractalWithSMA(SMAPeriod, tempOneMinuteHAPayload, tempFractalWithSMASignalPayload, tempFractalWithSMAEntryPayload, tempFractalWithSMAStoplossPayload)
                        Dim tempFractalRetracementEntrySignalPayload As Dictionary(Of Date, Integer) = Nothing
                        Dim tempFractalRetracementEntryPricePayload As Dictionary(Of Date, Double) = Nothing
                        StrategyRule.FractalRetracementEntryRule.CalculateFractalRetracementEntry(SMAPeriod, tempOneMinuteHAPayload, tempFractalRetracementEntrySignalPayload, tempFractalRetracementEntryPricePayload)

                        If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        OneMinutePayload.Add(stock, currentOneMinutePayload)
                        If OneMinuteHAPayload Is Nothing Then OneMinuteHAPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        OneMinuteHAPayload.Add(stock, tempOneMinuteHAPayload)
                        If FractalWithSMASignalPayload Is Nothing Then FractalWithSMASignalPayload = New Dictionary(Of String, Dictionary(Of Date, Color))
                        FractalWithSMASignalPayload.Add(stock, tempFractalWithSMASignalPayload)
                        If FractalWithSMAEntryPricePayload Is Nothing Then FractalWithSMAEntryPricePayload = New Dictionary(Of String, Dictionary(Of Date, Double))
                        FractalWithSMAEntryPricePayload.Add(stock, tempFractalWithSMAEntryPayload)
                        If FractalWithSMAStoplossPricePayload Is Nothing Then FractalWithSMAStoplossPricePayload = New Dictionary(Of String, Dictionary(Of Date, Double))
                        FractalWithSMAStoplossPricePayload.Add(stock, tempFractalWithSMAStoplossPayload)
                        If FractalRetracementEntrySignalPayload Is Nothing Then FractalRetracementEntrySignalPayload = New Dictionary(Of String, Dictionary(Of Date, Integer))
                        FractalRetracementEntrySignalPayload.Add(stock, tempFractalRetracementEntrySignalPayload)
                        If FractalRetracementEntryPricePayload Is Nothing Then FractalRetracementEntryPricePayload = New Dictionary(Of String, Dictionary(Of Date, Double))
                        FractalRetracementEntryPricePayload.Add(stock, tempFractalRetracementEntryPricePayload)
                    End If
                Next

                If OneMinutePayload IsNot Nothing AndAlso OneMinutePayload.Count > 0 Then
                    OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                    Dim startMinute As TimeSpan = ExchangeStartTime
                    Dim endMinute As TimeSpan = ExchangeEndTime
                    Dim lastLossTradeEntry As Boolean = False
                    While startMinute < endMinute
                        Dim startSecond As TimeSpan = startMinute
                        Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromSeconds(59))
                        Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                        Dim currentCandlePayload As Payload = Nothing
                        Dim runningTrade As Trade = Nothing
                        While startSecond < endSecond
                            Dim potentialTickSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                            For Each stockName In stockList.Keys
                                Dim currentTickPayload As List(Of Payload) = Nothing
                                If OneMinutePayload.ContainsKey(stockName) AndAlso OneMinutePayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                    currentCandlePayload = OneMinutePayload(stockName)(potentialCandleSignalTime)
                                End If
                                If currentCandlePayload IsNot Nothing AndAlso currentCandlePayload.Ticks IsNot Nothing Then
                                    currentTickPayload = currentCandlePayload.Ticks.FindAll(Function(x)
                                                                                                Return x.PayloadDate = potentialTickSignalTime
                                                                                            End Function)
                                End If
                                If currentTickPayload IsNot Nothing AndAlso currentTickPayload.Count > 0 Then
                                    Dim entryPrice As Double = FractalRetracementEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                    Dim potentialSLPrice As Double = FractalWithSMAStoplossPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                    Dim correctSLRemark As String = Nothing
                                    Dim correctSLPrice As Double = Nothing
                                    Dim correctTargetRemark As String = "No fixed Target"
                                    If FractalRetracementEntrySignalPayload(stockName).ContainsKey(currentCandlePayload.PreviousCandlePayload.PayloadDate) AndAlso
                                        FractalRetracementEntrySignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = 1 AndAlso
                                        runningTrade Is Nothing Then 'To prevent this from opening a trade on every tick
                                        entryPrice += CalculateBuffer(entryPrice, NumberManipulation.RoundOfType.Celing)
                                        potentialSLPrice -= CalculateBuffer(potentialSLPrice, NumberManipulation.RoundOfType.Floor)
                                        If potentialSLPrice > entryPrice AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) > entryPrice Then
                                            potentialSLPrice = Double.MinValue
                                        End If
                                        correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Buy, Double.MinValue, Double.MinValue, (entryPrice - entryPrice * slPercentage), potentialSLPrice, Nothing, correctSLRemark)
                                        runningTrade = New Trade(Me,
                                                                 currentCandlePayload.TradingSymbol,
                                                                 tradeStockType,
                                                                 currentCandlePayload.PayloadDate,
                                                                 Trade.TradeExecutionDirection.Buy,
                                                                 entryPrice,
                                                                 Trade.TradeType.MIS,
                                                                 Trade.TradeEntryCondition.Original,
                                                                 "Candle Retracement Trade Entry",
                                                                 stockList(stockName)(0),
                                                                 Double.MaxValue,
                                                                 correctTargetRemark,
                                                                 correctSLPrice,
                                                                 correctSLRemark,
                                                                 currentCandlePayload)
                                        If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                    ElseIf FractalRetracementEntrySignalPayload(stockName).ContainsKey(currentCandlePayload.PreviousCandlePayload.PayloadDate) AndAlso
                                        FractalRetracementEntrySignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = -1 AndAlso
                                        runningTrade Is Nothing Then
                                        entryPrice -= CalculateBuffer(entryPrice, NumberManipulation.RoundOfType.Floor)
                                        potentialSLPrice += CalculateBuffer(potentialSLPrice, NumberManipulation.RoundOfType.Celing)
                                        If potentialSLPrice < entryPrice AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) < entryPrice Then
                                            potentialSLPrice = Double.MaxValue
                                        End If
                                        correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Sell, Double.MaxValue, Double.MaxValue, (entryPrice + entryPrice * slPercentage), potentialSLPrice, Nothing, correctSLRemark)
                                        runningTrade = New Trade(Me,
                                                                 currentCandlePayload.TradingSymbol,
                                                                 tradeStockType,
                                                                 currentCandlePayload.PayloadDate,
                                                                 Trade.TradeExecutionDirection.Sell,
                                                                 entryPrice,
                                                                 Trade.TradeType.MIS,
                                                                 Trade.TradeEntryCondition.Original,
                                                                 "Candle Retracement Trade Entry",
                                                                 stockList(stockName)(0),
                                                                 Double.MinValue,
                                                                 correctTargetRemark,
                                                                 correctSLPrice,
                                                                 correctSLRemark,
                                                                 currentCandlePayload)
                                        If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                    End If

                                    Dim lastTrade As Trade = GetLastSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Close)
                                    If GetLastSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open) Is Nothing AndAlso
                                        GetLastSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress) Is Nothing Then
                                        If lastTrade IsNot Nothing AndAlso lastTrade.PLAfterBrokerage < 0 Then lastLossTradeEntry = True
                                    End If
                                    If lastLossTradeEntry AndAlso lastTrade IsNot Nothing AndAlso lastTrade.PLAfterBrokerage < 0 AndAlso runningTrade Is Nothing Then
                                        lastLossTradeEntry = False
                                        If lastTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                            If FractalWithSMASignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = Color.Green Then
                                                entryPrice = FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                entryPrice += CalculateBuffer(entryPrice, NumberManipulation.RoundOfType.Celing)
                                                potentialSLPrice = FractalWithSMAStoplossPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                potentialSLPrice -= CalculateBuffer(potentialSLPrice, NumberManipulation.RoundOfType.Floor)
                                                If potentialSLPrice > entryPrice AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) > entryPrice Then
                                                    potentialSLPrice = Double.MinValue
                                                End If
                                                correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Buy, Double.MinValue, Double.MinValue, (entryPrice - entryPrice * slPercentage), potentialSLPrice, Nothing, correctSLRemark)
                                                runningTrade = New Trade(Me,
                                                             currentCandlePayload.TradingSymbol,
                                                             tradeStockType,
                                                             currentCandlePayload.PayloadDate,
                                                             Trade.TradeExecutionDirection.Buy,
                                                             entryPrice,
                                                             Trade.TradeType.MIS,
                                                             Trade.TradeEntryCondition.Onward,
                                                             "Fractal Entry After SL",
                                                             stockList(stockName)(0),
                                                             Double.MaxValue,
                                                             correctTargetRemark,
                                                             correctSLPrice,
                                                             correctSLRemark,
                                                             currentCandlePayload)
                                                If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                            End If
                                        ElseIf lastTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                            If FractalWithSMASignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = Color.Red Then
                                                entryPrice = FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                entryPrice -= CalculateBuffer(entryPrice, NumberManipulation.RoundOfType.Floor)
                                                potentialSLPrice = FractalWithSMAStoplossPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                potentialSLPrice += CalculateBuffer(potentialSLPrice, NumberManipulation.RoundOfType.Celing)
                                                If potentialSLPrice < entryPrice AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) < entryPrice Then
                                                    potentialSLPrice = Double.MaxValue
                                                End If
                                                correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Sell, Double.MaxValue, Double.MaxValue, (entryPrice + entryPrice * slPercentage), potentialSLPrice, Nothing, correctSLRemark)
                                                runningTrade = New Trade(Me,
                                                             currentCandlePayload.TradingSymbol,
                                                             tradeStockType,
                                                             currentCandlePayload.PayloadDate,
                                                             Trade.TradeExecutionDirection.Sell,
                                                             entryPrice,
                                                             Trade.TradeType.MIS,
                                                             Trade.TradeEntryCondition.Onward,
                                                             "Fractal Entry After SL",
                                                             stockList(stockName)(0),
                                                             Double.MinValue,
                                                             correctTargetRemark,
                                                             correctSLPrice,
                                                             correctSLRemark,
                                                             currentCandlePayload)
                                                If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                            End If
                                        End If
                                    End If

                                    'Check the ticks that got retieved in that second
                                    For Each tick In currentTickPayload
                                        Dim inProgressTrades As List(Of Trade) = GetSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If inProgressTrades IsNot Nothing AndAlso inProgressTrades.Count > 0 Then
                                            For Each inProgressTrade In inProgressTrades
                                                If ExitTradeIfPossible(inProgressTrade, tick) Then
                                                    Console.WriteLine("")
                                                End If
                                            Next
                                        End If
                                        Dim stoplossMoveTrades As List(Of Trade) = GetSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If stoplossMoveTrades IsNot Nothing AndAlso stoplossMoveTrades.Count > 0 Then
                                            For Each stoplossMoveTrade In stoplossMoveTrades
                                                correctSLPrice = 0
                                                potentialSLPrice = 0
                                                correctSLRemark = Nothing
                                                Dim ltpRemark As String = Nothing
                                                potentialSLPrice = FractalWithSMAStoplossPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                Dim ltpSLPrice As Double = CalculateMathematicalTrailingSL(stoplossMoveTrade.EntryPrice, stoplossMoveTrade.CurrentLTP, stoplossMoveTrade.EntryDirection, slabPercentage, breakevenPercentage, ltpRemark)
                                                If stoplossMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                    potentialSLPrice -= CalculateBuffer(potentialSLPrice, RoundOfType.Celing)
                                                    If potentialSLPrice >= stoplossMoveTrade.CurrentLTP AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) >= stoplossMoveTrade.CurrentLTP Then
                                                        potentialSLPrice = Double.MinValue
                                                    End If
                                                    If FractalWithSMASignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = Color.Red Then
                                                        potentialSLPrice = Double.MinValue
                                                    End If
                                                    correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Buy, stoplossMoveTrade.PotentialStopLoss, ltpSLPrice, stoplossMoveTrade.EntryPrice - stoplossMoveTrade.EntryPrice * slPercentage, potentialSLPrice, ltpRemark, correctSLRemark)
                                                ElseIf stoplossMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                    potentialSLPrice += CalculateBuffer(potentialSLPrice, RoundOfType.Floor)
                                                    If potentialSLPrice <= stoplossMoveTrade.CurrentLTP AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) <= stoplossMoveTrade.CurrentLTP Then
                                                        potentialSLPrice = Double.MaxValue
                                                    End If
                                                    If FractalWithSMASignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = Color.Green Then
                                                        potentialSLPrice = Double.MaxValue
                                                    End If
                                                    correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Sell, stoplossMoveTrade.PotentialStopLoss, ltpSLPrice, stoplossMoveTrade.EntryPrice + stoplossMoveTrade.EntryPrice * slPercentage, potentialSLPrice, ltpRemark, correctSLRemark)
                                                End If
                                                MoveStopLoss(stoplossMoveTrade, correctSLPrice, If(correctSLRemark = "", stoplossMoveTrade.SLRemark, correctSLRemark))
                                            Next
                                        End If
                                        Dim openTrades As List(Of Trade) = GetSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If openTrades IsNot Nothing AndAlso openTrades.Count > 0 Then
                                            For Each openTrade In openTrades
                                                Dim newTrade As Trade = Nothing
                                                If EnterTradeIfPossible(openTrade, tick) Then
                                                    Console.WriteLine("")
                                                ElseIf openTrade.EntryCondition = Trade.TradeEntryCondition.Onward Then
                                                    If openTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                        If FractalWithSMASignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = Color.Red Then
                                                            CancelTrade(openTrade, tick, "Opposite Direction Signal")
                                                        Else
                                                            entryPrice = FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                            entryPrice += CalculateBuffer(entryPrice, NumberManipulation.RoundOfType.Celing)
                                                            potentialSLPrice = FractalWithSMAStoplossPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                            potentialSLPrice -= CalculateBuffer(potentialSLPrice, NumberManipulation.RoundOfType.Floor)
                                                            If potentialSLPrice > entryPrice AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) > entryPrice Then
                                                                potentialSLPrice = Double.MinValue
                                                            End If
                                                            correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Buy, Double.MinValue, Double.MinValue, (entryPrice - entryPrice * slPercentage), potentialSLPrice, Nothing, correctSLRemark)
                                                            newTrade = New Trade(Me,
                                                                         currentCandlePayload.TradingSymbol,
                                                                         tradeStockType,
                                                                         currentCandlePayload.PayloadDate,
                                                                         Trade.TradeExecutionDirection.Buy,
                                                                         entryPrice,
                                                                         Trade.TradeType.MIS,
                                                                         Trade.TradeEntryCondition.Onward,
                                                                         "Fractal Entry After SL",
                                                                         stockList(stockName)(0),
                                                                         Double.MaxValue,
                                                                         correctTargetRemark,
                                                                         correctSLPrice,
                                                                         correctSLRemark,
                                                                         currentCandlePayload)
                                                            If newTrade IsNot Nothing Then PlaceOrModifyOrder(openTrade, newTrade)
                                                        End If
                                                    ElseIf openTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                        If FractalWithSMASignalPayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) = Color.Green Then
                                                            CancelTrade(openTrade, tick, "Opposite Direction Signal")
                                                        Else
                                                            entryPrice = FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                            entryPrice -= CalculateBuffer(entryPrice, NumberManipulation.RoundOfType.Floor)
                                                            potentialSLPrice = FractalWithSMAStoplossPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate)
                                                            potentialSLPrice += CalculateBuffer(potentialSLPrice, NumberManipulation.RoundOfType.Celing)
                                                            If potentialSLPrice < entryPrice AndAlso FractalWithSMAEntryPricePayload(stockName)(currentCandlePayload.PreviousCandlePayload.PayloadDate) < entryPrice Then
                                                                potentialSLPrice = Double.MaxValue
                                                            End If
                                                            correctSLPrice = GetSLWithRemark(Trade.TradeExecutionDirection.Sell, Double.MaxValue, Double.MaxValue, (entryPrice + entryPrice * slPercentage), potentialSLPrice, Nothing, correctSLRemark)
                                                            newTrade = New Trade(Me,
                                                                         currentCandlePayload.TradingSymbol,
                                                                         tradeStockType,
                                                                         currentCandlePayload.PayloadDate,
                                                                         Trade.TradeExecutionDirection.Sell,
                                                                         entryPrice,
                                                                         Trade.TradeType.MIS,
                                                                         Trade.TradeEntryCondition.Onward,
                                                                         "Fractal Entry After SL",
                                                                         stockList(stockName)(0),
                                                                         Double.MinValue,
                                                                         correctTargetRemark,
                                                                         correctSLPrice,
                                                                         correctSLRemark,
                                                                         currentCandlePayload)
                                                            If newTrade IsNot Nothing Then PlaceOrModifyOrder(openTrade, newTrade)
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If

                                If startSecond >= endSecond Then
                                    Dim potentialCancelTrades As List(Of Trade) = GetSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                    For Each potentialCancelTrade In potentialCancelTrades
                                        If potentialCancelTrade.EntryCondition = Trade.TradeEntryCondition.Original Then
                                            If potentialCancelTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                If FractalRetracementEntrySignalPayload(stockName).ContainsKey(currentCandlePayload.PayloadDate) AndAlso
                                                        FractalRetracementEntrySignalPayload(stockName)(currentCandlePayload.PayloadDate) = 1 Then
                                                    CancelTrade(potentialCancelTrade, currentCandlePayload, "New Entry Point Triggered")
                                                End If
                                            ElseIf potentialCancelTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                If FractalRetracementEntrySignalPayload(stockName).ContainsKey(currentCandlePayload.PayloadDate) AndAlso
                                                        FractalRetracementEntrySignalPayload(stockName)(currentCandlePayload.PayloadDate) = -1 Then
                                                    CancelTrade(potentialCancelTrade, currentCandlePayload, "New Entry Point Triggered")
                                                End If
                                            End If
                                        End If
                                    Next
                                End If

                                If startMinute = EODExitTime Then
                                    Dim cancelTrades As List(Of Trade) = Nothing
                                    cancelTrades = GetSpecificTrades(currentCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                    If cancelTrades IsNot Nothing AndAlso cancelTrades.Count > 0 Then
                                        For Each cancelTradeItem In cancelTrades
                                            CancelTrade(cancelTradeItem, currentCandlePayload, "EOD Cancel")
                                        Next
                                    End If
                                End If
                            Next
                            startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                        End While 'Within the minute, seconds loop

                        startMinute = startMinute.Add(TimeSpan.FromMinutes(1))
                    End While 'That minute loop
                End If
            End If
            tradeCheckingDate = tradeCheckingDate.AddDays(1)
        End While 'Date loop

        'Excel Printing
        Dim curDate As DateTime = System.DateTime.Now
        Dim filename As String = String.Format("Fractal With MA Retracement Candle Strategy for {0} {1}-{2}-{3}_{4}-{5}-{6}.xlsx", "All Stock",
                                               curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
        PrintArrayToExcel(filename)
    End Function
    Private Function GetSLWithRemark(ByVal tradeDirection As Trade.TradeExecutionDirection, ByVal previousSLPrice As Double, ByVal ltpSLPrice As Double, ByVal entrySLPrice As Double, ByVal fractalSLPrice As Double, ByVal ltpRemark As String, ByRef slRemark As String) As Double
        Dim slPrice As Double = Nothing
        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
            If previousSLPrice > ltpSLPrice AndAlso previousSLPrice > entrySLPrice AndAlso previousSLPrice > fractalSLPrice Then
                slPrice = previousSLPrice
                slRemark = ""
            ElseIf ltpSLPrice > previousSLPrice AndAlso ltpSLPrice > entrySLPrice AndAlso ltpSLPrice > fractalSLPrice Then
                slPrice = ltpSLPrice
                slRemark = ltpRemark
            ElseIf entrySLPrice > previousSLPrice AndAlso entrySLPrice > ltpSLPrice AndAlso entrySLPrice > fractalSLPrice Then
                slPrice = entrySLPrice
                slRemark = "Move To 1% of Entry"
            Else
                slPrice = fractalSLPrice
                slRemark = "Move To Fractral"
            End If
        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
            If previousSLPrice < ltpSLPrice AndAlso previousSLPrice < entrySLPrice AndAlso previousSLPrice < fractalSLPrice Then
                slPrice = previousSLPrice
                slRemark = ""
            ElseIf ltpSLPrice < previousSLPrice AndAlso ltpSLPrice < entrySLPrice AndAlso ltpSLPrice < fractalSLPrice Then
                slPrice = ltpSLPrice
                slRemark = ltpRemark
            ElseIf entrySLPrice < previousSLPrice AndAlso entrySLPrice < ltpSLPrice AndAlso entrySLPrice < fractalSLPrice Then
                slPrice = entrySLPrice
                slRemark = "Move To 1% of Entry"
            Else
                slPrice = fractalSLPrice
                slRemark = "Move To Fractral"
            End If
        End If
        Return slPrice
    End Function
    Private Function GetPreMarketVolumeSpikePreMarketData(tradingDate As Date) As Dictionary(Of String, Double())
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        Dim dts As DataSet = Nothing
        Dim conn As MySqlConnection = cmn.OpenDBConnection
        Dim outputPayload As Dictionary(Of String, Double()) = Nothing

        If conn.State = ConnectionState.Open Then
            OnHeartbeat(String.Format("Fetching Pre Market Volume Spike Data for {0}", tradingDate.ToShortDateString))
            Dim cmd As New MySqlCommand("GET_PRE_MARKET_VOLUME_SPIKE_PRE_MARKET_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", tradingDate)
            cmd.Parameters.AddWithValue("@endDate", tradingDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@spikeChangePercentage", SpikeChangePercentageOfStockForPreMarketVolumeSpikeScreener)
            cmd.Parameters.AddWithValue("@minClose", 100)
            cmd.Parameters.AddWithValue("@maxClose", 1500)
            cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
            cmd.Parameters.AddWithValue("@perMinuteLots", 15)
            cmd.Parameters.AddWithValue("@sortColumn", "VolumeChangePercentage")

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
                For i = 0 To NumberOfTradeableStock - 1
                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Double())
                    outputPayload.Add(dt.Rows(i).Item(1), {dt.Rows(i).Item(12)})
                Next
            End If
        End If
        Return outputPayload
    End Function
    Private Function GetNarrowRangeStockData(databaseTableType As Common.DataBaseTable, tradingDate As Date) As Dictionary(Of String, Double())
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        Dim dts As DataSet = Nothing
        Dim conn As MySqlConnection = cmn.OpenDBConnection
        Dim outputPayload As Dictionary(Of String, Double()) = Nothing
        Dim previousTradingDate As Date = cmn.GetPreviousTradingDay(databaseTableType, tradingDate)

        If conn.State = ConnectionState.Open Then
            OnHeartbeat(String.Format("Fetching Narrow Range Stock Data for {0}", tradingDate.ToShortDateString))
            Dim cmd As New MySqlCommand("GET_NARROW_RANGE_STOCK_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", previousTradingDate)
            cmd.Parameters.AddWithValue("@endDate", previousTradingDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@numberOfDays", NumberOfDaysForNRxScreener)
            cmd.Parameters.AddWithValue("@minClose", 100)
            cmd.Parameters.AddWithValue("@maxClose", 1500)
            cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
            cmd.Parameters.AddWithValue("@perMinuteLots", 15)

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
                Dim sortedDataTable As DataTable = Nothing
                dt.DefaultView.Sort = "ATRPercentage DESC"
                sortedDataTable = dt.DefaultView.ToTable
                For i = 0 To NumberOfTradeableStock - 1
                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Double())
                    outputPayload.Add(sortedDataTable.Rows(i).Item(1), {sortedDataTable.Rows(i).Item(5)})
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
