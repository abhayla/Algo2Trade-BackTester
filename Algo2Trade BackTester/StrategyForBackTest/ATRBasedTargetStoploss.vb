Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading

Public Class ATRBasedTargetStoploss
    Inherits Strategy
    Implements IDisposable

    Private _signalTimeFrame As Integer
    Public Property TargetMultiplier As Decimal = 1
    Public Property StoplossMultiplier As Decimal = 1
    Public Property PartialExit As Boolean = False
    Public Property NumberOfTradePerDay As Integer = 1
    Public Property TrailingSL As Boolean = True
    Public Property ReverseSignalTrade As Boolean = True
    Public Property ExitOnFixedTargetStoploss As Boolean = False
    Private _maxProfitPerDay As Double = 15000
    Private _maxLossPerDay As Double = -10000

    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan,
                   ByVal signalTimeFrame As Integer)
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
        Me._signalTimeFrame = signalTimeFrame
    End Sub
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tradeStockType As Trade.TypeOfStock = Trade.TypeOfStock.Commodity
        Dim databaseTable As Common.DataBaseTable = Common.DataBaseTable.Intraday_Commodity
        Dim totalPL As Decimal = 0
        Dim tradeCheckingDate As Date = startDate
        While tradeCheckingDate <= endDate
            Dim stockList As List(Of String) = Nothing
            'TODO: Change Stocklist
            stockList = New List(Of String) From {"NATURALGAS"}

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim currentDayXMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim ATRXMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim XDayRuleSignalStocksPayload As Dictionary(Of String, Dictionary(Of Date, Integer)) = Nothing
                Dim XDayRuleEntryStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                'First lets build the payload for all the stocks
                For Each stock In stockList
                    Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim XDayXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim XDayXMinuteHAPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim ATRXMinutePayload As Dictionary(Of Date, Decimal) = Nothing
                    Dim XDayRuleSignalPayload As Dictionary(Of Date, Integer) = Nothing
                    Dim XDayRuleEntryPayload As Dictionary(Of Date, Decimal) = Nothing
                    Dim XDayRuleNextCandleExecutionPayload As Dictionary(Of Date, Boolean) = Nothing
                    'Get the currentDay payload
                    XDayOneMinutePayload = Cmn.GetRawPayload(databaseTable, stock, tradeCheckingDate.AddDays(-7), tradeCheckingDate)
                    'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)

                    If XDayOneMinutePayload IsNot Nothing AndAlso XDayOneMinutePayload.Count > 0 Then
                        OnHeartbeat(String.Format("Processing for {0}", tradeCheckingDate.ToShortDateString))
                        For Each runningPayload In XDayOneMinutePayload.Keys
                            If runningPayload.Date = tradeCheckingDate.Date Then
                                If currentDayOneMinutePayload Is Nothing Then currentDayOneMinutePayload = New Dictionary(Of Date, Payload)
                                currentDayOneMinutePayload.Add(runningPayload, XDayOneMinutePayload(runningPayload))
                            End If
                        Next
                        'Add all these payloads into the stock collections
                        If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                            If _signalTimeFrame > 1 Then
                                XDayXMinutePayload = Cmn.ConvertPayloadsToXMinutes(XDayOneMinutePayload, _signalTimeFrame)
                            Else
                                XDayXMinutePayload = currentDayOneMinutePayload
                            End If
                            Indicator.ATR.CalculateATR(14, XDayXMinutePayload, ATRXMinutePayload)
                            Indicator.HeikenAshi.ConvertToHeikenAshi(XDayXMinutePayload, XDayXMinuteHAPayload)

                            If XDayXMinuteHAPayload IsNot Nothing AndAlso XDayXMinuteHAPayload.Count > 0 Then
                                StrategyRules.HeikenAshiReversalRule.CalculateHeikenAshiReversalRule(XDayXMinuteHAPayload, XDayRuleSignalPayload, XDayRuleEntryPayload, XDayRuleNextCandleExecutionPayload)
                            End If

                            If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                            If currentDayXMinuteStocksPayload Is Nothing Then currentDayXMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            currentDayXMinuteStocksPayload.Add(stock, XDayXMinutePayload)
                            If ATRXMinuteStocksPayload Is Nothing Then ATRXMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            ATRXMinuteStocksPayload.Add(stock, ATRXMinutePayload)
                            If XDayRuleSignalStocksPayload Is Nothing Then XDayRuleSignalStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Integer))
                            XDayRuleSignalStocksPayload.Add(stock, XDayRuleSignalPayload)
                            If XDayRuleEntryStocksPayload Is Nothing Then XDayRuleEntryStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            XDayRuleEntryStocksPayload.Add(stock, XDayRuleEntryPayload)
                        End If
                    End If
                Next

                '------------------------------------------------------------------------------------------------------------------------------------------------

                If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                    OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                    Dim startMinute As TimeSpan = ExchangeStartTime
                    Dim endMinute As TimeSpan = ExchangeEndTime
                    While startMinute < endMinute
                        Dim startSecond As TimeSpan = startMinute
                        Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromMinutes(_signalTimeFrame - 1))
                        endSecond = endSecond.Add(TimeSpan.FromSeconds(59))
                        Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                        Dim potentialTickSignalTime As Date = Nothing
                        Dim currentMinuteCandlePayload As Payload = Nothing
                        Dim signalCandleTime As Date = Nothing

                        While startSecond <= endSecond
                            potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                            For Each stockName In stockList
                                Dim runningTrade As Trade = Nothing
                                Dim runningTrade2 As Trade = Nothing
                                'Get the current minute candle from the stock collection for this stock for that day
                                Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialTickSignalTime) Then
                                    currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialTickSignalTime)
                                End If
                                signalCandleTime = potentialTickSignalTime.AddMinutes(-_signalTimeFrame)
                                'Now get the ticks for this minute and second
                                If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.Ticks IsNot Nothing Then
                                    currentSecondTickPayload = currentMinuteCandlePayload.Ticks.FindAll(Function(x)
                                                                                                            Return x.PayloadDate = potentialTickSignalTime
                                                                                                        End Function)
                                End If

                                'Main Strategy Logic
                                Dim finalEntryPrice As Double = Nothing
                                Dim finalTargetPrice As Double = Nothing
                                Dim finalTargetRemark As String = Nothing
                                Dim finalStoplossPrice As Double = Nothing
                                Dim finalStoplossRemark As String = Nothing
                                Dim quantity As Integer = Nothing
                                Dim lotSize As Integer = Nothing
                                Dim entryBuffer As Decimal = Nothing
                                Dim stoplossBuffer As Decimal = Nothing

                                Dim tradeActive As Boolean = False

                                If currentMinuteCandlePayload IsNot Nothing AndAlso NumberOfTradesPerStockPerDay(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol) < NumberOfTradePerDay AndAlso
                                   XDayRuleSignalStocksPayload.ContainsKey(stockName) AndAlso
                                   XDayRuleSignalStocksPayload(stockName).ContainsKey(signalCandleTime) AndAlso
                                   XDayRuleSignalStocksPayload(stockName)(signalCandleTime) = 1 Then

                                    If ReverseSignalTrade Then
                                        tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Buy)
                                    Else
                                        tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                    End If

                                    If Not tradeActive Then
                                        finalEntryPrice = XDayRuleEntryStocksPayload(stockName)(signalCandleTime)
                                        entryBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)
                                        finalEntryPrice += entryBuffer

                                        If TrailingSL Then
                                            finalTargetPrice = finalEntryPrice + 10
                                        Else
                                            finalTargetPrice = finalEntryPrice + ATRXMinuteStocksPayload(stockName)(signalCandleTime) * TargetMultiplier
                                        End If
                                        finalTargetRemark = String.Format("Target: {0}", Math.Round(finalTargetPrice - finalEntryPrice, 2))

                                        finalStoplossPrice = finalEntryPrice - ATRXMinuteStocksPayload(stockName)(signalCandleTime) * StoplossMultiplier
                                        stoplossBuffer = 0
                                        finalStoplossPrice -= stoplossBuffer
                                        finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalEntryPrice - finalStoplossPrice, 2))

                                        quantity = 1250 * 10
                                        lotSize = 1250

                                        Dim tradeTag As String = System.Guid.NewGuid.ToString()

                                        If PartialExit Then
                                            Dim modifiedQuantity As Integer = quantity / 2
                                            runningTrade = New Trade(Me,
                                                         currentMinuteCandlePayload.TradingSymbol,
                                                         tradeStockType,
                                                         currentMinuteCandlePayload.PayloadDate,
                                                         Trade.TradeExecutionDirection.Buy,
                                                         finalEntryPrice,
                                                         entryBuffer,
                                                         Trade.TradeType.MIS,
                                                         Trade.TradeEntryCondition.Original,
                                                         "Signal Candle High Price",
                                                         modifiedQuantity,
                                                         lotSize,
                                                         finalTargetPrice,
                                                         finalTargetRemark,
                                                         finalStoplossPrice,
                                                         stoplossBuffer,
                                                         finalStoplossRemark,
                                                         currentDayOneMinuteStocksPayload(stockName)(signalCandleTime))

                                            finalTargetPrice = finalEntryPrice + 10
                                            finalTargetRemark = String.Format("Target: 10")

                                            runningTrade2 = New Trade(Me,
                                                         currentMinuteCandlePayload.TradingSymbol,
                                                         tradeStockType,
                                                         currentMinuteCandlePayload.PayloadDate,
                                                         Trade.TradeExecutionDirection.Buy,
                                                         finalEntryPrice,
                                                         entryBuffer,
                                                         Trade.TradeType.MIS,
                                                         Trade.TradeEntryCondition.Original,
                                                         "Signal Candle High Price",
                                                         quantity - modifiedQuantity,
                                                         lotSize,
                                                         finalTargetPrice,
                                                         finalTargetRemark,
                                                         finalStoplossPrice,
                                                         stoplossBuffer,
                                                         finalStoplossRemark,
                                                         currentDayOneMinuteStocksPayload(stockName)(signalCandleTime))

                                            runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(signalCandleTime))
                                            runningTrade2.UpdateTrade(Tag:=tradeTag, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(signalCandleTime))
                                        Else
                                            runningTrade = New Trade(Me,
                                                         currentMinuteCandlePayload.TradingSymbol,
                                                         tradeStockType,
                                                         currentMinuteCandlePayload.PayloadDate,
                                                         Trade.TradeExecutionDirection.Buy,
                                                         finalEntryPrice,
                                                         entryBuffer,
                                                         Trade.TradeType.MIS,
                                                         Trade.TradeEntryCondition.Original,
                                                         "Signal Candle High Price",
                                                         quantity,
                                                         lotSize,
                                                         finalTargetPrice,
                                                         finalTargetRemark,
                                                         finalStoplossPrice,
                                                         stoplossBuffer,
                                                         finalStoplossRemark,
                                                         currentDayOneMinuteStocksPayload(stockName)(signalCandleTime))

                                            runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(signalCandleTime))
                                        End If
                                        Dim potentialOpenTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialOpenTrades IsNot Nothing AndAlso potentialOpenTrades.Count > 0 Then
                                            For Each potentialOpenTrade In potentialOpenTrades
                                                If potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                    If PartialExit Then
                                                        If potentialOpenTrade.PotentialTarget - potentialOpenTrade.EntryPrice <= potentialOpenTrade.SquareOffValue - potentialOpenTrade.EntryPrice Then
                                                            If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                        ElseIf potentialOpenTrade.PotentialTarget - potentialOpenTrade.EntryPrice > potentialOpenTrade.SquareOffValue - potentialOpenTrade.EntryPrice Then
                                                            If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade2)
                                                        End If
                                                    Else
                                                        If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                    End If
                                                ElseIf potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                    Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                    dummyPayload.PayloadDate = potentialTickSignalTime
                                                    CancelTrade(potentialOpenTrade, dummyPayload, "Opposite Direction Signal")
                                                End If
                                            Next
                                        Else
                                            If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                            If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                        End If
                                    End If
                                ElseIf currentMinuteCandlePayload IsNot Nothing AndAlso NumberOfTradesPerStockPerDay(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol) < NumberOfTradePerDay AndAlso
                                    XDayRuleSignalStocksPayload.ContainsKey(stockName) AndAlso
                                    XDayRuleSignalStocksPayload(stockName).ContainsKey(signalCandleTime) AndAlso
                                    XDayRuleSignalStocksPayload(stockName)(signalCandleTime) = -1 Then

                                    If ReverseSignalTrade Then
                                        tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Sell)
                                    Else
                                        tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                    End If

                                    If Not tradeActive Then
                                        finalEntryPrice = XDayRuleEntryStocksPayload(stockName)(signalCandleTime)
                                        entryBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)
                                        finalEntryPrice -= entryBuffer

                                        If TrailingSL Then
                                            finalTargetPrice = finalEntryPrice - 10
                                        Else
                                            finalTargetPrice = finalEntryPrice - ATRXMinuteStocksPayload(stockName)(signalCandleTime) * TargetMultiplier
                                        End If
                                        finalTargetRemark = String.Format("Target: {0}", Math.Round(finalEntryPrice - finalTargetPrice, 2))

                                        finalStoplossPrice = finalEntryPrice + ATRXMinuteStocksPayload(stockName)(signalCandleTime) * StoplossMultiplier
                                        stoplossBuffer = 0
                                        finalStoplossPrice -= stoplossBuffer
                                        finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalStoplossPrice - finalEntryPrice, 2))

                                        quantity = 1250 * 10
                                        lotSize = 1250

                                        Dim tradeTag As String = System.Guid.NewGuid.ToString()

                                        If PartialExit Then
                                            Dim modifiedQuantity As Integer = quantity / 2
                                            runningTrade = New Trade(Me,
                                                            currentMinuteCandlePayload.TradingSymbol,
                                                            tradeStockType,
                                                            currentMinuteCandlePayload.PayloadDate,
                                                            Trade.TradeExecutionDirection.Sell,
                                                            finalEntryPrice,
                                                            entryBuffer,
                                                            Trade.TradeType.MIS,
                                                            Trade.TradeEntryCondition.Original,
                                                            "Signal Candle Low Price",
                                                            modifiedQuantity,
                                                            lotSize,
                                                            finalTargetPrice,
                                                            finalTargetRemark,
                                                            finalStoplossPrice,
                                                            stoplossBuffer,
                                                            finalStoplossRemark,
                                                            currentDayOneMinuteStocksPayload(stockName)(signalCandleTime))

                                            finalTargetPrice = finalEntryPrice - 10
                                            finalTargetRemark = String.Format("Target: 10")

                                            runningTrade2 = New Trade(Me,
                                                           currentMinuteCandlePayload.TradingSymbol,
                                                           tradeStockType,
                                                           currentMinuteCandlePayload.PayloadDate,
                                                           Trade.TradeExecutionDirection.Sell,
                                                           finalEntryPrice,
                                                           entryBuffer,
                                                           Trade.TradeType.MIS,
                                                           Trade.TradeEntryCondition.Original,
                                                           "Signal Candle Low Price",
                                                           quantity - modifiedQuantity,
                                                           lotSize,
                                                           finalTargetPrice,
                                                           finalTargetRemark,
                                                           finalStoplossPrice,
                                                           stoplossBuffer,
                                                           finalStoplossRemark,
                                                           currentDayOneMinuteStocksPayload(stockName)(signalCandleTime))

                                            runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(signalCandleTime))
                                            runningTrade2.UpdateTrade(Tag:=tradeTag, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(signalCandleTime))
                                        Else
                                            runningTrade = New Trade(Me,
                                                            currentMinuteCandlePayload.TradingSymbol,
                                                            tradeStockType,
                                                            currentMinuteCandlePayload.PayloadDate,
                                                            Trade.TradeExecutionDirection.Sell,
                                                            finalEntryPrice,
                                                            entryBuffer,
                                                            Trade.TradeType.MIS,
                                                            Trade.TradeEntryCondition.Original,
                                                            "Signal Candle Low Price",
                                                            quantity,
                                                            lotSize,
                                                            finalTargetPrice,
                                                            finalTargetRemark,
                                                            finalStoplossPrice,
                                                            stoplossBuffer,
                                                            finalStoplossRemark,
                                                            currentDayOneMinuteStocksPayload(stockName)(signalCandleTime))

                                            runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(signalCandleTime))
                                        End If
                                        Dim potentialOpenTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialOpenTrades IsNot Nothing AndAlso potentialOpenTrades.Count > 0 Then
                                            For Each potentialOpenTrade In potentialOpenTrades
                                                If potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                    If PartialExit Then
                                                        If potentialOpenTrade.EntryPrice - potentialOpenTrade.PotentialTarget <= potentialOpenTrade.EntryPrice - potentialOpenTrade.SquareOffValue Then
                                                            If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                        ElseIf potentialOpenTrade.EntryPrice - potentialOpenTrade.PotentialTarget > potentialOpenTrade.EntryPrice - potentialOpenTrade.SquareOffValue Then
                                                            If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade2)
                                                        End If
                                                    Else
                                                        If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                    End If
                                                ElseIf potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                    Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                    dummyPayload.PayloadDate = potentialTickSignalTime
                                                    CancelTrade(potentialOpenTrade, dummyPayload, "Opposite Direction Signal")
                                                End If
                                            Next
                                        Else
                                            If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                            If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                        End If
                                    End If
                                End If

                                If currentSecondTickPayload IsNot Nothing AndAlso currentSecondTickPayload.Count > 0 Then
                                    For Each tick In currentSecondTickPayload
                                        SetCurrentLTPForStock(currentMinuteCandlePayload, tick, Trade.TradeType.MIS)

                                        'MTM Check
                                        If ExitOnFixedTargetStoploss Then
                                            If StockPLAfterBrokerage(tradeCheckingDate, tick.TradingSymbol) >= _maxProfitPerDay OrElse
                                                StockPLAfterBrokerage(tradeCheckingDate, tick.TradingSymbol) <= _maxLossPerDay Then
                                                ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Max PL reached for the day")
                                            End If
                                        End If

                                        'Exit Trade
                                        Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                            For Each potentialExitTrade In potentialExitTrades
                                                ExitTradeIfPossible(potentialExitTrade, tick, False)
                                            Next
                                        End If

                                        'Trailing SL
                                        If TrailingSL Then
                                            Dim potentialSLMoveTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                            If potentialSLMoveTrades IsNot Nothing AndAlso potentialSLMoveTrades.Count > 0 Then
                                                For Each potentialSLMoveTrade In potentialSLMoveTrades
                                                    Dim slMoveRemark As String = Nothing
                                                    Dim slPrice As Double = CalculateMathematicalTrailingSL(potentialSLMoveTrade, tick.Open, 0.1, slMoveRemark)
                                                    If Math.Round(potentialSLMoveTrade.EntryPrice, 2) <> slPrice Then
                                                        If potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                                                            slPrice > potentialSLMoveTrade.PotentialStopLoss Then
                                                            MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, Math.Max(potentialSLMoveTrade.PotentialStopLoss, slPrice), slMoveRemark)
                                                        ElseIf potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                                                            slPrice < potentialSLMoveTrade.PotentialStopLoss Then
                                                            MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, Math.Min(potentialSLMoveTrade.PotentialStopLoss, slPrice), slMoveRemark)
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If

                                        'Enter Trade
                                        Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                            For Each potentialEntryTrade In potentialEntryTrades
                                                Dim placeOrderResponse As Tuple(Of Boolean, Date) = EnterTradeIfPossible(potentialEntryTrade, tick, False)
                                                If placeOrderResponse IsNot Nothing AndAlso placeOrderResponse.Item1 Then
                                                    Dim previousCandleTime As Date = GetPreviousXMinuteCandleTime(potentialCandleSignalTime, currentDayXMinuteStocksPayload(stockName), _signalTimeFrame)
                                                    Dim targetPoint As Decimal = ATRXMinuteStocksPayload(stockName)(previousCandleTime) * TargetMultiplier
                                                    Dim slPoint As Decimal = ATRXMinuteStocksPayload(stockName)(previousCandleTime) * StoplossMultiplier
                                                    Dim potentialTarget As Decimal = Nothing
                                                    Dim potentialSL As Decimal = Nothing
                                                    Dim targetRemark As String = Nothing
                                                    Dim slRemark As String = Nothing
                                                    If potentialEntryTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                        potentialTarget = potentialEntryTrade.EntryPrice + targetPoint
                                                        targetRemark = String.Format("Target: {0}", Math.Round(potentialTarget - potentialEntryTrade.EntryPrice, 2))
                                                        potentialSL = potentialEntryTrade.EntryPrice - slPoint
                                                        slRemark = String.Format("Stoploss: {0}", Math.Round(potentialEntryTrade.EntryPrice - potentialSL, 2))
                                                    ElseIf potentialEntryTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                        potentialTarget = potentialEntryTrade.EntryPrice - targetPoint
                                                        targetRemark = String.Format("Target: {0}", Math.Round(potentialEntryTrade.EntryPrice - potentialTarget, 2))
                                                        potentialSL = potentialEntryTrade.EntryPrice + slPoint
                                                        slRemark = String.Format("Stoploss: {0}", Math.Round(potentialSL - potentialEntryTrade.EntryPrice, 2))
                                                    End If
                                                    If TrailingSL Then
                                                        potentialEntryTrade.UpdateTrade(PotentialStopLoss:=potentialSL, SLRemark:=slRemark, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(previousCandleTime))
                                                    Else
                                                        potentialEntryTrade.UpdateTrade(PotentialTarget:=potentialTarget, TargetRemark:=targetRemark, PotentialStopLoss:=potentialSL, SLRemark:=slRemark, SquareOffValue:=ATRXMinuteStocksPayload(stockName)(previousCandleTime))
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                            startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                        End While   'Second Loop
                        'Force exit at day end
                        If startMinute = EODExitTime Then
                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, "EOD Force Exit")
                        End If
                        startMinute = startMinute.Add(TimeSpan.FromMinutes(_signalTimeFrame))
                    End While   'Minute Loop
                End If
            End If
            totalPL += AllPLAfterBrokerage(tradeCheckingDate)
            tradeCheckingDate = tradeCheckingDate.AddDays(1)
        End While   'Date Loop

        'Excel Printing
        Dim filename As String = String.Format("HA Reversal PL {10},Tgt-{3},Sl-{4},Partial-{5},NmbrTrd-{6},Trailing-{7},RvsTrd-{8},FxdMTM-{9} {0}-{1}-{2}.xlsx",
                                               Now.Hour, Now.Minute, Now.Second, TargetMultiplier, StoplossMultiplier,
                                               If(PartialExit, "T", "F"),
                                               If(NumberOfTradePerDay = Integer.MaxValue, "NoLimit", NumberOfTradePerDay),
                                               If(TrailingSL, "T", "F"),
                                               If(ReverseSignalTrade, "T", "F"),
                                               If(ExitOnFixedTargetStoploss, "T", "F"),
                                               Math.Round(totalPL, 0))

        PrintArrayToExcel(filename)
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
