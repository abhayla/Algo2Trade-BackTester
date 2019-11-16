Imports System.Threading
Imports Algo2TradeBLL
Imports MySql.Data.MySqlClient
Imports Utilities.Numbers
Imports System.Drawing

Public Class JOYMAStrategyOnHighATRStock
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)
    Public Sub Run(start_date As Date, end_date As Date)
        Try
            AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

            Dim fastATRPeriod As Integer = 1
            Dim fastATRTrailingStopMultiplier As Integer = 3
            Dim slowATRPeriod As Integer = 5
            Dim slowATRTrailingStopMultiplier As Integer = 3
            Dim ATRPeriod As Integer = 14
            Dim investmentPerStock As Double = 30000
            Dim targetMultiplier As Decimal = 1 / 3 'of Capital
            Dim maximumStopLoss As Double = -5000
            Dim maximumProfit As Double = 10000
            Dim maximumLoss As Double = -15000
            'Details to get stock
            Dim numberOfRecords As Integer = 5
            Dim minClose As Double = 100
            Dim maxClose As Double = 1500
            Dim minATRPercentage As Double = 3
            Dim minPerMiinuteLots As Double = 25

            'TODO:
            'Change Time
            Dim exchangeStartTime As Date = "09:15:00"
            Dim exchangeEndTime As Date = "15:30:00"
            Dim endOfDay As DateTime = "15:15:00"
            Dim lastTradeEntryTime As Date = "03:00:00"
            Dim firstTradeEntryTime As Date = "09:16:00"
            Dim mainSignalTimeFrame As Integer = 1

            Dim dataCheckDate As Date = end_date

            Dim from_date As DateTime = start_date
            Dim to_Date As Date = end_date
            Dim chk_date As Date = from_date
            Dim dateCtr As Integer = 0

            TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

            While chk_date <= to_Date
                Dim stockList As Dictionary(Of String, Double()) = Nothing
                Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
                dateCtr += 1
                OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))
                tempStocklist = GetStockListForTheDay(chk_date, numberOfRecords, minClose, maxClose, minATRPercentage, minPerMiinuteLots)
                'tempStocklist = New Dictionary(Of String, Double()) From {{"JINDALSTEL", {2250, 0, 0, 163.3}}}
                'tempStocklist = New Dictionary(Of String, Double()) From {{"RELCAPITAL", {1500, 0, 0, 0}}}

                If tempStocklist IsNot Nothing AndAlso tempStocklist.Count > 0 Then
                    For Each tradingSymbol In tempStocklist.Keys
                        If stockList Is Nothing Then stockList = New Dictionary(Of String, Double())
                        If Not stockList.ContainsKey(tradingSymbol) Then
                            stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0)})
                        End If
                    Next
                End If

                If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                    Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                    Dim OneMinuteHKPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                    Dim FastATRTrailingStopPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    Dim SlowATRTrailingStopPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    Dim FastATRTrailingStopColorPayload As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                    Dim SlowATRTrailingStopColorPayload As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                    Dim ATRPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    Dim FractalHighPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    Dim FractalLowPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                    For Each item In stockList.Keys
                        Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                        Dim currentOneMinuteHKPayload As Dictionary(Of Date, Payload) = Nothing
                        Dim tempOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                        Dim tempOneMinuteHKPayload As Dictionary(Of Date, Payload) = Nothing

                        'TODO:
                        'Change fetching data from database table name & start date
                        tempOneMinutePayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, item, chk_date.AddDays(-7), chk_date)

                        OnHeartbeat(String.Format("Processing Data for {0}", chk_date.ToShortDateString))

                        If tempOneMinutePayload IsNot Nothing AndAlso tempOneMinutePayload.Count > 0 Then
                            For Each tempKeys In tempOneMinutePayload.Keys
                                If tempKeys.Date = chk_date.Date Then
                                    If currentOneMinutePayload Is Nothing Then currentOneMinutePayload = New Dictionary(Of Date, Payload)
                                    currentOneMinutePayload.Add(tempKeys, tempOneMinutePayload(tempKeys))
                                End If
                            Next
                            If currentOneMinutePayload IsNot Nothing AndAlso currentOneMinutePayload.Count > 0 Then
                                Indicator.HeikenAshi.ConvertToHeikenAshi(tempOneMinutePayload, tempOneMinuteHKPayload)
                                If tempOneMinuteHKPayload IsNot Nothing AndAlso tempOneMinuteHKPayload.Count > 0 Then
                                    For Each tempKeys In tempOneMinuteHKPayload.Keys
                                        If tempKeys.Date = chk_date.Date Then
                                            If currentOneMinuteHKPayload Is Nothing Then currentOneMinuteHKPayload = New Dictionary(Of Date, Payload)
                                            currentOneMinuteHKPayload.Add(tempKeys, tempOneMinuteHKPayload(tempKeys))
                                        End If
                                    Next
                                    If currentOneMinuteHKPayload IsNot Nothing AndAlso currentOneMinuteHKPayload.Count > 0 Then
                                        Dim outputFastATRTrailingStop As Dictionary(Of Date, Decimal) = Nothing
                                        Dim outputFastColorATRTrailingStop As Dictionary(Of Date, Color) = Nothing
                                        Dim outputSlowATRTrailingStop As Dictionary(Of Date, Decimal) = Nothing
                                        Dim outputSlowColorATRTrailingStop As Dictionary(Of Date, Color) = Nothing
                                        Dim outputATR As Dictionary(Of Date, Decimal) = Nothing

                                        OnHeartbeat("Calculating ATR Trailing Stoploss")
                                        Indicator.ATRTrailingStop.CalculateATRTrailingStop(fastATRPeriod, fastATRTrailingStopMultiplier, tempOneMinuteHKPayload, outputFastATRTrailingStop, outputFastColorATRTrailingStop)
                                        Indicator.ATRTrailingStop.CalculateATRTrailingStop(slowATRPeriod, slowATRTrailingStopMultiplier, tempOneMinuteHKPayload, outputSlowATRTrailingStop, outputSlowColorATRTrailingStop)
                                        OnHeartbeat("Calculating ATR")
                                        Indicator.ATR.CalculateATR(ATRPeriod, tempOneMinuteHKPayload, outputATR)
                                        OnHeartbeat("Calculating Fractals")
                                        Dim outputFractalHigh As Dictionary(Of Date, Decimal) = Nothing
                                        Dim outputFractalLow As Dictionary(Of Date, Decimal) = Nothing
                                        Indicator.Fractals.CalculateFractal(5, tempOneMinuteHKPayload, outputFractalHigh, outputFractalLow)

                                        If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        OneMinutePayload.Add(item, currentOneMinutePayload)
                                        If OneMinuteHKPayload Is Nothing Then OneMinuteHKPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        OneMinuteHKPayload.Add(item, currentOneMinuteHKPayload)
                                        If FastATRTrailingStopPayload Is Nothing Then FastATRTrailingStopPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        FastATRTrailingStopPayload.Add(item, outputFastATRTrailingStop)
                                        If SlowATRTrailingStopPayload Is Nothing Then SlowATRTrailingStopPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        SlowATRTrailingStopPayload.Add(item, outputSlowATRTrailingStop)
                                        If FastATRTrailingStopColorPayload Is Nothing Then FastATRTrailingStopColorPayload = New Dictionary(Of String, Dictionary(Of Date, Color))
                                        FastATRTrailingStopColorPayload.Add(item, outputFastColorATRTrailingStop)
                                        If SlowATRTrailingStopColorPayload Is Nothing Then SlowATRTrailingStopColorPayload = New Dictionary(Of String, Dictionary(Of Date, Color))
                                        SlowATRTrailingStopColorPayload.Add(item, outputSlowColorATRTrailingStop)
                                        If ATRPayload Is Nothing Then ATRPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        ATRPayload.Add(item, outputATR)
                                        If FractalHighPayload Is Nothing Then FractalHighPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        FractalHighPayload.Add(item, outputFractalHigh)
                                        If FractalLowPayload Is Nothing Then FractalLowPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        FractalLowPayload.Add(item, outputFractalLow)
                                    End If
                                End If
                            End If
                        End If
                    Next

                    'Logical Portion Of Trading
                    If OneMinuteHKPayload IsNot Nothing AndAlso OneMinuteHKPayload.Count > 0 Then
                        Dim startTime As Date = exchangeStartTime
                        Dim endTime As Date = exchangeEndTime
                        Dim overallProfitLoss As Double = Nothing
                        Dim exitAllTradesForTargetAchieved As Boolean = False
                        Dim forceExitTick As Double = Nothing
                        While startTime < endTime
                            overallProfitLoss = GetOverallProfitLoss(chk_date, chk_date)
                            'OverAll Profit Exit
                            If overallProfitLoss > maximumProfit OrElse overallProfitLoss < maximumLoss Then
                                For Each stock In stockList.Keys
                                    Dim itemProgressTrade As List(Of Trade) = GetSpecificTrades(chk_date, stock, TradeExecutionStatus.Inprogress)
                                    If itemProgressTrade Is Nothing OrElse itemProgressTrade.Count = 0 Then
                                        Dim itemForceCancelTrade As List(Of Trade) = GetSpecificTrades(chk_date, stock, TradeExecutionStatus.Open)
                                        If itemForceCancelTrade IsNot Nothing AndAlso itemForceCancelTrade.Count > 0 Then
                                            CancelTrade(itemForceCancelTrade.LastOrDefault, OneMinutePayload(stock)(New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)))
                                        End If
                                    End If
                                Next
                                GoTo l
                            End If
                            Dim minStock As String = Nothing
                            Dim minTick As Integer = Integer.MaxValue
                            For Each stockName In stockList.Keys
                                OnHeartbeat(String.Format("Checking Trade for {0} on {1}", stockName, chk_date.ToShortDateString))
                                Dim totalProfitLoss As Double = GetProfitLossForStock(stockName, chk_date, chk_date)
                                If totalProfitLoss >= maximumProfit OrElse totalProfitLoss <= maximumStopLoss Then
                                    Continue For
                                End If
                                Dim potentialTargetPrice As Double = Math.Round((maximumProfit - totalProfitLoss), 2)
                                Dim potentialStopLossPrice As Double = Math.Round((maximumStopLoss - If(totalProfitLoss > 0, 0, totalProfitLoss)), 2)

                                Dim tempStockHKPayload As Dictionary(Of Date, Payload) = Nothing
                                Dim tempStockPayload As Dictionary(Of Date, Payload) = Nothing
                                If OneMinuteHKPayload.ContainsKey(stockName) Then
                                    tempStockHKPayload = OneMinuteHKPayload(stockName)          'Only Current Day Payload
                                End If
                                If OneMinutePayload.ContainsKey(stockName) Then
                                    tempStockPayload = OneMinutePayload(stockName)          'Only Current Day Payload
                                End If

                                If tempStockHKPayload IsNot Nothing AndAlso tempStockHKPayload.Count > 0 AndAlso
                                    tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 Then
                                    If startTime = exchangeStartTime Then
                                        startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockHKPayload.Keys.FirstOrDefault.Hour, tempStockHKPayload.Keys.FirstOrDefault.Minute, tempStockHKPayload.Keys.FirstOrDefault.Second)
                                        endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockHKPayload.Keys.LastOrDefault.Hour, tempStockHKPayload.Keys.LastOrDefault.Minute, tempStockHKPayload.Keys.LastOrDefault.Second)
                                        startTime = startTime.AddMinutes(mainSignalTimeFrame)
                                        endTime = endTime.AddMinutes(mainSignalTimeFrame)
                                        endOfDay = endTime.AddMinutes(-15)
                                        lastTradeEntryTime = endTime.AddMinutes(-30)
                                    End If

                                    Dim tempStockFastATRTrailingStopPayload As Dictionary(Of Date, Decimal) = FastATRTrailingStopPayload(stockName)                 'Full Data
                                    Dim tempStockSlowATRTrailingStopPayload As Dictionary(Of Date, Decimal) = SlowATRTrailingStopPayload(stockName)
                                    Dim tempStockFastATRTrailingStopColorPayload As Dictionary(Of Date, Color) = FastATRTrailingStopColorPayload(stockName)                 'Full Data
                                    Dim tempStockSlowATRTrailingStopColorPayload As Dictionary(Of Date, Color) = SlowATRTrailingStopColorPayload(stockName)
                                    Dim tempStockATRPayload As Dictionary(Of Date, Decimal) = ATRPayload(stockName)
                                    Dim tempStockFractalHighPayload As Dictionary(Of Date, Decimal) = FractalHighPayload(stockName)
                                    Dim tempStockFractalLowPayload As Dictionary(Of Date, Decimal) = FractalLowPayload(stockName)

                                    Dim potentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)

                                    If Not tempStockHKPayload.ContainsKey(potentialSignalTime) Then
                                        Continue For
                                    End If

                                    Dim currentActiveStockList As Dictionary(Of String, Double()) = Nothing
                                    For Each stock In stockList.Keys
                                        Dim activeTrade As List(Of Trade) = GetSpecificTrades(chk_date, stock, TradeExecutionStatus.Inprogress)
                                        minStock = Nothing
                                        If activeTrade IsNot Nothing AndAlso activeTrade.Count > 0 Then
                                            Dim stockPayload As Payload = OneMinutePayload(stock)(potentialSignalTime)
                                            If currentActiveStockList Is Nothing Then currentActiveStockList = New Dictionary(Of String, Double())
                                            currentActiveStockList.Add(stock, stockList(stock))
                                            If stockPayload.Ticks.Count < minTick Then
                                                minStock = stock
                                            End If
                                        End If
                                    Next
                                    If minStock IsNot Nothing Then
                                        Dim minStockPayload As Payload = OneMinutePayload(minStock)(potentialSignalTime)
                                        For i = 1 To minStockPayload.Ticks.Count
                                            Dim timeTraversedPercentage As Double = (i / minStockPayload.Ticks.Count) * 100
                                            Dim potentialForceExitStockList As Dictionary(Of String, Integer) = Nothing
                                            For Each potentialStock In currentActiveStockList.Keys
                                                Dim tickPosition As Integer = OneMinutePayload(potentialStock)(potentialSignalTime).Ticks.Count * timeTraversedPercentage / 100
                                                If tickPosition <> 0 Then
                                                    If potentialForceExitStockList Is Nothing Then potentialForceExitStockList = New Dictionary(Of String, Integer)
                                                    potentialForceExitStockList.Add(potentialStock, tickPosition - 1)
                                                End If
                                            Next
                                            If potentialForceExitStockList IsNot Nothing AndAlso potentialForceExitStockList.Count > 0 Then
                                                Dim currentMTM As Double = 0
                                                For Each data In potentialForceExitStockList.Keys
                                                    Dim currentTrade As List(Of Trade) = GetSpecificTrades(chk_date, data, TradeExecutionStatus.Inprogress)
                                                    Dim potentialPayload As Payload = OneMinutePayload(data)(potentialSignalTime)
                                                    If currentTrade IsNot Nothing AndAlso currentTrade.Count > 0 Then
                                                        If currentTrade.LastOrDefault.EntryDirection = TradeExecutionDirection.Buy Then
                                                            currentMTM += CalculateProfitLoss(currentTrade.LastOrDefault.TradingSymbol, currentTrade.LastOrDefault.EntryPrice, potentialPayload.Ticks(potentialForceExitStockList(data)), currentTrade.LastOrDefault.Quantity, TypeOfStock.Cash)
                                                        ElseIf currentTrade.LastOrDefault.EntryDirection = TradeExecutionDirection.Sell Then
                                                            currentMTM += CalculateProfitLoss(currentTrade.LastOrDefault.TradingSymbol, potentialPayload.Ticks(potentialForceExitStockList(data)), currentTrade.LastOrDefault.EntryPrice, currentTrade.LastOrDefault.Quantity, TypeOfStock.Cash)
                                                        End If
                                                    End If
                                                Next
                                                For Each stock In stockList.Keys
                                                    Dim exitedTrade As List(Of Trade) = GetSpecificTrades(chk_date, stock, TradeExecutionStatus.Close)
                                                    If exitedTrade IsNot Nothing AndAlso exitedTrade.Count > 0 Then
                                                        For Each item In exitedTrade
                                                            currentMTM += item.ProfitLoss
                                                        Next
                                                    End If
                                                Next
                                                If currentMTM > maximumProfit OrElse currentMTM < maximumLoss Then
                                                    Console.WriteLine(chk_date.ToShortDateString)
                                                    exitAllTradesForTargetAchieved = True
                                                    Dim stocks = potentialForceExitStockList.Keys.ToList
                                                    For stockIndex = 0 To stocks.Count - 1
                                                        Dim forceExitStock As String = stocks(stockIndex)
                                                        Dim itemForceExitTrade As List(Of Trade) = GetSpecificTrades(chk_date, forceExitStock, TradeExecutionStatus.Inprogress)
                                                        If itemForceExitTrade IsNot Nothing AndAlso itemForceExitTrade.Count > 0 Then
                                                            ExitTradeIfPossible(itemForceExitTrade.LastOrDefault, OneMinutePayload(forceExitStock)(potentialSignalTime), endOfDay, TypeOfStock.Cash, True, potentialForceExitStockList(forceExitStock))
                                                        End If
                                                    Next
                                                    For Each stock In stockList.Keys
                                                        Dim itemForceCancelTrade As List(Of Trade) = GetSpecificTrades(chk_date, stock, TradeExecutionStatus.Open)
                                                        If itemForceCancelTrade IsNot Nothing AndAlso itemForceCancelTrade.Count > 0 Then
                                                            CancelTrade(itemForceCancelTrade.LastOrDefault, OneMinutePayload(stock)(potentialSignalTime))
                                                        End If
                                                    Next
                                                    GoTo l
                                                End If
                                            End If
                                        Next
                                    End If
                                    'Trading Part
                                    Dim currentHKPayload As Payload = tempStockHKPayload(potentialSignalTime)
                                    Dim currentPayload As Payload = tempStockPayload(potentialSignalTime)
                                    Dim runningTrade As Trade = Nothing
                                    Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                    Dim itemInProgressTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    If tempStockSlowATRTrailingStopColorPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) = Color.Green AndAlso
                                    currentPayload.PreviousCandlePayload.Close > tempStockFractalLowPayload(currentPayload.PreviousCandlePayload.PayloadDate) Then
                                        Dim entryPrice As Double = tempStockFractalLowPayload(currentHKPayload.PreviousCandlePayload.PayloadDate)
                                        If tempStockFastATRTrailingStopColorPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) = Color.Green Then
                                            If currentHKPayload.PreviousCandlePayload.Low < tempStockFastATRTrailingStopPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) Then
                                                entryPrice = Math.Max(currentHKPayload.PreviousCandlePayload.Low, entryPrice)
                                            End If
                                        End If
                                        Dim modifySignal As Boolean = False
                                        Dim validSIgnal As Boolean = True
                                        If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                                            For Each item In itemSpecificTrade
                                                If item.EntryDirection = TradeExecutionDirection.Sell Then
                                                    modifySignal = True
                                                End If
                                            Next
                                        End If
                                        If itemInProgressTrade IsNot Nothing AndAlso itemInProgressTrade.Count > 0 Then
                                            For Each itemInprogress In itemInProgressTrade
                                                If itemInprogress.EntryDirection = TradeExecutionDirection.Sell Then
                                                    validSIgnal = False
                                                End If
                                            Next
                                        End If
                                        If Not modifySignal AndAlso validSIgnal AndAlso startTime < lastTradeEntryTime Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .EntryPrice = Math.Round((entryPrice - CalculateBuffer(entryPrice, RoundOfType.Celing)), 2)
                                                .EntryTime = currentHKPayload.PayloadDate
                                                .EntryType = "MIS"
                                                .SignalCandle = currentHKPayload
                                                .TradingSymbol = currentHKPayload.TradingSymbol
                                                .TradingDate = currentHKPayload.PayloadDate.Date
                                                .Quantity = CalculateTradeQuantity(stockList(stockName)(0), investmentPerStock, .EntryPrice)
                                                .PotentialTP = CalculateTargetOrStoploss(.TradingSymbol, .EntryPrice, .Quantity, Math.Min(.CapitalRequiredWithMargin * targetMultiplier, potentialTargetPrice), .EntryDirection, TypeOfStock.Cash)
                                                .PotentialSL = tempStockFractalHighPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) + CalculateBuffer(tempStockFractalHighPayload(currentHKPayload.PreviousCandlePayload.PayloadDate), RoundOfType.Celing)
                                                If CalculateProfitLoss(.TradingSymbol, .PotentialSL, .EntryPrice, .Quantity, TypeOfStock.Cash) < potentialStopLossPrice Then
                                                    .PotentialSL = CalculateTargetOrStoploss(.TradingSymbol, .EntryPrice, .Quantity, potentialStopLossPrice, .EntryDirection, TypeOfStock.Cash)
                                                End If
                                                .AbsoluteATR = tempStockATRPayload(currentHKPayload.PreviousCandlePayload.PayloadDate)
                                                .IndicatorCandleTime = currentHKPayload.PreviousCandlePayload.PayloadDate
                                                .TypeOfStock = TypeOfStock.Cash
                                            End With
                                        ElseIf modifySignal Then
                                            For Each item In itemSpecificTrade
                                                If item.EntryDirection = TradeExecutionDirection.Sell Then
                                                    Dim modifiedEntryPrice As Double = Math.Max(Math.Round((entryPrice - CalculateBuffer(entryPrice, RoundOfType.Celing)), 2), item.EntryPrice)
                                                    Dim stopLoss As Double = Math.Round((tempStockFractalHighPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) + CalculateBuffer(tempStockFractalHighPayload(currentHKPayload.PreviousCandlePayload.PayloadDate), RoundOfType.Celing)), 2)
                                                    Dim quantity As Integer = CalculateTradeQuantity(stockList(stockName)(0), investmentPerStock, modifiedEntryPrice)
                                                    Dim capital As Double = modifiedEntryPrice * quantity / 30
                                                    If CalculateProfitLoss(item.TradingSymbol, stopLoss, modifiedEntryPrice, quantity, TypeOfStock.Cash) < potentialStopLossPrice Then
                                                        stopLoss = CalculateTargetOrStoploss(item.TradingSymbol, modifiedEntryPrice, quantity, potentialStopLossPrice, item.EntryDirection, TypeOfStock.Cash)
                                                    End If
                                                    Dim target As Double = CalculateTargetOrStoploss(item.TradingSymbol, modifiedEntryPrice, quantity, Math.Min(capital * targetMultiplier, potentialTargetPrice), item.EntryDirection, TypeOfStock.Cash)
                                                    ModifyOrder(item, modifiedEntryPrice, target, stopLoss, quantity, currentHKPayload, tempStockATRPayload(currentHKPayload.PreviousCandlePayload.PayloadDate), currentHKPayload.PreviousCandlePayload.PayloadDate)
                                                End If
                                            Next
                                        End If
                                    ElseIf tempStockSlowATRTrailingStopColorPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) = Color.Red AndAlso
                                currentPayload.PreviousCandlePayload.Close < tempStockFractalHighPayload(currentPayload.PreviousCandlePayload.PayloadDate) Then
                                        Dim entryPrice As Double = tempStockFractalHighPayload(currentHKPayload.PreviousCandlePayload.PayloadDate)
                                        If tempStockFastATRTrailingStopColorPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) = Color.Red Then
                                            If currentHKPayload.PreviousCandlePayload.High > tempStockFastATRTrailingStopPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) Then
                                                entryPrice = Math.Min(currentHKPayload.PreviousCandlePayload.High, entryPrice)
                                            End If
                                        End If
                                        Dim modifySignal As Boolean = False
                                        Dim validSIgnal As Boolean = True
                                        If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                                            For Each item In itemSpecificTrade
                                                If item.EntryDirection = TradeExecutionDirection.Buy Then
                                                    modifySignal = True
                                                End If
                                            Next
                                        End If
                                        If itemInProgressTrade IsNot Nothing AndAlso itemInProgressTrade.Count > 0 Then
                                            For Each itemInprogress In itemInProgressTrade
                                                If itemInprogress.EntryDirection = TradeExecutionDirection.Buy Then
                                                    validSIgnal = False
                                                End If
                                            Next
                                        End If
                                        If Not modifySignal AndAlso validSIgnal AndAlso startTime < lastTradeEntryTime Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .EntryPrice = Math.Round((entryPrice + CalculateBuffer(entryPrice, RoundOfType.Celing)), 2)
                                                .EntryTime = currentHKPayload.PayloadDate
                                                .EntryType = "MIS"
                                                .SignalCandle = currentHKPayload
                                                .TradingSymbol = currentHKPayload.TradingSymbol
                                                .TradingDate = currentHKPayload.PayloadDate.Date
                                                .Quantity = CalculateTradeQuantity(stockList(stockName)(0), investmentPerStock, .EntryPrice)
                                                .PotentialTP = CalculateTargetOrStoploss(.TradingSymbol, .EntryPrice, .Quantity, Math.Min(.CapitalRequiredWithMargin * targetMultiplier, potentialTargetPrice), .EntryDirection, TypeOfStock.Cash)
                                                .PotentialSL = tempStockFractalLowPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) - CalculateBuffer(tempStockFractalLowPayload(currentHKPayload.PreviousCandlePayload.PayloadDate), RoundOfType.Celing)
                                                If CalculateProfitLoss(.TradingSymbol, .EntryPrice, .PotentialSL, .Quantity, TypeOfStock.Cash) < potentialStopLossPrice Then
                                                    .PotentialSL = CalculateTargetOrStoploss(.TradingSymbol, .EntryPrice, .Quantity, potentialStopLossPrice, .EntryDirection, TypeOfStock.Cash)
                                                End If
                                                .AbsoluteATR = tempStockATRPayload(currentHKPayload.PreviousCandlePayload.PayloadDate)
                                                .IndicatorCandleTime = currentHKPayload.PreviousCandlePayload.PayloadDate
                                                .TypeOfStock = TypeOfStock.Cash
                                            End With
                                        ElseIf modifySignal Then
                                            For Each item In itemSpecificTrade
                                                If item.EntryDirection = TradeExecutionDirection.Buy Then
                                                    Dim modifiedEntryPrice As Double = Math.Min(Math.Round((entryPrice + CalculateBuffer(entryPrice, RoundOfType.Celing)), 2), item.EntryPrice)
                                                    Dim stopLoss As Double = Math.Round((tempStockFractalLowPayload(currentHKPayload.PreviousCandlePayload.PayloadDate) - CalculateBuffer(tempStockFractalLowPayload(currentHKPayload.PreviousCandlePayload.PayloadDate), RoundOfType.Celing)), 2)
                                                    Dim quantity As Integer = CalculateTradeQuantity(stockList(stockName)(0), investmentPerStock, modifiedEntryPrice)
                                                    Dim capital As Double = modifiedEntryPrice * quantity / 30
                                                    If CalculateProfitLoss(item.TradingSymbol, modifiedEntryPrice, stopLoss, quantity, TypeOfStock.Cash) < potentialStopLossPrice Then
                                                        stopLoss = CalculateTargetOrStoploss(item.TradingSymbol, modifiedEntryPrice, quantity, potentialStopLossPrice, item.EntryDirection, TypeOfStock.Cash)
                                                    End If
                                                    Dim target As Double = CalculateTargetOrStoploss(item.TradingSymbol, modifiedEntryPrice, quantity, Math.Min(capital * targetMultiplier, potentialTargetPrice), item.EntryDirection, TypeOfStock.Cash)
                                                    ModifyOrder(item, modifiedEntryPrice, target, stopLoss, quantity, currentHKPayload, tempStockATRPayload(currentHKPayload.PreviousCandlePayload.PayloadDate), currentHKPayload.PreviousCandlePayload.PayloadDate)
                                                End If
                                            Next
                                        End If
                                    End If
                                    If runningTrade IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade)


                                    Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                        For Each itemExit In itemExitSpecificTrade
                                            If ExitTradeIfPossible(itemExit, currentPayload, endOfDay, TypeOfStock.Cash, False) Then
                                                Dim tradeExited As Boolean = True
                                            End If
                                        Next
                                    End If

                                    Dim profitLoss As Double = GetProfitLossForStock(stockName, chk_date, chk_date)
                                    If profitLoss >= maximumProfit OrElse profitLoss <= maximumStopLoss Then
                                        Dim itemCancelTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                        If itemCancelTrade IsNot Nothing AndAlso itemCancelTrade.Count > 0 Then
                                            For Each trades In itemCancelTrade
                                                CancelTrade(trades, currentPayload)
                                            Next
                                        End If
                                        Continue For
                                    End If

                                    Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                    If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                        For Each itemEntry In itemEntrySpecificTrade
                                            If EnterTradeIfPossible(itemEntry, currentPayload, TypeOfStock.Cash) Then
                                                Dim tradeEntered As Boolean = True
                                            End If
                                        Next
                                    End If

                                    Dim changeStopLossSpecificTrades As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    If changeStopLossSpecificTrades IsNot Nothing AndAlso changeStopLossSpecificTrades.Count > 0 Then
                                        For Each changeTrade In changeStopLossSpecificTrades
                                            If changeTrade.EntryDirection = TradeExecutionDirection.Buy Then
                                                Dim potentialStoploss As Double = tempStockFractalLowPayload(currentHKPayload.PayloadDate)
                                                If tempStockFastATRTrailingStopColorPayload(currentHKPayload.PayloadDate) = Color.Green Then
                                                    If currentHKPayload.Low < tempStockFastATRTrailingStopPayload(currentHKPayload.PayloadDate) Then
                                                        potentialStoploss = Math.Max(currentHKPayload.Low, potentialStoploss)
                                                    End If
                                                End If
                                                If CalculateProfitLoss(changeTrade.TradingSymbol, changeTrade.EntryPrice, potentialStoploss, changeTrade.Quantity, TypeOfStock.Cash) < potentialStopLossPrice Then
                                                    potentialStoploss = CalculateTargetOrStoploss(changeTrade.TradingSymbol, changeTrade.EntryPrice, changeTrade.Quantity, potentialStopLossPrice, changeTrade.EntryDirection, TypeOfStock.Cash)
                                                End If
                                                MoveStopLoss(changeTrade, potentialStoploss - CalculateBuffer(potentialStoploss, RoundOfType.Celing))
                                            ElseIf changeTrade.EntryDirection = TradeExecutionDirection.Sell Then
                                                Dim potentialStoploss As Double = tempStockFractalHighPayload(currentHKPayload.PayloadDate)
                                                If tempStockFastATRTrailingStopColorPayload(currentHKPayload.PayloadDate) = Color.Red Then
                                                    If currentHKPayload.High > tempStockFastATRTrailingStopPayload(currentHKPayload.PayloadDate) Then
                                                        potentialStoploss = Math.Min(currentHKPayload.High, potentialStoploss)
                                                    End If
                                                End If
                                                If CalculateProfitLoss(changeTrade.TradingSymbol, potentialStoploss, changeTrade.EntryPrice, changeTrade.Quantity, TypeOfStock.Cash) < potentialStopLossPrice Then
                                                    potentialStoploss = CalculateTargetOrStoploss(changeTrade.TradingSymbol, changeTrade.EntryPrice, changeTrade.Quantity, potentialStopLossPrice, changeTrade.EntryDirection, TypeOfStock.Cash)
                                                End If
                                                MoveStopLoss(changeTrade, potentialStoploss + CalculateBuffer(potentialStoploss, RoundOfType.Celing))
                                            End If
                                        Next
                                    End If
                                    If startTime >= lastTradeEntryTime Then
                                        Dim itemCancelTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                        If itemCancelTrade IsNot Nothing AndAlso itemCancelTrade.Count > 0 Then
                                            For Each trades In itemCancelTrade
                                                CancelTrade(trades, currentPayload)
                                            Next
                                        End If
                                    End If
                                End If
                            Next
                            startTime = startTime.AddMinutes(mainSignalTimeFrame)
                        End While
                    End If
                End If
l:              chk_date = chk_date.AddDays(1)
            End While
            'For Each stock In stockList.Keys
            Dim curDate As DateTime = System.DateTime.Now
            Dim filename As String = String.Format("JOYMA Strategy for {0} {1}-{2}-{3}_{4}-{5}-{6}.xlsx", "All Stock",
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
            PrintArrayToExcel(filename)
            'Next
        Catch ex As Exception
            'TODO:
            'Serialize Trades Taken Object
            Throw ex
        End Try
    End Sub
    Private Function GetStockListForTheDay(currentDate As Date, numberOfRecords As Integer, minClose As Double, maxClose As Double, minATRPercentage As Double, minPerMiinuteLots As Double) As Dictionary(Of String, Double())
        Try
            AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

            Dim previousDate As Date = cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, currentDate)

            Dim ret As Dictionary(Of String, Double()) = Nothing
            Dim outputdt As DataTable = Nothing
            Dim conn As MySqlConnection
            Dim dts As DataSet = Nothing
            conn = cmn.OpenDBConnection

            If conn.State = ConnectionState.Open Then
                OnHeartbeat("Fetching All Stock Data")
                Dim cmd As New MySqlCommand("GET_STOCK_CASH_DATA_ATR_VOLUME_ALL_DATES", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@startDate", previousDate)
                cmd.Parameters.AddWithValue("@endDate", previousDate)
                cmd.Parameters.AddWithValue("@numberOfRecords", numberOfRecords)
                cmd.Parameters.AddWithValue("@minClose", minClose)
                cmd.Parameters.AddWithValue("@maxClose", maxClose)
                cmd.Parameters.AddWithValue("@atrPercentage", minATRPercentage)
                cmd.Parameters.AddWithValue("@perMinuteLots", minPerMiinuteLots)

                Dim adapter As New MySqlDataAdapter(cmd)
                adapter.SelectCommand.CommandTimeout = 3000
                dts = New DataSet
                adapter.Fill(dts)

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
                    outputdt = dt
                End If
            End If

            If outputdt IsNot Nothing AndAlso outputdt.Rows.Count > 0 Then
                Dim i As Integer = 0
                While Not i = outputdt.Rows.Count()
                    If ret Is Nothing Then ret = New Dictionary(Of String, Double())
                    ret.Add(outputdt.Rows(i).Item(1), {outputdt.Rows(i).Item(5), outputdt.Rows(i).Item(2), outputdt.Rows(i).Item(4), outputdt.Rows(i).Item(3)})    'name,lotsize,absATR,ATRper,ATRClose
                    i += 1
                End While
            End If
            Return ret
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function CalculateTradeQuantity(lotSize As Integer, totalInvestment As Double, stockPrice As Double) As Integer
        Dim quantity As Integer = lotSize
        Dim quantityMultiplier As Double = totalInvestment / (quantity * stockPrice / 30)
        quantity = quantity * quantityMultiplier
        Return quantity
    End Function
End Class
