Imports System.Threading
Imports Algo2TradeBLL
Public Class ATRinTrendStrategy
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)
    Public Sub Run(start_date As Date, end_date As Date)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        'TODO:
        'Change Time
        Dim exchangeStartTime As Date = "09:15:00"
        Dim exchangeEndTime As Date = "15:30:00"
        Dim endOfDay As DateTime = "15:15:00"
        Dim lastTradeEntryTime As Date = "03:00:00"
        Dim firstTradeEntryTime As Date = "09:20:00"
        Dim mainSignalTimeFrame As Integer = 5

        Dim ATRPeriod As Integer = 14
        Dim stoplossMultiplier As Decimal = 1
        Dim targetMultiplier As Decimal = 1
        Dim signalMultiplier As Decimal = 1 / 2

        Dim dataCheckDate As Date = end_date
        Dim from_date As DateTime = start_date
        Dim to_Date As Date = end_date
        Dim chk_date As Date = from_date
        Dim dateCtr As Integer = 0

        Dim stockList As Dictionary(Of String, Double()) = Nothing
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        While chk_date <= to_Date
            Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))
            tempStocklist = New Dictionary(Of String, Double()) From {{"BAJFINANCE", {250}}}
            For Each tradingSymbol In tempStocklist.Keys
                If stockList Is Nothing Then stockList = New Dictionary(Of String, Double())
                If Not stockList.ContainsKey(tradingSymbol) Then
                    stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0)})
                End If
            Next

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim ATRPaylaod As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                For Each item In stockList.Keys
                    Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim outputATR As Dictionary(Of Date, Decimal) = Nothing

                    'TODO:
                    tempOneMinutePayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Futures, item, chk_date.AddDays(-7), chk_date)

                    For Each tempKeys In tempOneMinutePayload.Keys
                        If tempKeys.Date = chk_date.Date Then
                            If currentOneMinutePayload Is Nothing Then currentOneMinutePayload = New Dictionary(Of Date, Payload)
                            currentOneMinutePayload.Add(tempKeys, tempOneMinutePayload(tempKeys))
                        End If
                    Next
                    tempXMinutePayload = cmn.ConvertPayloadsToXMinutes(tempOneMinutePayload, mainSignalTimeFrame)
                    Indicator.ATR.CalculateATR(ATRPeriod, tempXMinutePayload, outputATR)

                    If currentOneMinutePayload IsNot Nothing AndAlso currentOneMinutePayload.Count > 0 Then
                        If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        OneMinutePayload.Add(item, currentOneMinutePayload)
                        If XMinutePayload Is Nothing Then XMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        XMinutePayload.Add(item, tempXMinutePayload)
                        If ATRPaylaod Is Nothing Then ATRPaylaod = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                        ATRPaylaod.Add(item, outputATR)
                    End If
                Next

                If OneMinutePayload IsNot Nothing AndAlso OneMinutePayload.Count > 0 Then
                    Dim startTime As Date = exchangeStartTime
                    Dim endTime As Date = exchangeEndTime
                    Dim firstCandleOfDay As Boolean = True
                    Dim exitTrade As Boolean = False
                    While startTime < endTime
                        OnHeartbeat(String.Format("Checking Trade for {0}", chk_date.ToShortDateString))
                        Dim tempStockPayload As Dictionary(Of Date, Payload) = Nothing
                        Dim tempStockXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                        For Each stockName In stockList.Keys
                            If OneMinutePayload.ContainsKey(stockName) Then
                                tempStockPayload = OneMinutePayload(stockName)          'Only Current Day Payload
                            End If
                            If XMinutePayload.ContainsKey(stockName) Then
                                tempStockXMinutePayload = XMinutePayload(stockName)
                            End If

                            If tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 AndAlso
                                tempStockXMinutePayload IsNot Nothing AndAlso tempStockXMinutePayload.Count > 0 Then
                                If startTime = exchangeStartTime Then
                                    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockPayload.Keys.FirstOrDefault.Hour, tempStockPayload.Keys.FirstOrDefault.Minute, tempStockPayload.Keys.FirstOrDefault.Second)
                                    startTime = startTime.AddMinutes(mainSignalTimeFrame)
                                    firstTradeEntryTime = startTime
                                    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockPayload.Keys.LastOrDefault.Hour, tempStockPayload.Keys.LastOrDefault.Minute, tempStockPayload.Keys.LastOrDefault.Second)
                                    endTime = endTime.AddMinutes(1)
                                    endOfDay = endTime.AddMinutes(-15)
                                    lastTradeEntryTime = endTime.AddMinutes(-30)
                                End If

                                Dim potentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)
                                If Not tempStockPayload.ContainsKey(potentialSignalTime) Then
                                    Continue For
                                End If

                                Dim currentPayload As Payload = tempStockPayload(potentialSignalTime)
                                Dim runningTrade As Trade = Nothing

                                Dim tickCount As Integer = 0

                                While tickCount < currentPayload.Ticks.Count - 1
                                    If startTime = firstTradeEntryTime AndAlso tickCount = 0 Then
                                        runningTrade = New Trade
                                        With runningTrade
                                            .TradingStatus = TradeExecutionStatus.Open
                                            .EntryPrice = currentPayload.Open
                                            .EntryDirection = TradeExecutionDirection.Sell
                                            .EntryTime = currentPayload.PayloadDate
                                            .SignalCandle = currentPayload
                                            .TradingSymbol = currentPayload.TradingSymbol
                                            .TradingDate = currentPayload.PayloadDate
                                            .IndicatorCandleTime = GetPreviousXMinuteCandleTime(.TradingDate, tempStockXMinutePayload, mainSignalTimeFrame)
                                            .AbsoluteATR = ATRPaylaod(stockName)(.IndicatorCandleTime)
                                            .PotentialSL = .EntryPrice + .AbsoluteATR * stoplossMultiplier
                                            .PotentialTP = .EntryPrice - .AbsoluteATR * targetMultiplier
                                            .Quantity = stockList(stockName)(0)
                                        End With
                                        If runningTrade IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade)
                                    End If

                                    If startTime < lastTradeEntryTime Then
                                        Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                        If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                            For Each item In itemEntrySpecificTrade
                                                Dim previousTickCount As Integer = tickCount
                                                'TODO:
                                                If EnterTradeIfPossible(item, currentPayload, TypeOfStock.Futures, previousTickCount) Then
                                                    tickCount = previousTickCount
                                                    Dim itemCancelSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                                    If itemCancelSpecificTrade IsNot Nothing AndAlso itemCancelSpecificTrade.Count > 0 Then
                                                        For Each cancelItem In itemCancelSpecificTrade
                                                            CancelTrade(cancelItem, currentPayload)
                                                        Next
                                                    End If
                                                    If item.EntryDirection = TradeExecutionDirection.Sell Then
                                                        runningTrade = New Trade
                                                        With runningTrade
                                                            .TradingStatus = TradeExecutionStatus.Open
                                                            .EntryPrice = item.PotentialSL
                                                            .EntryDirection = TradeExecutionDirection.Buy
                                                            .EntryTime = currentPayload.PayloadDate
                                                            .SignalCandle = currentPayload
                                                            .TradingSymbol = currentPayload.TradingSymbol
                                                            .TradingDate = currentPayload.PayloadDate
                                                            .IndicatorCandleTime = GetPreviousXMinuteCandleTime(.TradingDate, tempStockXMinutePayload, mainSignalTimeFrame)
                                                            .AbsoluteATR = ATRPaylaod(stockName)(.IndicatorCandleTime)
                                                            .PotentialSL = .EntryPrice - .AbsoluteATR * stoplossMultiplier
                                                            .PotentialTP = .EntryPrice + .AbsoluteATR * targetMultiplier
                                                            .Quantity = stockList(stockName)(0)
                                                        End With
                                                    ElseIf item.EntryDirection = TradeExecutionDirection.Buy Then
                                                        runningTrade = New Trade
                                                        With runningTrade
                                                            .TradingStatus = TradeExecutionStatus.Open
                                                            .EntryPrice = item.PotentialSL
                                                            .EntryDirection = TradeExecutionDirection.Sell
                                                            .EntryTime = currentPayload.PayloadDate
                                                            .SignalCandle = currentPayload
                                                            .TradingSymbol = currentPayload.TradingSymbol
                                                            .TradingDate = currentPayload.PayloadDate
                                                            .IndicatorCandleTime = GetPreviousXMinuteCandleTime(.TradingDate, tempStockXMinutePayload, mainSignalTimeFrame)
                                                            .AbsoluteATR = ATRPaylaod(stockName)(.IndicatorCandleTime)
                                                            .PotentialSL = .EntryPrice + .AbsoluteATR * stoplossMultiplier
                                                            .PotentialTP = .EntryPrice - .AbsoluteATR * targetMultiplier
                                                            .Quantity = stockList(stockName)(0)
                                                        End With
                                                    End If
                                                    If runningTrade IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade)
                                                Else
                                                    If itemEntrySpecificTrade.Count = 1 AndAlso previousTickCount = currentPayload.Ticks.Count - 1 Then
                                                        tickCount = 0
                                                    End If
                                                End If
                                            Next
                                            Dim itemProgressSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                            If itemProgressSpecificTrade Is Nothing OrElse itemProgressSpecificTrade.Count = 0 Then
                                                tickCount = currentPayload.Ticks.Count - 1
                                            End If
                                        End If
                                    Else
                                        Dim itemCancelSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                        If itemCancelSpecificTrade IsNot Nothing AndAlso itemCancelSpecificTrade.Count > 0 Then
                                            For Each cancelItem In itemCancelSpecificTrade
                                                CancelTrade(cancelItem, currentPayload)
                                            Next
                                        End If
                                        Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                        If itemSpecificTrade Is Nothing OrElse itemSpecificTrade.Count = 0 Then
                                            Exit While
                                        End If
                                    End If


                                    Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                        For Each item In itemExitSpecificTrade
                                            'TODO:
                                            If ExitTradeIfPossible(item, currentPayload, endOfDay, TypeOfStock.Futures, False, tickCount) Then
                                                If item.ExitCondition = TradeExitCondition.Target AndAlso startTime < lastTradeEntryTime Then
                                                    Dim itemCancelSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                                    If itemCancelSpecificTrade IsNot Nothing AndAlso itemCancelSpecificTrade.Count > 0 Then
                                                        For Each cancelItem In itemCancelSpecificTrade
                                                            CancelTrade(cancelItem, currentPayload)
                                                        Next
                                                    End If
                                                    Dim runningTrade1 As Trade = Nothing
                                                    Dim runningTrade2 As Trade = Nothing
                                                    runningTrade1 = New Trade
                                                    With runningTrade1
                                                        .TradingStatus = TradeExecutionStatus.Open
                                                        .EntryPrice = item.PotentialTP + ATRPaylaod(stockName)(GetPreviousXMinuteCandleTime(currentPayload.PayloadDate, tempStockXMinutePayload, mainSignalTimeFrame)) * signalMultiplier
                                                        .EntryDirection = TradeExecutionDirection.Buy
                                                        .EntryTime = currentPayload.PayloadDate
                                                        .SignalCandle = currentPayload
                                                        .TradingSymbol = currentPayload.TradingSymbol
                                                        .TradingDate = currentPayload.PayloadDate
                                                        .IndicatorCandleTime = GetPreviousXMinuteCandleTime(.TradingDate, tempStockXMinutePayload, mainSignalTimeFrame)
                                                        .AbsoluteATR = ATRPaylaod(stockName)(.IndicatorCandleTime)
                                                        .PotentialSL = .EntryPrice - .AbsoluteATR * stoplossMultiplier
                                                        .PotentialTP = .EntryPrice + .AbsoluteATR * targetMultiplier
                                                        .Quantity = stockList(stockName)(0)
                                                    End With
                                                    runningTrade2 = New Trade
                                                    With runningTrade2
                                                        .TradingStatus = TradeExecutionStatus.Open
                                                        .EntryPrice = item.PotentialTP - ATRPaylaod(stockName)(GetPreviousXMinuteCandleTime(currentPayload.PayloadDate, tempStockXMinutePayload, mainSignalTimeFrame)) * signalMultiplier
                                                        .EntryDirection = TradeExecutionDirection.Sell
                                                        .EntryTime = currentPayload.PayloadDate
                                                        .SignalCandle = currentPayload
                                                        .TradingSymbol = currentPayload.TradingSymbol
                                                        .TradingDate = currentPayload.PayloadDate
                                                        .IndicatorCandleTime = GetPreviousXMinuteCandleTime(.TradingDate, tempStockXMinutePayload, mainSignalTimeFrame)
                                                        .AbsoluteATR = ATRPaylaod(stockName)(.IndicatorCandleTime)
                                                        .PotentialSL = .EntryPrice + .AbsoluteATR * stoplossMultiplier
                                                        .PotentialTP = .EntryPrice - .AbsoluteATR * targetMultiplier
                                                        .Quantity = stockList(stockName)(0)
                                                    End With
                                                    If runningTrade1 IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade1)
                                                    If runningTrade2 IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade2)
                                                End If
                                            End If
                                        Next
                                    End If
                                End While

                                Dim itemModifySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                For Each item In itemModifySpecificTrade
                                    Dim targetPrice As Double = Nothing
                                    Dim stoplossPrice As Double = Nothing
                                    Dim previousXMinuteCandleTime As Date = GetPreviousXMinuteCandleTime(currentPayload.PayloadDate, tempStockXMinutePayload, mainSignalTimeFrame)
                                    Dim absoluteATR As Double = ATRPaylaod(stockName)(previousXMinuteCandleTime)
                                    If item.EntryDirection = TradeExecutionDirection.Buy Then
                                        targetPrice = item.EntryPrice + absoluteATR * targetMultiplier
                                        stoplossPrice = item.EntryPrice - absoluteATR * stoplossMultiplier
                                    ElseIf item.EntryDirection = TradeExecutionDirection.Sell Then
                                        targetPrice = item.EntryPrice - absoluteATR * targetMultiplier
                                        stoplossPrice = item.EntryPrice + absoluteATR * stoplossMultiplier
                                    End If
                                    ModifyOrder(item, item.EntryPrice, targetPrice, stoplossPrice, item.Quantity, item.SignalCandle, absoluteATR, previousXMinuteCandleTime)
                                Next
                            End If
                        Next
                        startTime = startTime.AddMinutes(1)
                    End While
                End If
            End If
            chk_date = chk_date.AddDays(1)
        End While
        For Each stock In stockList.Keys
            Dim curDate As DateTime = System.DateTime.Now
            Dim filename As String = String.Format("ATR in trend Strategy for {0} {1}-{2}-{3}_{4}-{5}-{6}.xlsx", stock,
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
            PrintArrayToExcel(filename, stock)
        Next
    End Sub
End Class
