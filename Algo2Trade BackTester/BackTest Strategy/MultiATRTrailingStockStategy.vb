Imports System.Threading
Imports Algo2TradeBLL
Public Class MultiATRTrailingStockStategy
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)
    Public Sub Run(start_date As Date, end_date As Date, signalTimeFrame As Integer)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim mainSignalTimeFrame As Integer = signalTimeFrame
        Dim firstATRPeriod As Integer = 14
        Dim firstATRTrailingStopMultiplier As Integer = 3
        Dim secondATRPeriod As Integer = 14
        Dim secondATRTrailingStopMultiplier As Integer = 6
        Dim stopLossLimit As Decimal = 15
        Dim targetLimit As Decimal = 10

        'Change Time
        Dim exchangeStartTime As Date = "10:00:00"
        Dim exchangeEndTime As Date = "23:30:00"
        Dim endOfDay As DateTime = "23:00:00"
        Dim lastTradeEntryTime As Date = exchangeEndTime.AddHours(-1)

        Dim dataCheckDate As Date = end_date

        Dim from_date As DateTime = start_date
        Dim to_Date As Date = end_date
        Dim chk_date As Date = from_date
        Dim dateCtr As Integer = 0

        Dim stockList As Dictionary(Of String, Decimal()) = Nothing
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tempStocklist As Dictionary(Of String, Decimal()) = Nothing
        'TODO:
        'Change Stock Name
        tempStocklist = New Dictionary(Of String, Decimal()) From {{"CRUDEOIL", {Nothing}}}
        For Each tradingSymbol In tempStocklist.Keys
            If stockList Is Nothing Then stockList = New Dictionary(Of String, Decimal())
            stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0)})
        Next

        While chk_date <= to_Date
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim OneMinuteHKPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim FirstATRTrailingStopPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SecondATRTrailingStopPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim FractalHighPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim FractalLowPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                For Each item In stockList.Keys
                    Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentOneMinuteHKPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempOneMinuteHKPayload As Dictionary(Of Date, Payload) = Nothing

                    'TODO:
                    'Change fetching data from database table name & start date
                    Dim currentTradingSymbol As String = cmn.GetCurrentTradingSymbol(Common.DataBaseTable.EOD_Commodity, item, chk_date)
                    If PastIntradayData IsNot Nothing AndAlso PastIntradayData.ContainsKey(currentTradingSymbol) Then
                        tempOneMinutePayload = PastIntradayData(currentTradingSymbol)
                    Else
                        tempOneMinutePayload = cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Commodity, currentTradingSymbol, chk_date.AddDays(-7), dataCheckDate)
                        If PastIntradayData Is Nothing Then PastIntradayData = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        PastIntradayData.Add(currentTradingSymbol, tempOneMinutePayload)
                    End If

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
                                    Dim outputFirstATRTrailingStop As Dictionary(Of Date, Decimal) = Nothing
                                    Dim outputSecondATRTrailingStop As Dictionary(Of Date, Decimal) = Nothing
                                    Dim outputColorFirstATRTrailingStop As Dictionary(Of Date, Color) = Nothing
                                    Dim outputColorSecondATRTrailingStop As Dictionary(Of Date, Color) = Nothing
                                    Indicator.ATRTrailingStop.CalculateATRTrailingStop(firstATRPeriod, firstATRTrailingStopMultiplier, tempOneMinuteHKPayload, outputFirstATRTrailingStop, outputColorFirstATRTrailingStop)
                                    Indicator.ATRTrailingStop.CalculateATRTrailingStop(secondATRPeriod, secondATRTrailingStopMultiplier, tempOneMinuteHKPayload, outputSecondATRTrailingStop, outputColorSecondATRTrailingStop)

                                    Dim outputFractalHigh As Dictionary(Of Date, Decimal) = Nothing
                                    Dim outputFractalLow As Dictionary(Of Date, Decimal) = Nothing
                                    If IndicatorData1 IsNot Nothing AndAlso IndicatorData1.ContainsKey(currentTradingSymbol) AndAlso
                                        IndicatorData2 IsNot Nothing AndAlso IndicatorData2.ContainsKey(currentTradingSymbol) Then
                                        outputFractalHigh = IndicatorData1(currentTradingSymbol)
                                        outputFractalLow = IndicatorData2(currentTradingSymbol)
                                    Else
                                        Indicator.Fractals.CalculateFractal(5, tempOneMinuteHKPayload, outputFractalHigh, outputFractalLow)
                                        If IndicatorData1 Is Nothing Then IndicatorData1 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        IndicatorData1.Add(currentTradingSymbol, outputFractalHigh)
                                        If IndicatorData2 Is Nothing Then IndicatorData2 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        IndicatorData2.Add(currentTradingSymbol, outputFractalLow)
                                    End If

                                    If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    OneMinutePayload.Add(item, currentOneMinutePayload)
                                    If OneMinuteHKPayload Is Nothing Then OneMinuteHKPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    OneMinuteHKPayload.Add(item, currentOneMinuteHKPayload)
                                    If FirstATRTrailingStopPayload Is Nothing Then FirstATRTrailingStopPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                    FirstATRTrailingStopPayload.Add(item, outputFirstATRTrailingStop)
                                    If SecondATRTrailingStopPayload Is Nothing Then SecondATRTrailingStopPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                    SecondATRTrailingStopPayload.Add(item, outputSecondATRTrailingStop)
                                    If FractalHighPayload Is Nothing Then FractalHighPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                    FractalHighPayload.Add(item, outputFractalHigh)
                                    If FractalLowPayload Is Nothing Then FractalLowPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                    FractalLowPayload.Add(item, outputFractalLow)
                                End If
                            End If
                        End If
                    End If
                Next

                If OneMinuteHKPayload IsNot Nothing AndAlso OneMinuteHKPayload.Count > 0 Then
                    Dim startTime As Date = exchangeStartTime
                    Dim endTime As Date = exchangeEndTime
                    Dim firstCandleOfDay As Boolean = True
                    Dim exitTrade As Boolean = False
                    While startTime < endTime
                        For Each stockName In stockList.Keys
                            OnHeartbeat(String.Format("Checking Trade for {0} on {1}", stockName, chk_date.ToShortDateString))
                            Dim plOfDay As Double = GetPLForDay(chk_date, stockName)
                            Dim tradeCount As Integer = CountTradesForDay(chk_date, stockName)
                            'TODO:
                            If plOfDay >= targetLimit OrElse plOfDay <= -40 Then
                                Exit While
                            End If
                            Dim potentialTargetLimit As Decimal = targetLimit - plOfDay + tradeCount * 1.5
                            Dim tempStockPayload As Dictionary(Of Date, Payload) = Nothing
                            If OneMinuteHKPayload.ContainsKey(stockName) Then
                                tempStockPayload = OneMinuteHKPayload(stockName)          'Only Current Day Payload
                            End If

                            If tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 Then
                                If startTime = exchangeStartTime Then
                                    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockPayload.Keys.FirstOrDefault.Hour, tempStockPayload.Keys.FirstOrDefault.Minute, tempStockPayload.Keys.FirstOrDefault.Second)
                                    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockPayload.Keys.LastOrDefault.Hour, tempStockPayload.Keys.LastOrDefault.Minute, tempStockPayload.Keys.LastOrDefault.Second)
                                    endTime = endTime.AddMinutes(mainSignalTimeFrame)
                                    endOfDay = endTime.AddMinutes(-30)
                                    lastTradeEntryTime = endTime.AddHours(-1)
                                End If

                                Dim tempStockFirstATRTrailingStopPayload As Dictionary(Of Date, Decimal) = FirstATRTrailingStopPayload(stockName)                 'Full Data
                                Dim tempStockSecondATRTrailingStopPayload As Dictionary(Of Date, Decimal) = SecondATRTrailingStopPayload(stockName)
                                Dim tempStockFractalHighPayload As Dictionary(Of Date, Decimal) = FractalHighPayload(stockName)
                                Dim tempStockFractalLowPayload As Dictionary(Of Date, Decimal) = FractalLowPayload(stockName)

                                Dim potentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)

                                If Not tempStockPayload.ContainsKey(potentialSignalTime) Then
                                    Continue For
                                End If

                                Dim validSignal As Boolean = False
                                Dim signalCandle As Payload = Nothing
                                Dim signalDirection As TradeExecutionDirection = TradeExecutionDirection.None
                                Dim entryPrice As Decimal = Nothing

                                Dim currentStockPayload As Payload = tempStockPayload(potentialSignalTime)
                                If tempStockFirstATRTrailingStopPayload(potentialSignalTime) > tempStockPayload(potentialSignalTime).Close AndAlso
                                            tempStockSecondATRTrailingStopPayload(potentialSignalTime) > tempStockPayload(potentialSignalTime).Close Then
                                    validSignal = True
                                    signalDirection = TradeExecutionDirection.Sell
                                ElseIf tempStockFirstATRTrailingStopPayload(potentialSignalTime) < tempStockPayload(potentialSignalTime).Close AndAlso
                                                tempStockSecondATRTrailingStopPayload(potentialSignalTime) < tempStockPayload(potentialSignalTime).Close Then
                                    validSignal = True
                                    signalDirection = TradeExecutionDirection.Buy
                                Else
                                    firstCandleOfDay = False
                                End If
                                If Not IsTradeActive(chk_date, stockName) AndAlso validSignal Then
                                    If firstCandleOfDay Then
                                        If signalDirection = TradeExecutionDirection.Sell Then
                                            If tempStockFractalLowPayload(potentialSignalTime) > tempStockFractalLowPayload(tempStockPayload(potentialSignalTime).PreviousCandlePayload.PayloadDate) Then
                                                signalCandle = currentStockPayload
                                                entryPrice = tempStockFractalLowPayload(potentialSignalTime) - 1
                                            End If
                                        ElseIf signalDirection = TradeExecutionDirection.Buy Then
                                            'Try
                                            If tempStockFractalHighPayload(potentialSignalTime) < tempStockFractalHighPayload(tempStockPayload(potentialSignalTime).PreviousCandlePayload.PayloadDate) Then
                                                    signalCandle = currentStockPayload
                                                    entryPrice = tempStockFractalHighPayload(potentialSignalTime) + 1
                                                End If
                                            'Catch
                                            '    Console.WriteLine("")
                                            'End Try
                                        End If
                                    Else
                                        If signalDirection = TradeExecutionDirection.Sell Then
                                            If tempStockFractalLowPayload(potentialSignalTime) <= currentStockPayload.Low Then
                                                signalCandle = currentStockPayload
                                                entryPrice = tempStockFractalLowPayload(potentialSignalTime) - 1
                                            End If
                                        ElseIf signalDirection = TradeExecutionDirection.Buy Then
                                            If tempStockFractalHighPayload(potentialSignalTime) >= currentStockPayload.High Then
                                                signalCandle = currentStockPayload
                                                entryPrice = tempStockFractalHighPayload(potentialSignalTime) + 1
                                            End If
                                        End If
                                    End If
                                    exitTrade = False
                                ElseIf IsTradeActive(chk_date, stockName) AndAlso validSignal Then
                                    Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    Dim tradeDirection As TradeExecutionDirection = TradeExecutionDirection.None
                                    For Each item In itemSpecificTrade
                                        tradeDirection = item.EntryDirection
                                    Next
                                    If tradeDirection <> signalDirection Then
                                        If tradeDirection = TradeExecutionDirection.Buy Then
                                            If tempStockFractalLowPayload(potentialSignalTime) <= currentStockPayload.Low Then
                                                signalCandle = currentStockPayload
                                                entryPrice = tempStockFractalLowPayload(potentialSignalTime) - 1
                                            End If
                                        ElseIf tradeDirection = TradeExecutionDirection.Sell Then
                                            If tempStockFractalHighPayload(potentialSignalTime) >= currentStockPayload.High Then
                                                signalCandle = currentStockPayload
                                                entryPrice = tempStockFractalLowPayload(potentialSignalTime) + 1
                                            End If
                                        End If
                                        exitTrade = True
                                    Else
                                        exitTrade = False
                                    End If
                                ElseIf Not validSignal Then
                                    exitTrade = True
                                End If
                                If signalCandle IsNot Nothing AndAlso startTime < lastTradeEntryTime Then
                                    Dim runningTrade As Trade = Nothing
                                    Dim orderSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                    If orderSpecificTrade Is Nothing OrElse orderSpecificTrade.Count = 0 Then
                                        If signalDirection = TradeExecutionDirection.Buy Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryPrice = entryPrice
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .EntryTime = signalCandle.PayloadDate
                                                .EntryType = "MIS"
                                                .SignalCandle = signalCandle
                                                .TradingSymbol = signalCandle.TradingSymbol
                                                .TradingDate = signalCandle.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = .EntryPrice + potentialTargetLimit
                                                .PotentialSL = .EntryPrice - stopLossLimit
                                            End With
                                        ElseIf signalDirection = TradeExecutionDirection.Sell Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryPrice = entryPrice
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .EntryTime = signalCandle.PayloadDate
                                                .EntryType = "MIS"
                                                .SignalCandle = signalCandle
                                                .TradingSymbol = signalCandle.TradingSymbol
                                                .TradingDate = signalCandle.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = .EntryPrice - potentialTargetLimit
                                                .PotentialSL = .EntryPrice + stopLossLimit
                                            End With
                                        End If
                                        If runningTrade IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade)
                                    Else
                                        If signalDirection = TradeExecutionDirection.Buy Then
                                            For Each order In orderSpecificTrade
                                                ModifyOrder(order, entryPrice, entryPrice + potentialTargetLimit, entryPrice - stopLossLimit, order.Quantity, signalCandle)
                                            Next
                                        ElseIf signalDirection = TradeExecutionDirection.Sell Then
                                            For Each order In orderSpecificTrade
                                                ModifyOrder(order, entryPrice, entryPrice - potentialTargetLimit, entryPrice + stopLossLimit, order.Quantity, signalCandle)
                                            Next
                                        End If
                                    End If
                                End If
                                If exitTrade Then
                                    Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    Dim tradeDirection As TradeExecutionDirection = TradeExecutionDirection.None
                                    If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                                        For Each item In itemSpecificTrade
                                            tradeDirection = item.EntryDirection
                                        Next
                                        If tradeDirection = TradeExecutionDirection.Sell Then
                                            If tempStockFractalHighPayload(potentialSignalTime) >= currentStockPayload.High Then
                                                For Each item In itemSpecificTrade
                                                    MoveStopLoss(item, tempStockFractalHighPayload(potentialSignalTime) + 1)
                                                Next
                                            End If
                                        ElseIf tradeDirection = TradeExecutionDirection.Buy Then
                                            If tempStockFractalLowPayload(potentialSignalTime) <= currentStockPayload.Low Then
                                                For Each item In itemSpecificTrade
                                                    MoveStopLoss(item, tempStockFractalLowPayload(potentialSignalTime) - 1)
                                                Next
                                            End If
                                        End If
                                    End If

                                    Dim itemCancelSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                    If itemCancelSpecificTrade IsNot Nothing AndAlso itemCancelSpecificTrade.Count > 0 Then
                                        For Each item In itemCancelSpecificTrade
                                            CancelTrade(item, currentStockPayload)
                                        Next
                                    End If
                                End If
                                    Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                Dim tradeExit As Boolean = False
                                If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                    For Each item In itemExitSpecificTrade
                                        tradeExit = ExitTradeIfPossible(item, currentStockPayload, endOfDay, TypeOfStock.Commodity, False)
                                    Next
                                    If tradeExit Then
                                        exitTrade = False
                                    End If
                                End If
                                Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                Dim tradeEntered As Boolean = False
                                If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                    For Each item In itemEntrySpecificTrade
                                        tradeEntered = EnterTradeIfPossible(item, currentStockPayload, TypeOfStock.Commodity)
                                    Next
                                    If tradeEntered Then
                                        firstCandleOfDay = False
                                    End If
                                End If
                            End If
                        Next
                        startTime = startTime.AddMinutes(mainSignalTimeFrame)
                    End While
                End If
            End If
            chk_date = chk_date.AddDays(1)
        End While
        For Each stock In stockList.Keys
            Dim curDate As Date = System.DateTime.Now
            Dim filename As String = String.Format("MultiATRTrailingStop Strategy {0} for {1} Minute Time Frame {2}-{3}-{4}_{5}-{6}-{7}.xlsx", stock, signalTimeFrame,
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
            PrintArrayToExcel(filename, stock)
        Next
    End Sub
End Class
