Imports System.Threading
Imports Algo2TradeBLL

Public Class MomentumReversalStrategy
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)
    Public Sub Run(start_date As DateTime, end_date As DateTime)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim chooseStockNumber As Integer = 0

        Dim signalTimeFrame As Integer = 5
        Dim currentCandleWicksSizePercentage As Double = CandleWickSizePercentage
        Dim targetPercentage As Double = WinRatio
        Dim cuurentMaxStopLossPercentage As Double = MaxStopLossPercentage
        Dim tradingStartTime As DateTime = "09:20:00"
        Dim tradingEndTime As DateTime = "15:30:00"
        Dim endOfDay As DateTime = "15:05:00"
        Dim lastTradeSignalCheckingTime As DateTime = "14:30:00"
        Dim maxTrades As Integer = 1
        Dim candleRangePercentage As Double = 0.1
        Dim PLforDay As Double = 0

        Dim from_date As DateTime = start_date
        Dim to_Date As Date = end_date
        Dim chk_date As Date = from_date
        Dim dateCtr As Integer = 0
        Dim outputATR As Dictionary(Of Date, Decimal) = Nothing
        Dim outputVolumeMA As Dictionary(Of Date, Decimal) = Nothing
        Dim outputHeikenAshiPayload As Dictionary(Of Date, Payload) = Nothing
        Dim stockList As Dictionary(Of String, Integer) = Nothing
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        While chk_date <= to_Date
            Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
            stockList = New Dictionary(Of String, Integer)
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))
            'tempStocklist = New Dictionary(Of String, Double()) From {{"SBIN", {300, 3000}}, {"CONCOR", {1480.4, 625}}, {"IDEA", {107.8, 7000}}, {"JINDALSTEL", {264.45, 4500}}, {"RELINFRA", {567.05, 1300}}, {"BHARTIARTL", {508.9, 1700}}, {"JETAIRWAYS", {845.25, 1200}}, {"INDIACEM", {193.9, 3500}}, {"PNB", {171.2, 4000}}, {"HINDALCO", {269.85, 3500}}, {"KPIT", {207.25, 4500}}}
            tempStocklist = New Dictionary(Of String, Double()) From {{"JINDALSTEL", {264.45, 4500}}}
            Dim counter As Integer = 0
            For Each tradingSymbol In tempStocklist.Keys
                'Dim targetPoint As Double = 1 + targetPercentage / 100
                'Dim quantity As Integer = CalculateQuantity(tempStocklist(tradingSymbol)(0), Math.Round((tempStocklist(tradingSymbol)(0)), 2) * targetPoint, NetProfitLossOfDay)
                'If tempStocklist(tradingSymbol)(1) >= quantity * perMinuteLots Then
                '    If stockList Is Nothing Then stockList = New Dictionary(Of String, Integer)
                '    stockList.Add(tradingSymbol, quantity)
                'End If
                'If counter = chooseStockNumber Then
                If stockList Is Nothing Then stockList = New Dictionary(Of String, Integer)
                stockList.Add(tradingSymbol, tempStocklist(tradingSymbol)(1))
                'End If
                counter += 1
            Next

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim HAPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim ATRPaylaod As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim VolumePayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                For Each item In stockList.Keys
                    Dim current_OneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_OneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_XMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim modulatedKey As String = String.Format("{0}_{1}", item, chk_date.ToShortDateString)
                    If Data.PastData IsNot Nothing AndAlso Data.PastData.ContainsKey(modulatedKey) Then
                        current_OneMinutePayload = Data.PastData(modulatedKey)
                    Else
                        current_OneMinutePayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, item, chk_date.AddDays(-7), chk_date)
                        If Data.PastData Is Nothing Then Data.PastData = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        Data.PastData.Add(modulatedKey, current_OneMinutePayload)
                    End If
                    For Each tempKeys In current_OneMinutePayload.Keys
                        If tempKeys.Date = chk_date.Date Then
                            If temp_OneMinutePayload Is Nothing Then temp_OneMinutePayload = New Dictionary(Of Date, Payload)
                            temp_OneMinutePayload.Add(tempKeys, current_OneMinutePayload(tempKeys))
                        End If
                    Next
                    outputATR = CalculateATR(chk_date, current_OneMinutePayload, signalTimeFrame)
                    outputVolumeMA = CalculateMA(chk_date, current_OneMinutePayload, signalTimeFrame)
                    outputHeikenAshiPayload = CalculateHeikenAshi(chk_date, current_OneMinutePayload, signalTimeFrame)
                    If temp_OneMinutePayload IsNot Nothing AndAlso temp_OneMinutePayload.Count > 0 Then
                        temp_XMinutePayload = New Dictionary(Of Date, Payload)
                        temp_XMinutePayload = cmn.ConvertPayloadsToXMinutes(temp_OneMinutePayload, signalTimeFrame)
                        If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        OneMinutePayload.Add(item, temp_OneMinutePayload)
                        If XMinutePayload Is Nothing Then XMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        XMinutePayload.Add(item, temp_XMinutePayload)
                        If HAPayload Is Nothing Then HAPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        HAPayload.Add(item, outputHeikenAshiPayload)
                        If ATRPaylaod Is Nothing Then ATRPaylaod = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                        ATRPaylaod.Add(item, outputATR)
                        If VolumePayload Is Nothing Then VolumePayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                        VolumePayload.Add(item, outputVolumeMA)
                    End If
                Next

                Dim tr_time As DateTime = tradingStartTime
                While tr_time < tradingEndTime
                    Dim tr_date As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, tr_time.Hour, tr_time.Minute, tr_time.Second)
                    If tr_date.TimeOfDay.Hours < lastTradeSignalCheckingTime.TimeOfDay.Hours OrElse (tr_date.TimeOfDay.Hours = lastTradeSignalCheckingTime.TimeOfDay.Hours And tr_date.TimeOfDay.Minutes < lastTradeSignalCheckingTime.TimeOfDay.Minutes) Then
                        For Each tradingSymbol In stockList.Keys
                            OnHeartbeat(String.Format("Checking Trade for {0} on {1}", tradingSymbol, chk_date.ToShortDateString))
                            If OneMinutePayload IsNot Nothing AndAlso XMinutePayload IsNot Nothing AndAlso
                                XMinutePayload.ContainsKey(tradingSymbol) AndAlso XMinutePayload(tradingSymbol).ContainsKey(tr_date) Then
                                If (Not IsTradeActive(chk_date, tradingSymbol)) AndAlso
                                    (TakeAllTrades OrElse (Not TakeAllTrades And GetTotalTrades(tr_date, tradingSymbol) < maxTrades)) Then
                                    Dim currentInstrumentXMinutePayload As Payload = XMinutePayload(tradingSymbol)(tr_date)
                                    Dim currentHAPayload As Payload = HAPayload(tradingSymbol)(tr_date)
                                    Dim currentATR As Decimal = ATRPaylaod(tradingSymbol)(tr_date)
                                    Dim currentVolume As Decimal = VolumePayload(tradingSymbol)(tr_date)
                                    Dim benchMarkWicksSize As Double = (currentInstrumentXMinutePayload.High - currentInstrumentXMinutePayload.Low) * currentCandleWicksSizePercentage / 100
                                    Dim runningTrade As Trade = Nothing
                                    If currentInstrumentXMinutePayload.CandleWicks.Top > benchMarkWicksSize Then
                                        If Not CandleRange Or (CandleRange AndAlso candleRangePercentage < currentInstrumentXMinutePayload.CandleRangePercentage) Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryPrice = currentInstrumentXMinutePayload.High + CalculateBuffer(currentInstrumentXMinutePayload.High)
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .TradingSymbol = tradingSymbol
                                                .TradingDate = currentInstrumentXMinutePayload.PayloadDate.Date
                                                .SignalCandle = currentInstrumentXMinutePayload
                                                .Quantity = stockList(tradingSymbol)
                                                .PotentialSL = .SignalCandle.Low - CalculateBuffer(.SignalCandle.Low)
                                                .PotentialTP = Math.Round((.EntryPrice + (.EntryPrice - .PotentialSL) * targetPercentage), 2)
                                                Dim pl As Double = CalculateProfitLoss(.TradingSymbol, .EntryPrice, .PotentialSL, .Quantity, TypeOfStock.Cash)
                                                If Math.Abs(pl) > .CapitalRequiredWithMargin * cuurentMaxStopLossPercentage / 100 Then
                                                    .PotentialSL = .EntryPrice - ((.CapitalRequiredWithMargin * cuurentMaxStopLossPercentage / 100) / .Quantity)
                                                    '.PotentialSL = CalculateTargetOrStoploss(.TradingSymbol, .EntryPrice, .Quantity, -(.CapitalRequiredWithMargin * cuurentMaxStopLossPercentage / 100), .EntryDirection, TypeOfStock.Cash)
                                                    .StopLossRemark = "Stop Loss moved because of Capital preservation"
                                                Else
                                                    .StopLossRemark = "Original"
                                                End If
                                                .AbsoluteATR = currentATR
                                                .ATRPercentage = .AbsoluteATR * 100 / .SignalCandle.Close
                                                .SignalCandleVolumeSMA = currentVolume
                                                .HeikenAshiSignalCandle = currentHAPayload
                                                .HeikenAshiSignalCandleStatus = GetHeikenAshiStatus(.HeikenAshiSignalCandle)
                                            End With
                                        End If
                                    ElseIf currentInstrumentXMinutePayload.CandleWicks.Bottom > benchMarkWicksSize Then
                                        If Not CandleRange Or (CandleRange AndAlso candleRangePercentage < currentInstrumentXMinutePayload.CandleRangePercentage) Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryPrice = currentInstrumentXMinutePayload.Low - CalculateBuffer(currentInstrumentXMinutePayload.Low)
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .TradingSymbol = tradingSymbol
                                                .TradingDate = currentInstrumentXMinutePayload.PayloadDate.Date
                                                .SignalCandle = currentInstrumentXMinutePayload
                                                .Quantity = stockList(tradingSymbol)
                                                .PotentialSL = .SignalCandle.High + CalculateBuffer(.SignalCandle.High)
                                                .PotentialTP = Math.Round((.EntryPrice - (.PotentialSL - .EntryPrice) * targetPercentage), 2)
                                                Dim pl As Double = CalculateProfitLoss(.TradingSymbol, .PotentialSL, .EntryPrice, .Quantity, TypeOfStock.Cash)
                                                If Math.Abs(pl) > .CapitalRequiredWithMargin * cuurentMaxStopLossPercentage / 100 Then
                                                    .PotentialSL = .EntryPrice + ((.CapitalRequiredWithMargin * cuurentMaxStopLossPercentage / 100) / .Quantity)
                                                    '.PotentialSL = CalculateTargetOrStoploss(.TradingSymbol, .EntryPrice, .Quantity, -(.CapitalRequiredWithMargin * cuurentMaxStopLossPercentage / 100), .EntryDirection, TypeOfStock.Cash)
                                                    .StopLossRemark = "Stop Loss moved because of Capital preservation"
                                                Else
                                                    .StopLossRemark = "Original"
                                                End If
                                                .AbsoluteATR = currentATR
                                                .ATRPercentage = .AbsoluteATR * 100 / .SignalCandle.Close
                                                .SignalCandleVolumeSMA = currentVolume
                                                .HeikenAshiSignalCandle = currentHAPayload
                                                .HeikenAshiSignalCandleStatus = GetHeikenAshiStatus(.HeikenAshiSignalCandle)
                                            End With
                                        End If
                                    End If
                                    If runningTrade IsNot Nothing Then EnterOrder(chk_date, tradingSymbol, runningTrade)
                                End If
                            End If
                        Next
                    End If

                    'OneMinuteLoop
                    Dim one_minute As DateTime = tr_time.AddMinutes(signalTimeFrame)
                    For minuteCtr = 0 To signalTimeFrame - 1
                        Dim oneMinutetrDate As DateTime = tr_date.AddMinutes(signalTimeFrame).AddMinutes(minuteCtr)
                        For Each tradingSymbol In stockList.Keys
                            If OneMinutePayload IsNot Nothing AndAlso XMinutePayload IsNot Nothing AndAlso
                                OneMinutePayload.ContainsKey(tradingSymbol) AndAlso
                                OneMinutePayload(tradingSymbol).ContainsKey(oneMinutetrDate) Then
                                Dim itemSpecificPayload As Payload = OneMinutePayload(tradingSymbol)(oneMinutetrDate)

                                Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, tradingSymbol, TradeExecutionStatus.Inprogress)
                                If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                    ExitTradeIfPossible(itemExitSpecificTrade(0), itemSpecificPayload, endOfDay, TypeOfStock.Cash, False)
                                End If
                                Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, tradingSymbol, TradeExecutionStatus.Open)
                                Dim tradeEntered As Boolean = False
                                If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                    tradeEntered = EnterTradeIfPossible(itemEntrySpecificTrade(0), itemSpecificPayload, TypeOfStock.Cash)
                                End If
                                If StopLossMoveToBreakEven = True Then
                                    If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                        If itemExitSpecificTrade(0).EntryDirection = TradeExecutionDirection.Buy Then
                                            Dim stopLossMoveTarget As Double = itemExitSpecificTrade(0).EntryPrice + (itemExitSpecificTrade(0).EntryPrice - itemExitSpecificTrade(0).PotentialSL)
                                            If itemSpecificPayload.High > stopLossMoveTarget Then
                                                MoveStopLoss(itemExitSpecificTrade(0), itemExitSpecificTrade(0).EntryPrice)
                                            End If
                                        ElseIf itemExitSpecificTrade(0).EntryDirection = TradeExecutionDirection.Sell Then
                                            Dim stopLossMoveTarget As Double = itemExitSpecificTrade(0).EntryPrice - (itemExitSpecificTrade(0).PotentialSL - itemExitSpecificTrade(0).EntryPrice)
                                            If itemSpecificPayload.Low < stopLossMoveTarget Then
                                                MoveStopLoss(itemExitSpecificTrade(0), itemExitSpecificTrade(0).EntryPrice)
                                            End If
                                        End If
                                    End If
                                End If
                                If minuteCtr = signalTimeFrame - 1 And itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 AndAlso Not tradeEntered Then
                                    CancelTrade(itemEntrySpecificTrade(0), itemSpecificPayload)
                                End If
                            End If
                        Next
                    Next
                    tr_time = tr_time.AddMinutes(signalTimeFrame)
                End While
            End If
            'PLforDay += GetPLForDay(chk_date.Date)
            chk_date = chk_date.AddDays(1)
        End While
        For Each stock In stockList.Keys
            Dim pl As Double = GetProfitLossForStock(stock, from_date, to_Date)
            Dim curDate As DateTime = System.DateTime.Now
            Dim filename As String = String.Format("MR {13}({9})-Candle Range-{10}, CandleWickSize%-{11}, MaxSL%-{12},TradeLimit-{7}, BreakEven-{8}, WR-{6} on {0}-{1}-{2}_{3}-{4}-{5}.xlsx",
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second,
                                                   WinRatio,
                                                   If(TakeAllTrades, "All", maxTrades),
                                                   If(StopLossMoveToBreakEven, "True", "False"),
                                                   Math.Round(pl, 0),
                                                   If(CandleRange, "True", "False"),
                                                    CandleWickSizePercentage, MaxStopLossPercentage, stock)
            PrintArrayToExcel(filename, stock)
        Next
    End Sub

    Private Function CalculateATR(tradingDate As Date, inputPayload As Dictionary(Of Date, Payload), candleTimeFrame As Integer) As Dictionary(Of Date, Decimal)
        Dim outputATR As Dictionary(Of Date, Decimal) = Nothing
        If inputPayload IsNot Nothing Then
            Dim tempOutputATR As Dictionary(Of Date, Decimal) = Nothing
            Dim xMinutePayload As Dictionary(Of Date, Payload) = Nothing
            If candleTimeFrame > 1 Then
                xMinutePayload = New Dictionary(Of Date, Payload)
                xMinutePayload = cmn.ConvertPayloadsToXMinutes(inputPayload, candleTimeFrame)
                If xMinutePayload IsNot Nothing Then
                    Indicator.ATR.CalculateATR(14, xMinutePayload, tempOutputATR)
                End If
            Else
                Indicator.ATR.CalculateATR(14, inputPayload, tempOutputATR)
            End If
            If tempOutputATR IsNot Nothing Then
                If outputATR Is Nothing Then outputATR = New Dictionary(Of Date, Decimal)
                For Each tempKeys In tempOutputATR.Keys
                    If tempKeys.Date = tradingDate.Date Then
                        outputATR.Add(tempKeys, tempOutputATR(tempKeys))
                    End If
                Next
            End If
        End If
        Return outputATR
    End Function

    Private Function CalculateMA(tradingDate As Date, inputPayload As Dictionary(Of Date, Payload), candleTimeFrame As Integer) As Dictionary(Of Date, Decimal)
        Dim outputPayload As Dictionary(Of Date, Decimal) = Nothing
        If inputPayload IsNot Nothing Then
            Dim tempOutputPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim xMinutePayload As Dictionary(Of Date, Payload) = Nothing
            If candleTimeFrame > 1 Then
                xMinutePayload = New Dictionary(Of Date, Payload)
                xMinutePayload = cmn.ConvertPayloadsToXMinutes(inputPayload, candleTimeFrame)
                If xMinutePayload IsNot Nothing Then
                    Indicator.SMA.CalculateSMA(25, Payload.PayloadFields.Volume, xMinutePayload, tempOutputPayload)
                End If
            Else
                Indicator.SMA.CalculateSMA(25, Payload.PayloadFields.Volume, inputPayload, tempOutputPayload)
            End If
            If tempOutputPayload IsNot Nothing Then
                If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Decimal)
                For Each tempKeys In tempOutputPayload.Keys
                    If tempKeys.Date = tradingDate.Date Then
                        outputPayload.Add(tempKeys, tempOutputPayload(tempKeys))
                    End If
                Next
            End If
        End If
        Return outputPayload
    End Function

    Private Function CalculateHeikenAshi(tradingDate As Date, inputPayload As Dictionary(Of Date, Payload), candleTimeFrame As Integer) As Dictionary(Of Date, Payload)
        Dim outputPayload As Dictionary(Of Date, Payload) = Nothing
        If inputPayload IsNot Nothing Then
            Dim tempOutputPayload As Dictionary(Of Date, Payload) = Nothing
            Dim xMinutePayload As Dictionary(Of Date, Payload) = Nothing
            If candleTimeFrame > 1 Then
                xMinutePayload = New Dictionary(Of Date, Payload)
                xMinutePayload = cmn.ConvertPayloadsToXMinutes(inputPayload, candleTimeFrame)
                If xMinutePayload IsNot Nothing Then
                    Indicator.HeikenAshi.ConvertToHeikenAshi(xMinutePayload, tempOutputPayload)
                End If
            Else
                Indicator.HeikenAshi.ConvertToHeikenAshi(inputPayload, tempOutputPayload)
            End If
            If tempOutputPayload IsNot Nothing Then
                If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Payload)
                For Each tempKeys In tempOutputPayload.Keys
                    If tempKeys.Date = tradingDate.Date Then
                        outputPayload.Add(tempKeys, tempOutputPayload(tempKeys))
                    End If
                Next
            End If
        End If
        Return outputPayload
    End Function
End Class
