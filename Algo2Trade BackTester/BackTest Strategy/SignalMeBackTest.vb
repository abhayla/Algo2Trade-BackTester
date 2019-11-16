Imports System.Threading
Imports Algo2TradeBLL
Public Class SignalMeBackTest
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)

    Public Sub Run(start_date As Date, end_date As Date, signalTimeFrame As Integer)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim mainSignalTimeFrame As Integer = signalTimeFrame
        Dim MACD_fastMAPeriod As Integer = 12
        Dim MACD_slowMAPeriod As Integer = 26
        Dim MACD_signalPeriod As Integer = 9
        Dim RSIPeriod As Integer = 4
        Dim CCIPeriod As Integer = 20
        Dim RSIOverBought As Double = 80
        Dim RSIOverSold As Double = 20
        Dim CCIOverBought As Double = 100
        Dim CCIOverSold As Double = -100
        Dim previousValidSignal As Boolean = False
        Dim previousSignalDirection As TradeExecutionDirection = TradeExecutionDirection.None

        'Change Time
        Dim exchangeStartTime As Date = "10:00:00"
        Dim exchangeEndTime As Date = "23:30:00"
        Dim endOfDay As DateTime = "23:00:00"

        Dim dataCheckDate As Date = end_date

        Dim from_date As DateTime = start_date
        Dim to_Date As Date = end_date
        Dim chk_date As Date = from_date
        Dim dateCtr As Integer = 0

        Dim stockList As Dictionary(Of String, Decimal()) = Nothing
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tempStocklist As Dictionary(Of String, Decimal()) = Nothing
        'TO DO:
        'Change Stock Name
        'tempStocklist = New Dictionary(Of String, Decimal()) From {{"ALUMINI", {1, 1.2, 0.1}}, {"COPPERM", {2, 2.2, 0.1}}, {"CRUDEOILM", {20, 32, 2}}, {"NATURALGAS", {1, 1.2, 0.1}}, {"NICKELM", {4, 6, 0.1}}, {"LEADMINI", {1, 1.2, 0.1}}, {"ZINCMINI", {1, 1.2, 0.1}}, {"COTTON", {100, 100, 0.1}}, {"GOLD", {10, 10, 0.1}}, {"SILVER", {200, 200, 20}}}        '{"StockName",{target,stoploss,buffer}}
        tempStocklist = New Dictionary(Of String, Decimal()) From {{"NATURALGAS", {1, 1.2, 0.1}}}        '{"StockName",{target,stoploss,buffer}}
        For Each tradingSymbol In tempStocklist.Keys
            If stockList Is Nothing Then stockList = New Dictionary(Of String, Decimal())
            stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0), tempStocklist(tradingSymbol)(1), tempStocklist(tradingSymbol)(2)})
        Next

        While chk_date <= to_Date
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinuteHKPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim MACDPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim MACDSignalPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim HistogramPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim RSIPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim CCIPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                For Each item In stockList.Keys
                    Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentXMinuteHKPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempXMinuteHKPayload As Dictionary(Of Date, Payload) = Nothing

                    'TO DO:
                    'Change fetching data from database table name & start date
                    Dim currentTradingSymbol As String = cmn.GetCurrentTradingSymbol(Common.DataBaseTable.EOD_Commodity, item, chk_date)
                    If PastIntradayData IsNot Nothing AndAlso PastIntradayData.ContainsKey(currentTradingSymbol) Then
                        tempOneMinutePayload = PastIntradayData(currentTradingSymbol)
                    Else
                        tempOneMinutePayload = cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Commodity, currentTradingSymbol, chk_date.AddMonths(-2), dataCheckDate)
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
                            tempXMinutePayload = cmn.ConvertPayloadsToXMinutes(tempOneMinutePayload, mainSignalTimeFrame)
                            If tempXMinutePayload IsNot Nothing AndAlso tempXMinutePayload.Count > 0 Then
                                Indicator.HeikenAshi.ConvertToHeikenAshi(tempXMinutePayload, tempXMinuteHKPayload)
                                If tempXMinuteHKPayload IsNot Nothing AndAlso tempXMinuteHKPayload.Count > 0 Then
                                    For Each tempKeys In tempXMinuteHKPayload.Keys
                                        If tempKeys.Date = chk_date.Date Then
                                            If currentXMinuteHKPayload Is Nothing Then currentXMinuteHKPayload = New Dictionary(Of Date, Payload)
                                            currentXMinuteHKPayload.Add(tempKeys, tempXMinuteHKPayload(tempKeys))
                                        End If
                                    Next
                                    If currentXMinuteHKPayload IsNot Nothing AndAlso currentXMinuteHKPayload.Count > 0 Then
                                        Dim ouputMACDPayload As Dictionary(Of Date, Decimal) = Nothing
                                        Dim outputsignalPayload As Dictionary(Of Date, Decimal) = Nothing
                                        Dim outputhistogramPayload As Dictionary(Of Date, Decimal) = Nothing
                                        Indicator.MACD.CalculateMACD(MACD_fastMAPeriod, MACD_slowMAPeriod, MACD_signalPeriod, tempXMinuteHKPayload, ouputMACDPayload, outputsignalPayload, outputhistogramPayload)

                                        Dim outputRSI As Dictionary(Of Date, Decimal) = Nothing
                                        Indicator.RSI.CalculateRSI(RSIPeriod, tempXMinuteHKPayload, outputRSI)

                                        Dim outputCCI As Dictionary(Of Date, Decimal) = Nothing
                                        Indicator.CCI.CalculateCCI(CCIPeriod, tempXMinuteHKPayload, outputCCI)

                                        If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        OneMinutePayload.Add(item, currentOneMinutePayload)
                                        If XMinutePayload Is Nothing Then XMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        XMinutePayload.Add(item, tempXMinutePayload)
                                        If XMinuteHKPayload Is Nothing Then XMinuteHKPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                        XMinuteHKPayload.Add(item, currentXMinuteHKPayload)
                                        If MACDPayload Is Nothing Then MACDPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        MACDPayload.Add(item, ouputMACDPayload)
                                        If MACDSignalPayload Is Nothing Then MACDSignalPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        MACDSignalPayload.Add(item, outputsignalPayload)
                                        If HistogramPayload Is Nothing Then HistogramPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        HistogramPayload.Add(item, outputhistogramPayload)
                                        If RSIPayload Is Nothing Then RSIPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        RSIPayload.Add(item, outputRSI)
                                        If CCIPayload Is Nothing Then CCIPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                        CCIPayload.Add(item, outputCCI)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next

                If XMinuteHKPayload IsNot Nothing AndAlso XMinuteHKPayload.Count > 0 Then
                    Dim startTime As Date = exchangeStartTime
                    Dim endTime As Date = exchangeEndTime
                    Dim lastTradeEntryTime As Date = exchangeEndTime.AddHours(-1)
                    Dim validSignal As Boolean = False
                    Dim reverseSignal As Boolean = False
                    While startTime < endTime
                        Dim oneMinuteStartTime As Date = Nothing
                        For Each stockName In stockList.Keys
                            OnHeartbeat(String.Format("Checking Trade for {0} on {1}", stockName, chk_date.ToShortDateString))
                            Dim tempStockPayload As Dictionary(Of Date, Payload) = Nothing
                            If XMinuteHKPayload.ContainsKey(stockName) Then
                                tempStockPayload = XMinuteHKPayload(stockName)          'Only Current Day Payload
                            End If

                            If tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 Then
                                If startTime = exchangeStartTime Then
                                    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockPayload.Keys.FirstOrDefault.Hour, tempStockPayload.Keys.FirstOrDefault.Minute, tempStockPayload.Keys.FirstOrDefault.Second)
                                    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockPayload.Keys.LastOrDefault.Hour, tempStockPayload.Keys.LastOrDefault.Minute, tempStockPayload.Keys.LastOrDefault.Second)
                                    endTime = endTime.AddMinutes(mainSignalTimeFrame)
                                    endOfDay = endTime.AddMinutes(-30)
                                    lastTradeEntryTime = endTime.AddHours(-1)
                                End If
                                oneMinuteStartTime = startTime
                                Dim tempStockMACDPayload As Dictionary(Of Date, Decimal) = MACDPayload(stockName)               'Full Data
                                Dim tempStockMACDSignalPayload As Dictionary(Of Date, Decimal) = MACDSignalPayload(stockName)   'Full Data
                                Dim tempStockHistogramPayload As Dictionary(Of Date, Decimal) = HistogramPayload(stockName)     'Full Data
                                Dim tempStockRSIPayload As Dictionary(Of Date, Decimal) = RSIPayload(stockName)                 'Full Data
                                Dim tempStockCCIPayload As Dictionary(Of Date, Decimal) = CCIPayload(stockName)                 'Full Data

                                Dim potentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)
                                If Not tempStockPayload.ContainsKey(potentialSignalTime) Then
                                    Continue For
                                End If

                                Dim signalDirection As TradeExecutionDirection = TradeExecutionDirection.Buy
                                Dim currentSignalPayload As Payload = tempStockPayload(potentialSignalTime)

                                'Checking for Valid Signal Start
                                If startTime < lastTradeEntryTime AndAlso Not IsTradeActive(chk_date, stockName) Then
                                    validSignal = True
                                    validSignal = validSignal And GetValidSignal(previousValidSignal, currentSignalPayload, signalDirection,
                                                                               tempStockMACDPayload, tempStockMACDSignalPayload, tempStockHistogramPayload,
                                                                               tempStockRSIPayload, RSIOverBought, RSIOverSold,
                                                                               tempStockCCIPayload, CCIOverBought, CCIOverSold)
                                    If validSignal = False Then
                                        validSignal = True
                                        signalDirection = TradeExecutionDirection.Sell
                                        validSignal = validSignal And GetValidSignal(previousValidSignal, currentSignalPayload, signalDirection,
                                                                               tempStockMACDPayload, tempStockMACDSignalPayload, tempStockHistogramPayload,
                                                                               tempStockRSIPayload, RSIOverBought, RSIOverSold,
                                                                               tempStockCCIPayload, CCIOverBought, CCIOverSold)
                                    End If
                                ElseIf IsTradeActive(chk_date, stockName) Then
                                    'Check Reverse Signal
                                    Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                                        For Each item In itemSpecificTrade
                                            signalDirection = item.EntryDirection
                                        Next
                                        reverseSignal = GetValidSignal(previousValidSignal, currentSignalPayload, signalDirection,
                                                                               tempStockMACDPayload, tempStockMACDSignalPayload, tempStockHistogramPayload,
                                                                               tempStockRSIPayload, RSIOverBought, RSIOverSold,
                                                                               tempStockCCIPayload, CCIOverBought, CCIOverSold)

                                        If reverseSignal = False Then
                                            For Each item In itemSpecificTrade
                                                signalDirection = If(item.EntryDirection = TradeExecutionDirection.Buy, TradeExecutionDirection.Sell, TradeExecutionDirection.Buy)
                                            Next
                                            reverseSignal = GetValidSignal(False, currentSignalPayload, signalDirection,
                                                                               tempStockMACDPayload, tempStockMACDSignalPayload, tempStockHistogramPayload,
                                                                               tempStockRSIPayload, RSIOverBought, RSIOverSold,
                                                                               tempStockCCIPayload, CCIOverBought, CCIOverSold)
                                            If reverseSignal Then
                                                validSignal = True
                                                previousValidSignal = False
                                            End If
                                        End If
                                    End If
                                End If
                                'Checking for Valid Signal End

                                Dim runningTrade As Trade = Nothing
                                'Entry Trade Condition Start
                                If validSignal AndAlso validSignal <> previousValidSignal AndAlso signalDirection = TradeExecutionDirection.Buy Then
                                    runningTrade = New Trade
                                    With runningTrade
                                        .TradingStatus = TradeExecutionStatus.Open
                                        .EntryPrice = currentSignalPayload.High + stockList(stockName)(2)
                                        .EntryDirection = TradeExecutionDirection.Buy
                                        .EntryTime = currentSignalPayload.PayloadDate
                                        .SignalCandle = currentSignalPayload
                                        .TradingSymbol = currentSignalPayload.TradingSymbol
                                        .TradingDate = currentSignalPayload.PayloadDate
                                        .Quantity = 1
                                        .PotentialTP = .EntryPrice + stockList(stockName)(0)
                                        .PotentialSL = .EntryPrice - stockList(stockName)(1)
                                        .MaximumDrawDown = .EntryPrice
                                        .MaximumDrawUp = .EntryPrice
                                    End With
                                ElseIf validSignal AndAlso validSignal <> previousValidSignal AndAlso signalDirection = TradeExecutionDirection.Sell Then
                                    runningTrade = New Trade
                                    With runningTrade
                                        .TradingStatus = TradeExecutionStatus.Open
                                        .EntryPrice = currentSignalPayload.Low - stockList(stockName)(2)
                                        .EntryDirection = TradeExecutionDirection.Sell
                                        .EntryTime = currentSignalPayload.PayloadDate
                                        .SignalCandle = currentSignalPayload
                                        .TradingSymbol = currentSignalPayload.TradingSymbol
                                        .TradingDate = currentSignalPayload.PayloadDate
                                        .Quantity = 1
                                        .PotentialTP = .EntryPrice - stockList(stockName)(0)
                                        .PotentialSL = .EntryPrice + stockList(stockName)(1)
                                        .MaximumDrawDown = .EntryPrice
                                        .MaximumDrawUp = .EntryPrice
                                    End With
                                End If
                                If runningTrade IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade)
                                'Entry Trade Condition End
                            End If

                            If (previousValidSignal And Not validSignal) OrElse startTime >= endOfDay Then
                                Dim signalCancelMinute As Date = startTime.AddMinutes(mainSignalTimeFrame)
                                signalCancelMinute = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, signalCancelMinute.Hour, signalCancelMinute.Minute, signalCancelMinute.Second)
                                If OneMinutePayload IsNot Nothing AndAlso OneMinutePayload.Count > 0 AndAlso
                                       OneMinutePayload.ContainsKey(stockName) AndAlso
                                       OneMinutePayload(stockName).ContainsKey(signalCancelMinute) Then

                                    Dim itemCancelSpecificPayload As Payload = OneMinutePayload(stockName)(signalCancelMinute)
                                    Dim itemCancelSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                    If itemCancelSpecificTrade IsNot Nothing AndAlso itemCancelSpecificTrade.Count > 0 Then
                                        For Each item In itemCancelSpecificTrade
                                            CancelTrade(item, itemCancelSpecificPayload)
                                        Next
                                    End If
                                End If
                            End If
                        Next

                        'One Minute Loop Start
                        If startTime <= endOfDay Then
                            Dim oneMinute As Date = startTime.AddMinutes(mainSignalTimeFrame)
                            For minuteCtr = 0 To mainSignalTimeFrame - 1
                                Dim oneMinutePotentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, oneMinute.Hour, oneMinute.Minute, oneMinute.Second)
                                oneMinutePotentialSignalTime = oneMinutePotentialSignalTime.AddMinutes(minuteCtr)
                                For Each stockName In stockList.Keys
                                    If OneMinutePayload IsNot Nothing AndAlso OneMinutePayload.Count > 0 AndAlso
                                        OneMinutePayload.ContainsKey(stockName) AndAlso
                                        OneMinutePayload(stockName).ContainsKey(oneMinutePotentialSignalTime) Then

                                        Dim itemSpecificPayload As Payload = OneMinutePayload(stockName)(oneMinutePotentialSignalTime)

                                        Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                        Dim tradeExit As Boolean = False
                                        If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                            For Each item In itemExitSpecificTrade
                                                tradeExit = ExitTradeIfPossible(item, itemSpecificPayload, endOfDay, TypeOfStock.Commodity, False)
                                            Next
                                        End If
                                        Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                        Dim tradeEntered As Boolean = False
                                        If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                            For Each item In itemEntrySpecificTrade
                                                tradeEntered = EnterTradeIfPossible(itemEntrySpecificTrade(0), itemSpecificPayload, TypeOfStock.Commodity)
                                            Next
                                        End If
                                    End If
                                Next
                            Next
                        End If

                        'One Minute Loop End
                        previousValidSignal = validSignal
                        startTime = startTime.AddMinutes(mainSignalTimeFrame)
                    End While
                End If
            End If
            chk_date = chk_date.AddDays(1)
        End While
        For Each stock In stockList.Keys
            Dim curDate As Date = System.DateTime.Now
            Dim filename As String = String.Format("Signal Me Backtest {0} for {1} Minute Time Frame {2}-{3}-{4}_{5}-{6}-{7}.xlsx", stock, signalTimeFrame,
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
            PrintArrayToExcel(filename, stock)
        Next
    End Sub

    Private Function GetValidSignal(ByVal previousValidSignal As Boolean, ByVal currentPayload As Payload, ByVal signalDirection As TradeExecutionDirection, ParamArray requiredIndicators() As Object) As Boolean
        Dim ret As Boolean = False
        Dim MACDState As IndicatorState = IndicatorState.None
        Dim RSIState As IndicatorState = IndicatorState.None
        Dim CCIState As IndicatorState = IndicatorState.None

        'MACD
        If requiredIndicators(0)(currentPayload.PayloadDate) > requiredIndicators(1)(currentPayload.PayloadDate) AndAlso
            requiredIndicators(1)(currentPayload.PayloadDate) > 0 Then

            MACDState = IndicatorState.CrossingUp
        ElseIf requiredIndicators(0)(currentPayload.PayloadDate) < requiredIndicators(1)(currentPayload.PayloadDate) AndAlso
                requiredIndicators(1)(currentPayload.PayloadDate) < 0 Then

            MACDState = IndicatorState.CrossingDown
        End If

        'RSI
        If Not previousValidSignal Then
            If requiredIndicators(3)(currentPayload.PayloadDate) > requiredIndicators(4) AndAlso
            requiredIndicators(3)(currentPayload.PreviousCandlePayload.PayloadDate) < requiredIndicators(4) Then

                RSIState = IndicatorState.CrossingUp
            ElseIf requiredIndicators(3)(currentPayload.PayloadDate) < requiredIndicators(5) AndAlso
                requiredIndicators(3)(currentPayload.PreviousCandlePayload.PayloadDate) > requiredIndicators(5) Then

                RSIState = IndicatorState.CrossingDown
            End If
        Else
            If requiredIndicators(3)(currentPayload.PayloadDate) > requiredIndicators(4) Then
                RSIState = IndicatorState.CrossingUp
            ElseIf requiredIndicators(3)(currentPayload.PayloadDate) < requiredIndicators(5) Then
                RSIState = IndicatorState.CrossingDown
            End If
        End If

        'CCI
        If requiredIndicators(6)(currentPayload.PayloadDate) > requiredIndicators(7) Then
            CCIState = IndicatorState.Overbought
        ElseIf requiredIndicators(6)(currentPayload.PayloadDate) < requiredIndicators(8) Then
            CCIState = IndicatorState.Oversold
        End If

        'Entry Condition
        If RSIState = IndicatorState.CrossingUp And CCIState = IndicatorState.Overbought And MACDState = IndicatorState.CrossingUp Then
            If signalDirection = TradeExecutionDirection.Buy Then
                ret = True
            End If
        ElseIf RSIState = IndicatorState.CrossingDown And CCIState = IndicatorState.Oversold And MACDState = IndicatorState.CrossingDown Then
            If signalDirection = TradeExecutionDirection.Sell Then
                ret = True
            End If
        End If
        Return ret
    End Function
    Public Enum IndicatorState
        None = 1
        Overbought
        Oversold
        Outside
        CrossingUp
        CrossingDown
        ROverboughtSold
    End Enum
End Class
