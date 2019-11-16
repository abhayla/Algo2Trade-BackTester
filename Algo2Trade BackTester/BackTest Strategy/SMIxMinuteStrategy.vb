Imports System.Threading
Imports Algo2TradeBLL
Public Class SMIxMinuteStrategy
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)
    Public Sub Run(start_date As Date, end_date As Date, signalTimeFrame As Integer)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim mainSignalTimeFrame As Integer = signalTimeFrame
        Dim SMI_KPeriod As Integer = 10
        Dim SMI_KSmoothingPeriod As Integer = 3
        Dim SMI_KDoubleSmoothingPeriod As Integer = 3
        Dim SMI_DPeriod As Integer = 10
        Dim ATR_Period As Integer = 14
        Dim SMILowerLimit As Integer = -40
        Dim SMIUpperLimit As Integer = 40
        'Dim stopLossLimit As Decimal = 200
        'Dim targetLimit As Decimal = 10
        Dim targetATRMultiplier As Decimal = 2
        Dim stopLossATRMultiplier As Decimal = 1

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
        tempStocklist = New Dictionary(Of String, Decimal()) From {{"CRUDEOILM", {Nothing}}}
        For Each tradingSymbol In tempStocklist.Keys
            If stockList Is Nothing Then stockList = New Dictionary(Of String, Decimal())
            stockList.Add(tradingSymbol, {tempStocklist(tradingSymbol)(0)})
        Next

        While chk_date <= to_Date
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim SMISignalPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim ATRPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing

                For Each item In stockList.Keys
                    Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim tempXMinutePayload As Dictionary(Of Date, Payload) = Nothing

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
                                For Each tempKeys In tempXMinutePayload.Keys
                                    If tempKeys.Date = chk_date.Date Then
                                        If currentXMinutePayload Is Nothing Then currentXMinutePayload = New Dictionary(Of Date, Payload)
                                        currentXMinutePayload.Add(tempKeys, tempXMinutePayload(tempKeys))
                                    End If
                                Next
                                If currentXMinutePayload IsNot Nothing AndAlso currentXMinutePayload.Count > 0 Then
                                    Dim outputSMISignal As Dictionary(Of Date, Decimal) = Nothing
                                    Dim outputSMIEMASignal As Dictionary(Of Date, Decimal) = Nothing
                                    Indicator.SMI.CalculateSMI(SMI_KPeriod, SMI_KSmoothingPeriod, SMI_KDoubleSmoothingPeriod, SMI_DPeriod, tempXMinutePayload, outputSMISignal, outputSMIEMASignal)

                                    Dim outputATR As Dictionary(Of Date, Decimal) = Nothing
                                    Indicator.ATR.CalculateATR(ATR_Period, tempXMinutePayload, outputATR)

                                    If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    OneMinutePayload.Add(item, currentOneMinutePayload)
                                    If XMinutePayload Is Nothing Then XMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                    XMinutePayload.Add(item, currentXMinutePayload)
                                    If SMISignalPayload Is Nothing Then SMISignalPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                    SMISignalPayload.Add(item, outputSMISignal)
                                    If ATRPayload Is Nothing Then ATRPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                    ATRPayload.Add(item, outputATR)
                                End If
                            End If
                        End If
                    End If
                Next

                If XMinutePayload IsNot Nothing AndAlso XMinutePayload.Count > 0 Then
                    Dim startTime As Date = exchangeStartTime
                    Dim endTime As Date = exchangeEndTime
                    Dim lastTradeEntryTime As Date = exchangeEndTime.AddHours(-1)

                    While startTime < endTime
                        For Each stockName In stockList.Keys
                            OnHeartbeat(String.Format("Checking Trade for {0} on {1}", stockName, chk_date.ToShortDateString))
                            Dim plOfDay As Double = GetPLForDay(chk_date, stockName)
                            'If plOfDay >= targetLimit OrElse plOfDay <= -stopLossLimit Then
                            '    Exit While
                            'End If
                            Dim tempStockPayload As Dictionary(Of Date, Payload) = Nothing
                            If XMinutePayload.ContainsKey(stockName) Then
                                tempStockPayload = XMinutePayload(stockName)          'Only Current Day Payload
                            End If

                            If tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 Then
                                If startTime = exchangeStartTime Then
                                    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockPayload.Keys.FirstOrDefault.Hour, tempStockPayload.Keys.FirstOrDefault.Minute, tempStockPayload.Keys.FirstOrDefault.Second)
                                    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockPayload.Keys.LastOrDefault.Hour, tempStockPayload.Keys.LastOrDefault.Minute, tempStockPayload.Keys.LastOrDefault.Second)
                                    endTime = endTime.AddMinutes(mainSignalTimeFrame)
                                    endOfDay = endTime.AddMinutes(-30)
                                    lastTradeEntryTime = endTime.AddHours(-1)
                                End If
                                Dim tempStockSMISignalPayload As Dictionary(Of Date, Decimal) = SMISignalPayload(stockName)                 'Full Data
                                Dim tempStockATRPayload As Dictionary(Of Date, Decimal) = ATRPayload(stockName)
                                Dim potentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)
                                If Not tempStockPayload.ContainsKey(potentialSignalTime) Then
                                    Continue For
                                End If

                                Dim validSignal As Boolean = False
                                Dim signalDirection As TradeExecutionDirection = TradeExecutionDirection.None
                                If tempStockSMISignalPayload IsNot Nothing AndAlso tempStockSMISignalPayload.Count > 0 Then
                                    Dim currentPayload As Payload = tempStockPayload(potentialSignalTime)
                                    If startTime < lastTradeEntryTime AndAlso Not IsTradeActive(chk_date, stockName) Then
                                        If tempStockSMISignalPayload(currentPayload.PayloadDate) > SMILowerLimit AndAlso
                                            tempStockSMISignalPayload(currentPayload.PayloadDate) < SMIUpperLimit AndAlso
                                            tempStockSMISignalPayload(currentPayload.PreviousCandlePayload.PayloadDate) < SMILowerLimit Then
                                            validSignal = True
                                            signalDirection = TradeExecutionDirection.Buy
                                        ElseIf tempStockSMISignalPayload(currentPayload.PayloadDate) > SMILowerLimit AndAlso
                                                tempStockSMISignalPayload(currentPayload.PayloadDate) < SMIUpperLimit AndAlso
                                                tempStockSMISignalPayload(currentPayload.PreviousCandlePayload.PayloadDate) > SMIUpperLimit Then
                                            validSignal = True
                                            signalDirection = TradeExecutionDirection.Sell
                                        End If
                                    ElseIf IsTradeActive(chk_date, stockName) Then
                                        Dim reverseSignal As Boolean = False
                                        Dim reverseDirection As TradeExecutionDirection = TradeExecutionDirection.None
                                        Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                        If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                                            For Each item In itemSpecificTrade
                                                reverseDirection = item.EntryDirection
                                            Next
                                            If reverseDirection = TradeExecutionDirection.Buy Then
                                                If tempStockSMISignalPayload(currentPayload.PayloadDate) > SMIUpperLimit Then
                                                    reverseSignal = True
                                                End If
                                            ElseIf reverseDirection = TradeExecutionDirection.Sell Then
                                                If tempStockSMISignalPayload(currentPayload.PayloadDate) < SMILowerLimit Then
                                                    reverseSignal = True
                                                End If
                                            End If
                                            If reverseSignal Then
                                                For Each item In itemSpecificTrade
                                                    MoveStopLoss(item, currentPayload.Close)
                                                    ExitTradeIfPossible(item, tempStockPayload(potentialSignalTime.AddMinutes(mainSignalTimeFrame)), endOfDay, TypeOfStock.Commodity, True)
                                                Next
                                            End If
                                        End If

                                    End If
                                    Dim runningTrade As Trade = Nothing
                                    If validSignal Then
                                        If signalDirection = TradeExecutionDirection.Buy Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryPrice = currentPayload.Close
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .EntryTime = currentPayload.PayloadDate
                                                .SignalCandle = currentPayload
                                                .TradingSymbol = currentPayload.TradingSymbol
                                                .TradingDate = currentPayload.PayloadDate
                                                .Quantity = 60
                                                .PotentialTP = .EntryPrice + targetATRMultiplier * tempStockATRPayload(currentPayload.PayloadDate)
                                                .PotentialSL = .EntryPrice - stopLossATRMultiplier * tempStockATRPayload(currentPayload.PayloadDate)
                                                .MaximumDrawDown = .EntryPrice
                                                .MaximumDrawUp = .EntryPrice
                                            End With
                                        ElseIf signalDirection = TradeExecutionDirection.Sell Then
                                            runningTrade = New Trade
                                            With runningTrade
                                                .TradingStatus = TradeExecutionStatus.Open
                                                .EntryPrice = currentPayload.Close
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .EntryTime = currentPayload.PayloadDate
                                                .SignalCandle = currentPayload
                                                .TradingSymbol = currentPayload.TradingSymbol
                                                .TradingDate = currentPayload.PayloadDate
                                                .Quantity = 60
                                                .PotentialTP = .EntryPrice - targetATRMultiplier * tempStockATRPayload(currentPayload.PayloadDate)
                                                .PotentialSL = .EntryPrice + stopLossATRMultiplier * tempStockATRPayload(currentPayload.PayloadDate)
                                                .MaximumDrawDown = .EntryPrice
                                                .MaximumDrawUp = .EntryPrice
                                            End With
                                        End If
                                        If runningTrade IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade)
                                    End If
                                    If startTime <= endOfDay Then
                                        Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                        Dim tradeExit As Boolean = False
                                        If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                            For Each item In itemExitSpecificTrade
                                                tradeExit = ExitTradeIfPossible(item, currentPayload, endOfDay, TypeOfStock.Commodity, False)
                                            Next
                                        End If
                                        Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                        Dim tradeEntered As Boolean = False
                                        Dim itemEntrySpecificPayload As Payload = tempStockPayload(potentialSignalTime.AddMinutes(mainSignalTimeFrame))
                                        If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                            For Each item In itemEntrySpecificTrade
                                                tradeEntered = EnterTradeIfPossible(itemEntrySpecificTrade(0), itemEntrySpecificPayload, TypeOfStock.Commodity)
                                            Next
                                        End If
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
            Dim filename As String = String.Format("SMI X-Minute Strategy {0} for {1} Minute Time Frame {2}-{3}-{4}_{5}-{6}-{7}.xlsx", stock, signalTimeFrame,
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
            PrintArrayToExcel(filename, stock)
        Next
    End Sub
End Class
