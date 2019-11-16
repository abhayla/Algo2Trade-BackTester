Imports System.Threading
Imports Algo2TradeBLL
Imports MySql.Data.MySqlClient
Public Class Nifty50StockBreakingPreviousDayHighLow
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)
    Public Sub Run(start_date As Date, end_date As Date)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim mainSignalTimeFrame As Integer = 15

        'TODO:
        'Change Time
        Dim exchangeStartTime As Date = "09:15:00"
        Dim exchangeEndTime As Date = "15:30:00"
        Dim endOfDay As DateTime = "15:15:00"
        Dim lastTradeEntryTime As Date = "09:45:00"

        Dim dataCheckDate As Date = end_date

        Dim from_date As DateTime = start_date
        Dim to_Date As Date = end_date
        Dim chk_date As Date = from_date
        Dim dateCtr As Integer = 0

        Dim stockList As Dictionary(Of String, Decimal()) = Nothing
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tempStocklist As String() = {"ADANIPORTS",
                                        "ASIANPAINT",
                                        "AXISBANK",
                                        "BAJAJ-AUTO",
                                        "BAJAJFINSV",
                                        "BAJFINANCE",
                                        "BHARTIARTL",
                                        "BPCL",
                                        "CIPLA",
                                        "COALINDIA",
                                        "DRREDDY",
                                        "EICHERMOT",
                                        "GAIL",
                                        "GRASIM",
                                        "HCLTECH",
                                        "HDFC",
                                        "HDFCBANK",
                                        "HEROMOTOCO",
                                        "HINDALCO",
                                        "HINDPETRO",
                                        "HINDUNILVR",
                                        "IBULHSGFIN",
                                        "ICICIBANK",
                                        "INDUSINDBK",
                                        "INFRATEL",
                                        "INFY",
                                        "IOC",
                                        "ITC",
                                        "JSWSTEEL",
                                        "KOTAKBANK",
                                        "LT",
                                        "M&M",
                                        "MARUTI",
                                        "NTPC",
                                        "ONGC",
                                        "POWERGRID",
                                        "RELIANCE",
                                        "SBIN",
                                        "SUNPHARMA",
                                        "TATAMOTORS",
                                        "TATASTEEL",
                                        "TCS",
                                        "TECHM",
                                        "TITAN",
                                        "ULTRACEMCO",
                                        "UPL",
                                        "VEDL",
                                        "WIPRO",
                                        "YESBANK",
                                        "ZEEL"}
        'TODO:
        'Change Stock Name
        Dim overAllStockList As Dictionary(Of String, Decimal()) = Nothing
        For Each tradingSymbol In tempStocklist
            If overAllStockList Is Nothing Then overAllStockList = New Dictionary(Of String, Decimal())
            overAllStockList.Add(tradingSymbol, {Nothing})
        Next

        While chk_date <= to_Date
            Dim activeCapital As Double = 0
            Dim passiveCapital As Double = 0
            Dim plDrawdown As Double = 0
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))

            stockList = GetPrevious5DaysAverageVolume(chk_date, overAllStockList)

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim HighBreakStock As Dictionary(Of String, Decimal()) = Nothing
                Dim LowBreakStock As Dictionary(Of String, Decimal()) = Nothing
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim EODPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing

                For Each item In stockList.Keys
                    Dim currentOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim previousDayEODPayload As Dictionary(Of Date, Payload) = Nothing

                    'TODO:
                    'Change fetching data from database table name & start date
                    Dim previousTradingDay As Date = cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, item, chk_date)
                    currentOneMinutePayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, item, chk_date, chk_date)
                    previousDayEODPayload = cmn.GetRawPayload(Common.DataBaseTable.EOD_Cash, item, previousTradingDay, previousTradingDay)

                    OnHeartbeat(String.Format("Processing Data for {0}", chk_date.ToShortDateString))

                    If currentOneMinutePayload IsNot Nothing AndAlso currentOneMinutePayload.Count > 0 Then
                        currentXMinutePayload = cmn.ConvertPayloadsToXMinutes(currentOneMinutePayload, 15)

                        If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        OneMinutePayload.Add(item, currentOneMinutePayload)
                        If XMinutePayload Is Nothing Then XMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        XMinutePayload.Add(item, currentXMinutePayload)
                        If EODPayload Is Nothing Then EODPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        EODPayload.Add(item, previousDayEODPayload)

                        Dim runningTrade As Trade = Nothing
                        For Each candle In currentXMinutePayload.Keys
                            If currentXMinutePayload(candle).High > previousDayEODPayload.LastOrDefault.Value.High Then
                                If HighBreakStock Is Nothing Then HighBreakStock = New Dictionary(Of String, Decimal())
                                HighBreakStock.Add(item, {Nothing})
                                runningTrade = New Trade
                                With runningTrade
                                    .TradingStatus = TradeExecutionStatus.Open
                                    .EntryPrice = currentXMinutePayload(candle).High + 0.1
                                    .EntryDirection = TradeExecutionDirection.Buy
                                    .EntryTime = currentXMinutePayload(candle).PayloadDate
                                    .SignalCandle = currentXMinutePayload(candle)
                                    .TradingSymbol = currentXMinutePayload(candle).TradingSymbol
                                    .TradingDate = currentXMinutePayload(candle).PayloadDate
                                    '.PotentialSL = currentXMinutePayload(candle).Low - 0.1
                                    '.PotentialTP = If(currentXMinutePayload(candle).CandleRange > 0.01 * .EntryPrice, .EntryPrice + currentXMinutePayload(candle).CandleRange * 0.5, .EntryPrice + .EntryPrice * 0.005)
                                    .PotentialSL = .EntryPrice - .EntryPrice * 0.005
                                    .PotentialTP = .EntryPrice + .EntryPrice * 0.005
                                    .Quantity = CalculateQuantity(.TradingSymbol, .EntryPrice, .PotentialSL, -3000, TypeOfStock.Cash)
                                    .MaximumDrawDown = .EntryPrice
                                    .MaximumDrawUp = .EntryPrice
                                End With
                            ElseIf currentXMinutePayload(candle).Low < previousDayEODPayload.LastOrDefault.Value.Low Then
                                If LowBreakStock Is Nothing Then LowBreakStock = New Dictionary(Of String, Decimal())
                                LowBreakStock.Add(item, {Nothing})
                                runningTrade = New Trade
                                With runningTrade
                                    .TradingStatus = TradeExecutionStatus.Open
                                    .EntryPrice = currentXMinutePayload(candle).Low - 0.1
                                    .EntryDirection = TradeExecutionDirection.Sell
                                    .EntryTime = currentXMinutePayload(candle).PayloadDate
                                    .SignalCandle = currentXMinutePayload(candle)
                                    .TradingSymbol = currentXMinutePayload(candle).TradingSymbol
                                    .TradingDate = currentXMinutePayload(candle).PayloadDate
                                    '.PotentialSL = currentXMinutePayload(candle).High + 0.1
                                    '.PotentialTP = If(currentXMinutePayload(candle).CandleRange > 0.01 * .EntryPrice, .EntryPrice - currentXMinutePayload(candle).CandleRange * 0.5, .EntryPrice - .EntryPrice * 0.005)
                                    .PotentialSL = .EntryPrice + .EntryPrice * 0.005
                                    .PotentialTP = .EntryPrice - .EntryPrice * 0.005
                                    .Quantity = CalculateQuantity(.TradingSymbol, .PotentialSL, .EntryPrice, -3000, TypeOfStock.Cash)
                                    .MaximumDrawDown = .EntryPrice
                                    .MaximumDrawUp = .EntryPrice
                                End With
                            End If
                            If runningTrade IsNot Nothing Then
                                EnterOrder(chk_date, item, runningTrade)
                                passiveCapital += runningTrade.CapitalRequiredWithMargin
                            End If
                            Exit For
                        Next
                    End If
                Next

                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(chk_date) Then
                    Dim startTime As Date = exchangeStartTime
                    Dim endTime As Date = exchangeEndTime
                    Dim firstCandleOfDay As Boolean = True
                    Dim exitTrade As Boolean = False
                    While startTime < endTime
                        OnHeartbeat(String.Format("Checking Trade for {0}", chk_date.ToShortDateString))
                        Dim tempStockPayload As Dictionary(Of Date, Payload) = Nothing
                        For Each stockName In TradesTaken(chk_date).Keys
                            If OneMinutePayload.ContainsKey(stockName) Then
                                tempStockPayload = OneMinutePayload(stockName)          'Only Current Day Payload
                            End If

                            If tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 Then
                                If startTime = exchangeStartTime Then
                                    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockPayload.Keys.FirstOrDefault.Hour, tempStockPayload.Keys.FirstOrDefault.Minute, tempStockPayload.Keys.FirstOrDefault.Second)
                                    startTime = startTime.AddMinutes(mainSignalTimeFrame)
                                    lastTradeEntryTime = startTime.AddMinutes(mainSignalTimeFrame)
                                    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockPayload.Keys.LastOrDefault.Hour, tempStockPayload.Keys.LastOrDefault.Minute, tempStockPayload.Keys.LastOrDefault.Second)
                                    endTime = endTime.AddMinutes(1)
                                    endOfDay = endTime.AddMinutes(-15)
                                End If

                                'OneMinuteLoop
                                Dim oneMinuteCandleTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)
                                If startTime < endTime Then
                                    'For Each stock In TradesTaken(chk_date).Keys
                                    If OneMinutePayload IsNot Nothing AndAlso OneMinutePayload.ContainsKey(stockName) Then
                                        Dim itemSpecificPayload As Payload = Nothing
                                        itemSpecificPayload = OneMinutePayload(stockName)(oneMinuteCandleTime)

                                        Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                        If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                            For Each item In itemExitSpecificTrade
                                                If item.EntryDirection = TradeExecutionDirection.Buy Then
                                                    item.DrawDownPL = CalculateProfitLoss(item.TradingSymbol, item.EntryPrice, itemSpecificPayload.Low, item.Quantity, TypeOfStock.Cash)
                                                ElseIf item.EntryDirection = TradeExecutionDirection.Sell Then
                                                    item.DrawDownPL = CalculateProfitLoss(item.TradingSymbol, itemSpecificPayload.High, item.EntryPrice, item.Quantity, TypeOfStock.Cash)
                                                End If
                                                ExitTradeIfPossible(item, itemSpecificPayload, endOfDay, TypeOfStock.Cash, False)
                                            Next
                                        End If

                                        If startTime < lastTradeEntryTime Then
                                            Dim itemEntrySpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                            If itemEntrySpecificTrade IsNot Nothing AndAlso itemEntrySpecificTrade.Count > 0 Then
                                                For Each item In itemEntrySpecificTrade
                                                    If EnterTradeIfPossible(item, itemSpecificPayload, TypeOfStock.Cash) Then
                                                        activeCapital += item.CapitalRequiredWithMargin
                                                    End If
                                                Next
                                            End If
                                        End If

                                        If startTime = lastTradeEntryTime Then
                                            Dim itemCancelSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Open)
                                            If itemCancelSpecificTrade IsNot Nothing AndAlso itemCancelSpecificTrade.Count > 0 Then
                                                For Each item In itemCancelSpecificTrade
                                                    CancelTrade(item, itemSpecificPayload)
                                                Next
                                            End If
                                        End If
                                    End If
                                    'Next
                                End If
                            End If
                        Next
                        Dim minuteDrawdownPL As Double = 0
                        For Each stockName In TradesTaken(chk_date).Keys
                            For Each stockTrades In TradesTaken(chk_date)(stockName)
                                If stockTrades.TradingStatus = TradeExecutionStatus.Inprogress Then
                                    minuteDrawdownPL += stockTrades.DrawDownPL
                                End If
                            Next
                        Next
                        If plDrawdown = 0 Then
                            plDrawdown = minuteDrawdownPL
                        Else
                            plDrawdown = Math.Min(plDrawdown, minuteDrawdownPL)
                        End If
                        startTime = startTime.AddMinutes(1)
                    End While
                End If
            End If
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(chk_date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(chk_date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        For Each stockTrade In stockTrades(stock)
                            stockTrade.PassiveCapitalWithMargin = passiveCapital
                            stockTrade.ActiveCapitalWithMargin = activeCapital
                            stockTrade.DrawDownPL = plDrawdown
                        Next
                    Next
                End If
            End If
            chk_date = chk_date.AddDays(1)
        End While
        Dim curDate As Date = System.DateTime.Now
        Dim filename As String = String.Format("NIFTY50 Stocks Strategy {0}-{1}-{2}_{3}-{4}-{5}.xlsx",
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
        PrintArrayToExcel(filename)
    End Sub
    Private Function GetPrevious5DaysAverageVolume(currentDate As Date, stockList As Dictionary(Of String, Decimal())) As Dictionary(Of String, Decimal())
        OnHeartbeat("Checking Average Volume for NIFTY50 Stocks")
        Dim ret As Dictionary(Of String, Decimal()) = Nothing
        Dim dt As DataTable = Nothing
        Dim conn As MySqlConnection = cmn.OpenDBConnection()
        Dim cmd As MySqlCommand = Nothing
        If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
            Dim stocks As String = Nothing
            For Each item In stockList.Keys
                If stocks Is Nothing Then
                    stocks = String.Format("'{0}'", item)
                Else
                    stocks = String.Format("{0},'{1}'", stocks, item)
                End If
            Next
            'stocks = Left(stocks, Len(stocks) - 1)
            cmd = New MySqlCommand(String.Format("SELECT  a.`TradingSymbol`, AVG(a.`Volume`)
                                FROM (SELECT  `TradingSymbol`,`Volume`,`SnapshotDate`,
                                 (CASE TradingSymbol
                                 WHEN @curTradingSymbol COLLATE utf8_unicode_ci THEN @curRow := @curRow + 1
                                 ELSE @curRow := 1 AND @curTradingSymbol := TradingSymbol END 
                                 )+1 AS rank
                                 FROM `eod_prices_cash`,
                                 (SELECT @curRow :=0 , @curTradingSymbol :='' COLLATE utf8_unicode_ci) r
                                 WHERE TradingSymbol IN ({0})
                                 AND `SnapshotDate`<=@cd
                                 AND `SnapshotDate`>=DATE_SUB(@cd,INTERVAL 15 DAY)
                                 ORDER BY TradingSymbol,`SnapshotDate` DESC) a
                                WHERE a.rank<=5
                                GROUP BY a.`TradingSymbol`", stocks), conn)


            'cmd.Parameters.AddWithValue("@trd", stocks)
            cmd.Parameters.AddWithValue("@cd", currentDate)
            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 300
            dt = New DataTable()
            adapter.Fill(dt)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim count As Integer = 0
                While Not count = dt.Rows.Count
                    If dt.Rows(count).Item(1) > 1000000 Then
                        If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
                        ret.Add(dt.Rows(count).Item(0), {dt.Rows(count).Item(1)})
                    End If
                    count += 1
                End While
            End If
        End If
        Return ret
    End Function
End Class
