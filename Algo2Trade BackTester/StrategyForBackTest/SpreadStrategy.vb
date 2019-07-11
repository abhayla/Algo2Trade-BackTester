Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading
Imports MySql.Data.MySqlClient
Imports System.IO

Public Class SpreadStrategy
    Inherits Strategy
    Implements IDisposable

    Private ReadOnly _SignalTimeFrame As Integer
    Private ReadOnly _UseHeikenAshi As Boolean
    Public Property PartialExit As Boolean = False
    Public Property NumberOfTradePerStockPerDay As Integer = Integer.MaxValue
    Public Property NumberOfTradePerDay As Integer = Integer.MaxValue
    Public Property TrailingSL As Boolean = False
    Public Property ReverseSignalTrade As Boolean = False
    Public Property ExitOnStockFixedTargetStoploss As Boolean = False
    Public Property StockMaxProfitPerDay As Double = 15000
    Public Property StockMaxLossPerDay As Double = -10000
    Public Property ExitOnOverAllFixedTargetStoploss As Boolean = False
    Public Property ModifyTargetStoploss As Boolean = False
    Public Property NIFTY50Stocks As String()

    Private ReadOnly _StockData As Dictionary(Of String, Decimal())
    Private ReadOnly _nifty50FilePath As String = Path.Combine(My.Application.Info.DirectoryPath, "NIFTY50 Stocks.txt")
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan,
                   ByVal signalTimeFrame As Integer,
                   ByVal UseHeikenAshi As Boolean,
                   ByVal associatedStockData As Dictionary(Of String, Decimal()))
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
        Me._SignalTimeFrame = signalTimeFrame
        Me._UseHeikenAshi = UseHeikenAshi
        Me._StockData = associatedStockData
        Me.NIFTY50Stocks = File.ReadAllLines(_nifty50FilePath)
    End Sub
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tradeStockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
        Dim databaseTable As Common.DataBaseTable = Common.DataBaseTable.Intraday_Futures
        Dim totalPL As Decimal = 0
        Dim tradeCheckingDate As Date = startDate
        While tradeCheckingDate <= endDate
            Dim stockList As Dictionary(Of String, Decimal()) = Nothing
            'Dim tempStockList As IEnumerable(Of KeyValuePair(Of String, Integer)) = GetStockData(tradeCheckingDate).Take(1)
            'stockList = tempStockList.ToDictionary(Of String, Integer)(Function(x)
            '                                                               Return x.Key
            '                                                           End Function, Function(y)
            '                                                                             Return y.Value
            '                                                                         End Function)
            If Me._StockData IsNot Nothing AndAlso Me._StockData.Count > 0 Then
                stockList = New Dictionary(Of String, Decimal())
                Dim lotSize As Integer = Cmn.GetAppropiateLotSize(databaseTable, _StockData.FirstOrDefault.Key, tradeCheckingDate.Date)
                stockList.Add(_StockData.FirstOrDefault.Key, {lotSize, lotSize})
            Else
                'stockList = GetStockData(tradeCheckingDate)
            End If
            stockList = New Dictionary(Of String, Decimal())
            'stockList.Add("StockName", {Quantity, LotSize})
            'stockList.Add("CRUDEOIL19APRFUT", {1000, 100})
            'stockList.Add("CRUDEOIL19MAYFUT", {1000, 100})
            stockList.Add("BANKNIFTY19JANFUT", {200, 20})
            stockList.Add("BANKNIFTY19FEBFUT", {200, 20})
            'stockList.Add("JINDALSTEL", {2250, 2250})
            'stockList.Add("ADANIENT", {4000, 4000})
            'stockList.Add("ZEEL", {1300, 1, -1})
            'stockList.Add("CRUDEOIL", {100, 100})
            'stockList.Add("NIFTY", {75, 1})

            If stockList IsNot Nothing AndAlso stockList.Count = 2 Then
                Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XDayXMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XDayXMinuteHAStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XDayRuleSignalStocksPayload As Dictionary(Of String, Dictionary(Of Date, Integer)) = Nothing
                Dim XDayRuleEntryStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim XDayRuleTargetStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim XDayRuleStoplossStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim XDayRuleQuantityStocksPayload As Dictionary(Of String, Dictionary(Of Date, Integer)) = Nothing
                Dim XDayRuleSupporting1StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting2StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting3StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting4StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting5StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting6StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting7StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting8StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting9StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                Dim XDayRuleSupporting10StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing

                'First lets build the payload for all the stocks
                Dim stockCount As Integer = 0
                For Each stock In stockList.Keys
                    stockCount += 1
                    Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim XDayXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim XDayXMinuteHAPayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing

                    'Get payload
                    XDayOneMinutePayload = Cmn.GetRawPayloadForSpecificTradingSymbol(databaseTable, stock, tradeCheckingDate.AddDays(-4), tradeCheckingDate)

                    'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)
                    If XDayOneMinutePayload IsNot Nothing AndAlso XDayOneMinutePayload.Count > 0 Then
                        OnHeartbeat(String.Format("Processing for {0} on {1}. Stock Counter: [ {2}/{3} ]", stock, tradeCheckingDate.ToShortDateString, stockCount, stockList.Count))
                        For Each runningPayload In XDayOneMinutePayload.Keys
                            If runningPayload.Date = tradeCheckingDate.Date Then
                                If currentDayOneMinutePayload Is Nothing Then currentDayOneMinutePayload = New Dictionary(Of Date, Payload)
                                currentDayOneMinutePayload.Add(runningPayload, XDayOneMinutePayload(runningPayload))
                            End If
                        Next
                        'Add all these payloads into the stock collections
                        If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                            If _SignalTimeFrame > 1 Then
                                XDayXMinutePayload = Cmn.ConvertPayloadsToXMinutes(XDayOneMinutePayload, _SignalTimeFrame)
                            Else
                                XDayXMinutePayload = XDayOneMinutePayload
                            End If
                            If _UseHeikenAshi Then
                                Indicator.HeikenAshi.ConvertToHeikenAshi(XDayXMinutePayload, XDayXMinuteHAPayload)
                            Else
                                XDayXMinuteHAPayload = XDayXMinutePayload
                            End If

                            If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                            If XDayXMinuteStocksPayload Is Nothing Then XDayXMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            XDayXMinuteStocksPayload.Add(stock, XDayXMinutePayload)
                            If XDayXMinuteHAStocksPayload Is Nothing Then XDayXMinuteHAStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            XDayXMinuteHAStocksPayload.Add(stock, XDayXMinuteHAPayload)
                        End If
                    End If
                Next

                If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count = 2 Then
                    Dim signalPayload As Dictionary(Of Date, Payload) = CloseGap.CalculateRule(XDayXMinuteHAStocksPayload.FirstOrDefault.Value, XDayXMinuteHAStocksPayload.LastOrDefault.Value)
                    If signalPayload IsNot Nothing AndAlso signalPayload.Count > 0 Then
                        Dim XDayRuleOutputPayload As Dictionary(Of String, Object) = Nothing
                        Using strategyBaseRule As New HazeOutsideBollinger(signalPayload, TickSize, stockList.FirstOrDefault.Value(0), Canceller)
                            strategyBaseRule.CalculateRule(XDayRuleOutputPayload)
                        End Using
                        If XDayRuleOutputPayload IsNot Nothing AndAlso XDayRuleOutputPayload.Count > 0 Then
                            Dim ruleSignalPayload As Dictionary(Of Date, Integer) = Nothing
                            If XDayRuleOutputPayload.ContainsKey("Signal") Then ruleSignalPayload = CType(XDayRuleOutputPayload("Signal"), Dictionary(Of Date, Integer))
                            Dim count As Integer = 0
                            For Each stock In stockList.Keys
                                count += 1
                                Dim XDayRuleSignalPayload As Dictionary(Of Date, Integer) = Nothing
                                Dim XDayRuleEntryPayload As Dictionary(Of Date, Decimal) = Nothing
                                Dim XDayRuleTargetPayload As Dictionary(Of Date, Decimal) = Nothing
                                Dim XDayRuleStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
                                Dim XDayRuleQuantityPayload As Dictionary(Of Date, Integer) = Nothing

                                Dim lastExistingPayload As Payload = Nothing
                                For Each runningPayload In ruleSignalPayload.Keys
                                    Dim signal As Integer = 0
                                    Dim entryPrice As Decimal = 0
                                    Dim slPrice As Decimal = 0
                                    Dim targetPrice As Decimal = 0
                                    Dim quantity As Integer = stockList(stock)(0)
                                    Dim currentPayload As Payload = Nothing
                                    If XDayXMinuteHAStocksPayload(stock).ContainsKey(runningPayload) Then
                                        currentPayload = XDayXMinuteHAStocksPayload(stock)(runningPayload)
                                    Else
                                        currentPayload = lastExistingPayload
                                    End If
                                    lastExistingPayload = currentPayload

                                    If count = 1 Then
                                        If signalPayload(runningPayload).Close > 0 Then
                                            signal = ruleSignalPayload(runningPayload) * -1
                                            entryPrice = If(signal = 1, 0, If(signal = -1, 10000000, 0))
                                            slPrice = If(signal = 1, currentPayload.Close - 1000, If(signal = -1, currentPayload.Close + 1000, 0))
                                            targetPrice = If(signal = 1, currentPayload.Close + 1000, If(signal = -1, currentPayload.Close - 1000, 0))
                                            'Else
                                            '    signal = ruleSignalPayload(runningPayload)
                                            '    entryPrice = If(signal = 1, 0, If(signal = -1, 10000000, 0))
                                            '    slPrice = If(signal = 1, currentPayload.Close - 1000, If(signal = -1, currentPayload.Close + 1000, 0))
                                            '    targetPrice = If(signal = 1, currentPayload.Close + 1000, If(signal = -1, currentPayload.Close - 1000, 0))
                                        End If
                                    ElseIf count = 2 Then
                                        If signalPayload(runningPayload).Close > 0 Then
                                            signal = ruleSignalPayload(runningPayload)
                                            entryPrice = If(signal = 1, 0, If(signal = -1, 10000000, 0))
                                            slPrice = If(signal = 1, currentPayload.Close - 1000, If(signal = -1, currentPayload.Close + 1000, 0))
                                            targetPrice = If(signal = 1, currentPayload.Close + 1000, If(signal = -1, currentPayload.Close - 1000, 0))
                                            'Else
                                            '    signal = ruleSignalPayload(runningPayload) * -1
                                            '    entryPrice = If(signal = 1, 0, If(signal = -1, 10000000, 0))
                                            '    slPrice = If(signal = 1, currentPayload.Close - 1000, If(signal = -1, currentPayload.Close + 1000, 0))
                                            '    targetPrice = If(signal = 1, currentPayload.Close + 1000, If(signal = -1, currentPayload.Close - 1000, 0))
                                        End If
                                    End If
                                    If XDayRuleSignalPayload Is Nothing Then XDayRuleSignalPayload = New Dictionary(Of Date, Integer)
                                    XDayRuleSignalPayload.Add(runningPayload, signal)
                                    If XDayRuleEntryPayload Is Nothing Then XDayRuleEntryPayload = New Dictionary(Of Date, Decimal)
                                    XDayRuleEntryPayload.Add(runningPayload, entryPrice)
                                    If XDayRuleStoplossPayload Is Nothing Then XDayRuleStoplossPayload = New Dictionary(Of Date, Decimal)
                                    XDayRuleStoplossPayload.Add(runningPayload, slPrice)
                                    If XDayRuleTargetPayload Is Nothing Then XDayRuleTargetPayload = New Dictionary(Of Date, Decimal)
                                    XDayRuleTargetPayload.Add(runningPayload, targetPrice)
                                    If XDayRuleQuantityPayload Is Nothing Then XDayRuleQuantityPayload = New Dictionary(Of Date, Integer)
                                    XDayRuleQuantityPayload.Add(runningPayload, quantity)
                                Next

                                If XDayRuleSignalStocksPayload Is Nothing Then XDayRuleSignalStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Integer))
                                XDayRuleSignalStocksPayload.Add(stock, XDayRuleSignalPayload)
                                If XDayRuleEntryStocksPayload Is Nothing Then XDayRuleEntryStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                XDayRuleEntryStocksPayload.Add(stock, XDayRuleEntryPayload)
                                If XDayRuleTargetStocksPayload Is Nothing Then XDayRuleTargetStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                XDayRuleTargetStocksPayload.Add(stock, XDayRuleTargetPayload)
                                If XDayRuleStoplossStocksPayload Is Nothing Then XDayRuleStoplossStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                XDayRuleStoplossStocksPayload.Add(stock, XDayRuleStoplossPayload)
                                If XDayRuleQuantityStocksPayload Is Nothing Then XDayRuleQuantityStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Integer))
                                XDayRuleQuantityStocksPayload.Add(stock, XDayRuleQuantityPayload)
                            Next
                        End If
                    End If
                End If
                '------------------------------------------------------------------------------------------------------------------------------------------------

                If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count = 2 Then
                    OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                    Dim startMinute As TimeSpan = ExchangeStartTime
                    Dim endMinute As TimeSpan = ExchangeEndTime
                    While startMinute < endMinute
                        Dim startSecond As TimeSpan = startMinute
                        Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromMinutes(_SignalTimeFrame - 1))
                        endSecond = endSecond.Add(TimeSpan.FromSeconds(59))
                        Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                        Dim potentialTickSignalTime As Date = Nothing
                        Dim currentMinuteCandlePayload As Payload = Nothing
                        Dim signalCandleTime As Date = Nothing

                        While startSecond <= endSecond
                            potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                            If potentialTickSignalTime.Second = 0 Then
                                potentialCandleSignalTime = potentialTickSignalTime
                            End If
                            For Each stockName In stockList.Keys
                                Dim runningTrade As Trade = Nothing
                                Dim runningTrade2 As Trade = Nothing
                                'Get the current minute candle from the stock collection for this stock for that day
                                Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                    currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialCandleSignalTime)
                                End If
                                signalCandleTime = potentialTickSignalTime.AddMinutes(-_SignalTimeFrame)
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
                                Dim squareOffValue As Double = Nothing
                                Dim quantity As Integer = Nothing
                                Dim lotSize As Integer = stockList(stockName)(1)
                                Dim entryBuffer As Decimal = Nothing
                                Dim stoplossBuffer As Decimal = Nothing
                                Dim supporting1 As String = If(XDayRuleSupporting1StocksPayload IsNot Nothing AndAlso XDayRuleSupporting1StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting1StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting1StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting1StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting2 As String = If(XDayRuleSupporting2StocksPayload IsNot Nothing AndAlso XDayRuleSupporting2StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting2StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting2StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting2StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting3 As String = If(XDayRuleSupporting3StocksPayload IsNot Nothing AndAlso XDayRuleSupporting3StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting3StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting3StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting3StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting4 As String = If(XDayRuleSupporting4StocksPayload IsNot Nothing AndAlso XDayRuleSupporting4StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting4StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting4StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting4StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting5 As String = If(XDayRuleSupporting5StocksPayload IsNot Nothing AndAlso XDayRuleSupporting5StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting5StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting5StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting5StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting6 As String = If(XDayRuleSupporting6StocksPayload IsNot Nothing AndAlso XDayRuleSupporting6StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting6StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting6StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting6StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting7 As String = If(XDayRuleSupporting7StocksPayload IsNot Nothing AndAlso XDayRuleSupporting7StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting7StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting7StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting7StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting8 As String = If(XDayRuleSupporting8StocksPayload IsNot Nothing AndAlso XDayRuleSupporting8StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting8StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting8StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting8StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting9 As String = If(XDayRuleSupporting9StocksPayload IsNot Nothing AndAlso XDayRuleSupporting9StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting9StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting9StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting9StocksPayload(stockName)(signalCandleTime), "")
                                Dim supporting10 As String = If(XDayRuleSupporting10StocksPayload IsNot Nothing AndAlso XDayRuleSupporting10StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting10StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting10StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting10StocksPayload(stockName)(signalCandleTime), "")

                                Dim tradeActive As Boolean = False

                                If currentMinuteCandlePayload IsNot Nothing AndAlso NumberOfTradesPerDay(currentMinuteCandlePayload.PayloadDate) < NumberOfTradePerDay Then
                                    If currentMinuteCandlePayload IsNot Nothing AndAlso NumberOfTradesPerStockPerDay(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol) < NumberOfTradePerStockPerDay AndAlso
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

                                            finalStoplossPrice = XDayRuleStoplossStocksPayload(stockName)(signalCandleTime)
                                            stoplossBuffer = CalculateBuffer(finalStoplossPrice, RoundOfType.Floor)
                                            finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalEntryPrice - finalStoplossPrice, 4))

                                            If TrailingSL Then
                                                finalTargetPrice = finalEntryPrice + 100
                                            Else
                                                finalTargetPrice = XDayRuleTargetStocksPayload(stockName)(signalCandleTime)
                                            End If
                                            finalTargetRemark = String.Format("Target: {0}", Math.Round(finalTargetPrice - finalEntryPrice, 4))

                                            squareOffValue = XDayRuleTargetStocksPayload(stockName)(signalCandleTime) - finalEntryPrice
                                            quantity = XDayRuleQuantityStocksPayload(stockName)(signalCandleTime)

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
                                                     currentMinuteCandlePayload)

                                                runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)

                                                finalTargetPrice = finalEntryPrice + 100
                                                finalTargetRemark = String.Format("Target: 100")

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
                                                     currentMinuteCandlePayload)

                                                runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
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
                                                     currentMinuteCandlePayload)

                                                runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
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
                                    ElseIf currentMinuteCandlePayload IsNot Nothing AndAlso NumberOfTradesPerStockPerDay(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol) < NumberOfTradePerStockPerDay AndAlso
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

                                            finalStoplossPrice = XDayRuleStoplossStocksPayload(stockName)(signalCandleTime)
                                            stoplossBuffer = CalculateBuffer(finalStoplossPrice, RoundOfType.Floor)
                                            finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalStoplossPrice - finalEntryPrice, 4))

                                            If TrailingSL Then
                                                finalTargetPrice = finalEntryPrice - 100
                                            Else
                                                finalTargetPrice = XDayRuleTargetStocksPayload(stockName)(signalCandleTime)
                                            End If
                                            finalTargetRemark = String.Format("Target: {0}", Math.Round(finalEntryPrice - finalTargetPrice, 4))

                                            squareOffValue = finalEntryPrice - XDayRuleTargetStocksPayload(stockName)(signalCandleTime)
                                            quantity = XDayRuleQuantityStocksPayload(stockName)(signalCandleTime)

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
                                                        currentMinuteCandlePayload)

                                                runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)

                                                finalTargetPrice = finalEntryPrice - 100
                                                finalTargetRemark = String.Format("Target: 100")

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
                                                       currentMinuteCandlePayload)

                                                runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
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
                                                        currentMinuteCandlePayload)

                                                runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
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
                                Else
                                    Dim extraCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                    If extraCancelTrades IsNot Nothing AndAlso extraCancelTrades.Count > 0 Then
                                        For Each extraCancelTrade In extraCancelTrades
                                            Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                            dummyPayload.PayloadDate = potentialTickSignalTime
                                            CancelTrade(extraCancelTrade, dummyPayload, String.Format("{0} Trade already executed", NumberOfTradePerDay))
                                        Next
                                    End If
                                End If
                                If currentSecondTickPayload IsNot Nothing AndAlso currentSecondTickPayload.Count > 0 Then
                                    For Each tick In currentSecondTickPayload
                                        SetCurrentLTPForStock(currentMinuteCandlePayload, tick, Trade.TradeType.MIS)

                                        'Modify Target Stoploss
                                        If ModifyTargetStoploss Then
                                            Dim potentialModifyTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                            If potentialModifyTrades IsNot Nothing AndAlso potentialModifyTrades.Count > 0 Then
                                                For Each potentialModifyTrade In potentialModifyTrades
                                                    If potentialModifyTrade.CoreTradingSymbol = stockName Then
                                                        Dim signalCheckTime As Date = GetCurrentXMinuteCandleTime(potentialCandleSignalTime, XDayXMinuteStocksPayload(stockName))
                                                        Dim modifiedTarget As Decimal = 0
                                                        Dim modifiedStoploss As Decimal = 0
                                                        Dim modifiedTargetRemark As String = Nothing
                                                        Dim modifiedStoplossRemark As String = Nothing
                                                        If XDayRuleTargetStocksPayload.ContainsKey(stockName) AndAlso
                                                        XDayRuleTargetStocksPayload(stockName).ContainsKey(signalCheckTime) AndAlso
                                                        XDayRuleTargetStocksPayload(stockName)(signalCheckTime) <> 0 Then
                                                            modifiedTarget = XDayRuleTargetStocksPayload(stockName)(signalCheckTime)
                                                            modifiedTargetRemark = String.Format("Modified Target: {0}", If(potentialModifyTrade.EntryDirection = Trade.TradeExecutionDirection.Buy, Math.Round(modifiedTarget - potentialModifyTrade.EntryPrice, 4), Math.Round(potentialModifyTrade.EntryPrice - modifiedTarget, 4)))
                                                            potentialModifyTrade.UpdateTrade(PotentialTarget:=modifiedTarget, TargetRemark:=modifiedTargetRemark)
                                                        End If
                                                        If XDayRuleStoplossStocksPayload.ContainsKey(stockName) AndAlso
                                                        XDayRuleStoplossStocksPayload(stockName).ContainsKey(signalCheckTime) AndAlso
                                                        XDayRuleStoplossStocksPayload(stockName)(signalCheckTime) <> 0 Then
                                                            modifiedStoploss = XDayRuleStoplossStocksPayload(stockName)(signalCheckTime)
                                                            modifiedStoplossRemark = String.Format("Modified Stoploss: {0}", If(potentialModifyTrade.EntryDirection = Trade.TradeExecutionDirection.Buy, Math.Round(potentialModifyTrade.EntryPrice - modifiedStoploss, 4), Math.Round(modifiedStoploss - potentialModifyTrade.EntryPrice, 4)))
                                                            potentialModifyTrade.UpdateTrade(PotentialStopLoss:=modifiedStoploss, SLRemark:=modifiedStoplossRemark)
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If

                                        'Stock Point Check
                                        'If StockPLPoint(tradeCheckingDate, tick.TradingSymbol) >= 0.05 OrElse
                                        '        StockPLPoint(tradeCheckingDate, tick.TradingSymbol) <= -0.1 Then
                                        '    ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Max Stock PL Point reached for the day")
                                        'End If

                                        'Stock MTM Check
                                        If ExitOnStockFixedTargetStoploss Then
                                            If StockPLAfterBrokerage(tradeCheckingDate, tick.TradingSymbol) >= StockMaxProfitPerDay OrElse
                                            StockPLAfterBrokerage(tradeCheckingDate, tick.TradingSymbol) <= StockMaxLossPerDay Then
                                                ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Max Stock PL reached for the day")
                                            End If
                                        End If

                                        'OverAll MTM Check
                                        If ExitOnOverAllFixedTargetStoploss Then
                                            If AllPLAfterBrokerage(tradeCheckingDate) >= OverAllProfitPerDay OrElse
                                            AllPLAfterBrokerage(tradeCheckingDate) <= OverAllLossPerDay Then
                                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, "Max PL reached for the day")
                                            End If
                                        End If

                                        'Exit Trade
                                        Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                            For Each potentialExitTrade In potentialExitTrades
                                                If ExitTradeIfPossible(potentialExitTrade, tick) Then
                                                    Console.WriteLine("")
                                                End If
                                            Next
                                        End If

                                        'Trailing SL
                                        If TrailingSL Then
                                            Dim potentialSLMoveTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                            If potentialSLMoveTrades IsNot Nothing AndAlso potentialSLMoveTrades.Count > 0 Then
                                                For Each potentialSLMoveTrade In potentialSLMoveTrades
                                                    Dim slMoveRemark As String = Nothing
                                                    Dim slPrice As Double = CalculateMathematicalTrailingSL(potentialSLMoveTrade, tick.Open, 0.005, slMoveRemark)
                                                    If Math.Round(potentialSLMoveTrade.EntryPrice, 4) <> Math.Round(slPrice, 4) Then
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
                                                If EnterTradeIfPossible(potentialEntryTrade, tick, Nothing, True) Then
                                                    Console.WriteLine("")
                                                    'If NumberOfTradesPerDay(currentMinuteCandlePayload.PayloadDate) = 1 Then
                                                    '    Dim potentialPL As Decimal = Decimal.MinValue
                                                    '    If potentialEntryTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                    '        potentialPL = Strategy.CalculatePL(potentialEntryTrade.CoreTradingSymbol, potentialEntryTrade.EntryPrice, potentialEntryTrade.EntryPrice + 0.05, potentialEntryTrade.Quantity, potentialEntryTrade.LotSize, potentialEntryTrade.StockType)
                                                    '    ElseIf potentialEntryTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                    '        potentialPL = Strategy.CalculatePL(potentialEntryTrade.CoreTradingSymbol, potentialEntryTrade.EntryPrice - 0.05, potentialEntryTrade.EntryPrice, potentialEntryTrade.Quantity, potentialEntryTrade.LotSize, potentialEntryTrade.StockType)
                                                    '    End If
                                                    '    If potentialPL <> Decimal.MinValue AndAlso potentialPL > 0 Then
                                                    '        Me.StockMaxProfitPerDay = potentialPL
                                                    '        Me.StockMaxLossPerDay = -1 * potentialPL * 2
                                                    '        Me.OverAllProfitPerDay = potentialPL
                                                    '        Me.OverAllLossPerDay = -1 * potentialPL * 2
                                                    '    End If
                                                    'End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                                'Cancel trade
                                If startSecond.Hours = endSecond.Hours AndAlso startSecond.Minutes = endSecond.Minutes AndAlso startSecond.Seconds = endSecond.Seconds Then
                                    Dim potentialCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                    If potentialCancelTrades IsNot Nothing AndAlso potentialCancelTrades.Count > 0 Then
                                        For Each potentialCancelTrade In potentialCancelTrades
                                            If potentialCancelTrade.CoreTradingSymbol = stockName Then
                                                Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                dummyPayload.PayloadDate = potentialTickSignalTime
                                                Dim signalCheckTime As Date = GetCurrentXMinuteCandleTime(potentialCandleSignalTime, XDayXMinuteStocksPayload(stockName))
                                                If XDayRuleSignalStocksPayload.ContainsKey(stockName) AndAlso
                                                XDayRuleSignalStocksPayload(stockName).ContainsKey(signalCheckTime) AndAlso
                                                XDayRuleSignalStocksPayload(stockName)(signalCheckTime) = 0 Then
                                                    CancelTrade(potentialCancelTrade, dummyPayload, String.Format("Invalid Signal"))
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                            startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                        End While   'Second Loop
                        'Force exit at day end
                        If startMinute = EODExitTime Then
                            ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, "EOD Force Exit")
                        End If
                        startMinute = startMinute.Add(TimeSpan.FromMinutes(_SignalTimeFrame))
                    End While   'Minute Loop
                End If
            End If
            totalPL += AllPLAfterBrokerage(tradeCheckingDate)
            tradeCheckingDate = tradeCheckingDate.AddDays(1)
        End While   'Date Loop

        'Excel Printing
        Dim filename As String = String.Format("Strategy Output PL {8},Partial-{3},NmbrTrd-{4},Trailing-{5},RvsTrd-{6},FxdMTM-{7},MaxP-{9},MaxL-{10} {0}-{1}-{2}.xlsx",
                                               Now.Hour, Now.Minute, Now.Second,
                                               If(PartialExit, "T", "F"),
                                               If(NumberOfTradePerStockPerDay = Integer.MaxValue, "NoLimit", NumberOfTradePerStockPerDay),
                                               If(TrailingSL, "T", "F"),
                                               If(ReverseSignalTrade, "T", "F"),
                                               If(ExitOnStockFixedTargetStoploss, "T", "F"),
                                               Math.Round(totalPL, 0),
                                               StockMaxProfitPerDay,
                                               Math.Abs(StockMaxLossPerDay))

        PrintArrayToExcel(filename)
    End Function

    'Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, Decimal())
    '    AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
    '    Dim dt As DataTable = Nothing
    '    Dim conn As MySqlConnection = Cmn.OpenDBConnection
    '    Dim ret As Dictionary(Of String, Decimal()) = Nothing

    '    Dim signalCheckingDate As Date = Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, tradingDate.Date)
    '    If conn.State = ConnectionState.Open Then
    '        OnHeartbeat("Fetching All Stock Data")
    '        Dim cmd As New MySqlCommand("GET_STOCK_CASH_DATA_ATR_VOLUME_ALL_DATES", conn)
    '        cmd.CommandType = CommandType.StoredProcedure
    '        cmd.Parameters.AddWithValue("@startDate", signalCheckingDate)
    '        cmd.Parameters.AddWithValue("@endDate", signalCheckingDate)
    '        cmd.Parameters.AddWithValue("@numberOfRecords", 5)
    '        cmd.Parameters.AddWithValue("@minClose", 100)
    '        cmd.Parameters.AddWithValue("@maxClose", 1500)
    '        cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
    '        cmd.Parameters.AddWithValue("@potentialAmount", 500000)

    '        Dim adapter As New MySqlDataAdapter(cmd)
    '        adapter.SelectCommand.CommandTimeout = 3000
    '        dt = New DataTable
    '        adapter.Fill(dt)
    '    End If
    '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '        For i = 0 To dt.Rows.Count - 1
    '            If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
    '            ret.Add(dt.Rows(i).Item(1), {dt.Rows(i).Item(5), dt.Rows(i).Item(5)})
    '        Next
    '    End If
    '    Return ret
    'End Function

    'Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, Decimal())
    '    AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
    '    Dim dt As DataTable = Nothing
    '    Dim conn As MySqlConnection = Cmn.OpenDBConnection
    '    Dim ret As Dictionary(Of String, Decimal()) = Nothing

    '    If conn.State = ConnectionState.Open Then
    '        OnHeartbeat("Fetching Top Gainer Looser Data")
    '        Dim cmd As New MySqlCommand("GET_DAY_TOP_GAINER_LOOSER_DATA_ATR_VOLUME_ALL_DATES", conn)
    '        cmd.CommandType = CommandType.StoredProcedure
    '        cmd.Parameters.AddWithValue("@startDate", tradingDate)
    '        cmd.Parameters.AddWithValue("@endDate", tradingDate)
    '        cmd.Parameters.AddWithValue("@userTime", "09:19:00")
    '        cmd.Parameters.AddWithValue("@numberOfRecords", 0)
    '        cmd.Parameters.AddWithValue("@minClose", 100)
    '        cmd.Parameters.AddWithValue("@maxClose", 1500)
    '        cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
    '        cmd.Parameters.AddWithValue("@potentialAmount", 1000000)

    '        Dim adapter As New MySqlDataAdapter(cmd)
    '        adapter.SelectCommand.CommandTimeout = 3000
    '        dt = New DataTable
    '        adapter.Fill(dt)
    '    End If
    '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '        For i = 0 To dt.Rows.Count - 1
    '            'If dt.Rows(i).Item(11) >= 0.5 Then
    '            If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
    '            ret.Add(dt.Rows(i).Item(1), {dt.Rows(i).Item(5), 1, dt.Rows(i).Item(3)})
    '            'End If
    '            If i = 19 Then Exit For
    '        Next
    '    End If
    '    Return ret
    'End Function

    Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, Decimal())
        AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
        Dim dt As DataTable = Nothing
        Dim conn As MySqlConnection = Cmn.OpenDBConnection
        Dim ret As Dictionary(Of String, Decimal()) = Nothing

        If conn.State = ConnectionState.Open Then
            OnHeartbeat("Fetching Pre Market Data")
            Dim cmd As New MySqlCommand("GET_PRE_MARKET_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", tradingDate)
            cmd.Parameters.AddWithValue("@endDate", tradingDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@minClose", 500)
            cmd.Parameters.AddWithValue("@maxClose", 1000)
            cmd.Parameters.AddWithValue("@atrPercentage", 0)
            cmd.Parameters.AddWithValue("@potentialAmount", 0)
            cmd.Parameters.AddWithValue("@sortColumn", "ChangePer")

            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 3000
            dt = New DataTable
            adapter.Fill(dt)
        End If
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                If Me.NIFTY50Stocks.Contains(dt.Rows(i).Item(1)) Then
                    If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
                    ret.Add(dt.Rows(i).Item(1), {100, 100, dt.Rows(i).Item(5)})
                End If
            Next
        End If
        Return ret
    End Function

    'Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, Decimal())
    '    AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
    '    Dim dts As DataSet = Nothing
    '    Dim conn As MySqlConnection = Cmn.OpenDBConnection
    '    Dim ret As Dictionary(Of String, Decimal()) = Nothing

    '    If conn.State = ConnectionState.Open Then
    '        OnHeartbeat(String.Format("Fetching Pre Market Data for {0}", tradingDate.ToShortDateString))
    '        Dim cmd As New MySqlCommand("GET_PRE_MARKET_DATA_ATR_VOLUME_ALL_DATES", conn)
    '        cmd.CommandType = CommandType.StoredProcedure
    '        cmd.Parameters.AddWithValue("@startDate", tradingDate)
    '        cmd.Parameters.AddWithValue("@endDate", tradingDate)
    '        cmd.Parameters.AddWithValue("@numberOfRecords", 0)
    '        cmd.Parameters.AddWithValue("@minClose", 100)
    '        cmd.Parameters.AddWithValue("@maxClose", 1500)
    '        cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
    '        cmd.Parameters.AddWithValue("@potentialAmount", 1000000)
    '        cmd.Parameters.AddWithValue("@sortColumn", "ChangePer")

    '        Dim adapter As New MySqlDataAdapter(cmd)
    '        adapter.SelectCommand.CommandTimeout = 3000
    '        dts = New DataSet
    '        adapter.Fill(dts)
    '    End If
    '    If dts IsNot Nothing AndAlso dts.Tables.Count > 0 Then
    '        Dim totalTables As Integer = dts.Tables.Count
    '        Dim dt As DataTable = Nothing
    '        Dim count As Integer = 0
    '        While Not count > totalTables - 1
    '            Dim temp_dt As New DataTable
    '            temp_dt = dts.Tables(count)
    '            If temp_dt.Rows.Count > 0 Then
    '                Dim tempCol As DataColumn = temp_dt.Columns.Add("Abs", GetType(Double), "Iif(ChangePer <0 ,ChangePer*(-1),ChangePer)")
    '                temp_dt = temp_dt.Select("Abs>=0", "Abs DESC").Cast(Of DataRow).Take(Integer.MaxValue).CopyToDataTable
    '            End If
    '            If dt Is Nothing Then dt = New DataTable
    '            dt.Merge(temp_dt)
    '            count += 1
    '        End While
    '        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '            For i = 0 To dt.Rows.Count - 1
    '                Dim stockName As String = dt.Rows(i).Item(1)
    '                Dim currentDayPayload As Dictionary(Of Date, Payload) = Cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, stockName, tradingDate, tradingDate)
    '                Dim previousTradingDate As Date = Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, stockName, tradingDate)
    '                Dim previousDayPayload As Dictionary(Of Date, Payload) = Cmn.GetRawPayload(Common.DataBaseTable.EOD_Cash, stockName, previousTradingDate, previousTradingDate)
    '                Dim signalDirection As Integer = If(dt.Rows(i).Item(5) >= 0, -1, 1)
    '                If signalDirection = 1 Then
    '                    If currentDayPayload.FirstOrDefault.Value.Open > previousDayPayload.FirstOrDefault.Value.Low Then
    '                        If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
    '                        ret.Add(stockName, {dt.Rows(i).Item(11), signalDirection, previousDayPayload.FirstOrDefault.Value.Close, dt.Rows(i).Item(8)})
    '                    End If
    '                ElseIf signalDirection = -1 Then
    '                    If currentDayPayload.FirstOrDefault.Value.Open < previousDayPayload.FirstOrDefault.Value.High Then
    '                        If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
    '                        ret.Add(stockName, {dt.Rows(i).Item(11), signalDirection, previousDayPayload.FirstOrDefault.Value.Close, dt.Rows(i).Item(8)})
    '                    End If
    '                End If
    '            Next
    '        End If
    '    End If
    '    Return ret
    'End Function

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