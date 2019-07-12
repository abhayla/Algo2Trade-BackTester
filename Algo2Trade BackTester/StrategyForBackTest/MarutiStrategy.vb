Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading
Public Class MarutiStrategy
    Inherits Strategy
    Implements IDisposable

    Private ReadOnly _SignalTimeFrame As Integer = 30
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan)
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
    End Sub
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tradeStockType As Trade.TypeOfStock = Trade.TypeOfStock.Cash
        Dim databaseTable As Common.DataBaseTable = Common.DataBaseTable.Intraday_Cash
        Dim totalPL As Decimal = 0
        Dim tradeCheckingDate As Date = startDate
        While tradeCheckingDate <= endDate
            Dim stockList As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
            stockList.Add("MARUTI", 75)

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim XDayXMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing

                'First lets build the payload for all the stocks
                For Each stock In stockList.Keys
                    Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim XDayXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing

                    'Get payload
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
                            XDayXMinutePayload = Cmn.ConvertPayloadsToXMinutes(XDayOneMinutePayload, _SignalTimeFrame)

                            If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                            If XDayXMinuteStocksPayload Is Nothing Then XDayXMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            XDayXMinuteStocksPayload.Add(stock, XDayXMinutePayload)
                        End If
                    End If
                Next

                '------------------------------------------------------------------------------------------------------------------------------------------------

                If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                    OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                    Dim startMinute As TimeSpan = ExchangeStartTime
                    Dim endMinute As TimeSpan = ExchangeEndTime
                    Dim oncePerDay As Boolean = True
                    While startMinute < endMinute
                        Dim startSecond As TimeSpan = startMinute
                        Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromMinutes(_SignalTimeFrame - 1))
                        endSecond = endSecond.Add(TimeSpan.FromSeconds(59))
                        Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                        Dim potentialTickSignalTime As Date = Nothing
                        Dim currentMinuteCandlePayload As Payload = Nothing
                        Dim signalCandlePayload As Payload = Nothing

                        While startSecond <= endSecond
                            potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                            If potentialTickSignalTime.Second = 0 Then
                                potentialCandleSignalTime = potentialTickSignalTime
                            End If
                            For Each stockName In stockList.Keys
                                Dim runningTrade1 As Trade = Nothing
                                Dim runningTrade2 As Trade = Nothing
                                'Get the current minute candle from the stock collection for this stock for that day
                                Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                    currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialCandleSignalTime)
                                End If
                                'Signal Candle
                                Dim signalCandleTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, 9, 45, 0)
                                If XDayXMinuteStocksPayload.ContainsKey(stockName) AndAlso XDayXMinuteStocksPayload(stockName).ContainsKey(signalCandleTime) Then
                                    signalCandlePayload = XDayXMinuteStocksPayload(stockName)(signalCandleTime)
                                End If
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
                                Dim lotSize As Integer = 1
                                Dim entryBuffer As Decimal = Nothing
                                Dim stoplossBuffer As Decimal = Nothing
                                If oncePerDay AndAlso
                                    currentMinuteCandlePayload IsNot Nothing AndAlso
                                    currentMinuteCandlePayload.PayloadDate.Hour = 10 AndAlso
                                    currentMinuteCandlePayload.PayloadDate.Minute = 15 Then
                                    oncePerDay = False

                                    'Buy Order
                                    finalEntryPrice = signalCandlePayload.High
                                    entryBuffer = 1
                                    finalEntryPrice += entryBuffer
                                    finalStoplossPrice = finalEntryPrice - 30
                                    stoplossBuffer = 0
                                    finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalEntryPrice - finalStoplossPrice, 2))
                                    finalTargetPrice = finalEntryPrice + 60
                                    finalTargetRemark = String.Format("Target: {0}", Math.Round(finalTargetPrice - finalEntryPrice, 2))
                                    quantity = stockList(stockName) * lotSize

                                    runningTrade1 = New Trade(Me,
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
                                                         signalCandlePayload)

                                    'Sell Order
                                    finalEntryPrice = signalCandlePayload.Low
                                    entryBuffer = 1
                                    finalEntryPrice -= entryBuffer
                                    finalStoplossPrice = finalEntryPrice + 30
                                    stoplossBuffer = 0
                                    finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalStoplossPrice - finalEntryPrice, 2))
                                    finalTargetPrice = finalEntryPrice - 60
                                    finalTargetRemark = String.Format("Target: {0}", Math.Round(finalEntryPrice - finalTargetPrice, 2))
                                    quantity = stockList(stockName) * lotSize

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
                                                         quantity,
                                                         lotSize,
                                                         finalTargetPrice,
                                                         finalTargetRemark,
                                                         finalStoplossPrice,
                                                         stoplossBuffer,
                                                         finalStoplossRemark,
                                                         signalCandlePayload)

                                    If runningTrade1 IsNot Nothing Then PlaceOrModifyOrder(runningTrade1, Nothing)
                                    If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                End If

                                If currentSecondTickPayload IsNot Nothing AndAlso currentSecondTickPayload.Count > 0 Then
                                    For Each tick In currentSecondTickPayload
                                        SetCurrentLTPForStock(currentMinuteCandlePayload, tick, Trade.TradeType.MIS)
                                        'Exit Trade
                                        Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                        If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                            For Each potentialExitTrade In potentialExitTrades
                                                ExitTradeIfPossible(potentialExitTrade, tick, False)
                                            Next
                                        End If

                                        'Enter Trade
                                        Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                            For Each potentialEntryTrade In potentialEntryTrades
                                                Dim placeOrderResponse As Tuple(Of Boolean, Date) = EnterTradeIfPossible(potentialEntryTrade, tick, False)
                                                If placeOrderResponse IsNot Nothing AndAlso placeOrderResponse.Item1 Then
                                                    Console.WriteLine("")
                                                    Dim oppositeCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                                    If oppositeCancelTrades IsNot Nothing AndAlso oppositeCancelTrades.Count > 0 Then
                                                        For Each oppositeCancelTrade In oppositeCancelTrades
                                                            CancelTrade(oppositeCancelTrade, tick, String.Format("Opposite Trade Triggered"))
                                                        Next
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                                'Cancel trade
                                If startSecond.Hours = 15 AndAlso startSecond.Minutes = 10 AndAlso startSecond.Seconds = 0 Then
                                    Dim potentialCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                    If potentialCancelTrades IsNot Nothing AndAlso potentialCancelTrades.Count > 0 Then
                                        For Each potentialCancelTrade In potentialCancelTrades
                                            If potentialCancelTrade.CoreTradingSymbol = stockName Then
                                                Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                dummyPayload.PayloadDate = potentialTickSignalTime
                                                CancelTrade(potentialCancelTrade, dummyPayload, String.Format("Invalid Signal"))
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
        Dim filename As String = String.Format("Maruti Strategy {0}-{1}-{2}.xlsx",
                                               Now.Hour, Now.Minute, Now.Second)

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
