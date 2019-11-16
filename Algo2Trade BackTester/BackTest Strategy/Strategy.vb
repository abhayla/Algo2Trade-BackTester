Imports System.Threading
Imports Utilities.DAL
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public MustInherit Class Strategy
    Private _canceller As CancellationTokenSource
#Region "Events/Event handlers"
    Public Event DocumentDownloadComplete()
    Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
    Public Event Heartbeat(ByVal msg As String)
    Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
    'The below functions are needed to allow the derived classes to raise the above two events
    Protected Overridable Sub OnDocumentDownloadComplete()
        RaiseEvent DocumentDownloadComplete()
    End Sub
    Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        RaiseEvent DocumentRetryStatus(currentTry, totalTries)
    End Sub
    Protected Overridable Sub OnHeartbeat(ByVal msg As String)
        RaiseEvent Heartbeat(msg)
    End Sub
    Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
    End Sub
#End Region

#Region "Enum"
    Public Enum TradeExecutionDirection
        Buy = 1
        Sell
        None
    End Enum
    Public Enum TradeExecutionStatus
        Open = 1
        Inprogress
        Close
        Cancel
    End Enum
    Public Enum StopLossPoint
        Fixed = 1
        MinumumOfFixedAndHighLow
    End Enum
    Public Enum TypeOfStock
        Cash = 1
        Currency
        Commodity
        Futures
        None
    End Enum
    Public Enum TradeExitCondition
        Target = 1
        StopLoss
        EndOfDay
        Cancelled
        ForceExit
    End Enum
#End Region

#Region "Trade Class"
    Public Class Trade
        Public Property TradingSymbol As String
        Public Property TradingDate As Date
        Public Property TradingStatus As Strategy.TradeExecutionStatus
        Public Property EntryTime As DateTime
        Public Property EntryDirection As Strategy.TradeExecutionDirection
        Public Property EntryPrice As Double
        Public Property EntryType As String
        Public Property EntryCondition As String
        Public Property ExitTime As DateTime
        Public Property ExitPrice As Double
        Public Property ExitCondition As Strategy.TradeExitCondition
        Public Property StopLossRemark As String
        Public Property Quantity As Integer
        Public Property PotentialTP As Double
        Public Property PotentialSL As Double
        Public Property ActiveCapitalWithMargin As Double
        Public Property PassiveCapitalWithMargin As Double
        Public Property LeftOverCapitalBeforeTakingTrade As Double
        Public Property ProfitLoss As Double
        Public ReadOnly Property CapitalRequiredWithMargin As Double
            Get
                Return Me.EntryPrice * Me.Quantity / 30
            End Get
        End Property
        Private Property _MaximumDrawUp As Double = Double.MinValue
        Public Property MaximumDrawUp As Double
            Get
                If _MaximumDrawUp = Double.MinValue Then
                    _MaximumDrawUp = Me.EntryPrice
                End If
                Return _MaximumDrawUp
            End Get
            Set(value As Double)
                _MaximumDrawUp = value
            End Set
        End Property
        Private Property _MaximumDrawDown As Double = Double.MinValue
        Public Property MaximumDrawDown As Double
            Get
                If _MaximumDrawDown = Double.MinValue Then
                    _MaximumDrawDown = Me.EntryPrice
                End If
                Return _MaximumDrawDown
            End Get
            Set(value As Double)
                _MaximumDrawDown = value
            End Set
        End Property
        Public Property DrawDownPL As Double
        Public Property SignalCandle As Payload
        Public Property ATRPercentage As Double
        Public Property AbsoluteATR As Double
        Public Property IndicatorCandleTime As Date
        Public Property SignalCandleVolumeSMA As Double
        Public Property HeikenAshiSignalCandle As Payload
        Public Property HeikenAshiSignalCandleStatus As String
        Public Property FirstProfitableTradeOfDay As Boolean
        Public Property TypeOfStock As TypeOfStock

        Private _DurationOfTrade As TimeSpan
        Public ReadOnly Property DurationOfTrade As TimeSpan
            Get
                _DurationOfTrade = Me.ExitTime - Me.EntryTime
                Return _DurationOfTrade
            End Get
        End Property
        Private _PLPoint As Double
        Public ReadOnly Property PLPoint As Double
            Get
                If Me.EntryDirection = TradeExecutionDirection.Buy Then
                    _PLPoint = Me.ExitPrice - Me.EntryPrice
                ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                    _PLPoint = Me.EntryPrice - Me.ExitPrice
                End If
                Return Math.Round(_PLPoint, 2)
            End Get
        End Property
        Private _ProfitLossWithoutBrokerage As Double
        Public ReadOnly Property ProfitLossWithoutBrokerage As Double
            Get
                If Me.EntryDirection = TradeExecutionDirection.Buy Then
                    _ProfitLossWithoutBrokerage = Math.Round((Me.ExitPrice - Me.EntryPrice) * Me.Quantity, 2)
                ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                    _ProfitLossWithoutBrokerage = Math.Round((Me.EntryPrice - Me.ExitPrice) * Me.Quantity, 2)
                End If
                Return _ProfitLossWithoutBrokerage
            End Get
        End Property
    End Class
#End Region

#Region "Public Variables"
    'Public Property TradesTaken As List(Of Trade)
    Public Property StopLossSet As StopLossPoint
    Public Property StopLossMoveToBreakEven As Boolean
    Public Property WinRatio As Double
    Public Property CandleWickSizePercentage As Integer
    Public Property MaxStopLossPercentage As Integer
    Public Property TakeAllTrades As Boolean
    Public Property CandleRange As Boolean
    Public Property TradesTaken As Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
#End Region

#Region "Public Functions"
    Public Function CalculateBuffer(ByVal price As Double) As Double
        Dim bufferPrice As Double = Nothing
        If price < 200 Then
            bufferPrice = 0.05
        ElseIf price >= 200 And price <= 600 Then
            bufferPrice = 0.1
        ElseIf price > 600 And price <= 1200 Then
            bufferPrice = 0.15
        ElseIf price > 1200 And price <= 2000 Then
            bufferPrice = 0.2
        ElseIf price > 2000 And price <= 3000 Then
            bufferPrice = 0.25
        Else
            bufferPrice = 0.4
        End If
        Return bufferPrice
    End Function
    Public Function CalculateBuffer(ByVal price As Double, ByVal floorOrCeiling As RoundOfType) As Double
        Dim bufferPrice As Double = Nothing
        bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, 0.05, floorOrCeiling)
        Return bufferPrice
    End Function
    Public Function IsTradeActive(ByVal tradeDate As Date, ByVal tradingSymbol As String) As Boolean
        Dim ret As Boolean = False
        ret = TradesTaken IsNot Nothing AndAlso
                 TradesTaken.Count > 0 AndAlso
                TradesTaken.ContainsKey(tradeDate.Date) AndAlso
                TradesTaken(tradeDate.Date).ContainsKey(tradingSymbol) AndAlso
                TradesTaken(tradeDate.Date)(tradingSymbol).Find(Function(x)
                                                                    Return x.TradingStatus = TradeExecutionStatus.Inprogress
                                                                End Function) IsNot Nothing
        Return ret
    End Function
    Public Function GetTotalTrades(ByVal tradeDate As Date, ByVal tradingSymbol As String) As Integer
        Dim ret As Integer = 0
        If TradesTaken IsNot Nothing AndAlso
                 TradesTaken.Count > 0 AndAlso
                TradesTaken.ContainsKey(tradeDate.Date) AndAlso
                TradesTaken(tradeDate.Date).ContainsKey(tradingSymbol) Then
            Dim allTradesForThatDayForThatSymbol As List(Of Trade) = TradesTaken(tradeDate.Date)(tradingSymbol).FindAll(Function(x)
                                                                                                                            Return x.TradingStatus = TradeExecutionStatus.Close
                                                                                                                        End Function)
            If allTradesForThatDayForThatSymbol IsNot Nothing Then
                ret = allTradesForThatDayForThatSymbol.Count
            End If
        End If
        Return ret
    End Function
    Public Sub EnterOrder(ByVal tradeDate As Date, ByVal tradingSymbol As String, ByVal currentTrade As Trade)
        With currentTrade
            If .EntryDirection = TradeExecutionDirection.Buy Then
                .EntryPrice = ConvertFloorCeling(.EntryPrice, 0.05, RoundOfType.Celing)
                .PotentialTP = ConvertFloorCeling(.PotentialTP, 0.05, RoundOfType.Floor)
                .PotentialSL = ConvertFloorCeling(.PotentialSL, 0.05, RoundOfType.Floor)
            ElseIf .EntryDirection = TradeExecutionDirection.Sell Then
                .EntryPrice = ConvertFloorCeling(.EntryPrice, 0.05, RoundOfType.Floor)
                .PotentialTP = ConvertFloorCeling(.PotentialTP, 0.05, RoundOfType.Floor)
                .PotentialSL = ConvertFloorCeling(.PotentialSL, 0.05, RoundOfType.Celing)
            End If
        End With
        If TradesTaken.ContainsKey(tradeDate) Then
            If TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
                TradesTaken(tradeDate)(tradingSymbol).Add(currentTrade)
            Else
                TradesTaken(tradeDate).Add(tradingSymbol, New List(Of Trade) From {currentTrade})
            End If
        Else
            TradesTaken.Add(tradeDate, New Dictionary(Of String, List(Of Trade)) From {{tradingSymbol, New List(Of Trade) From {currentTrade}}})
        End If
    End Sub
    Public Function GetSpecificTrades(ByVal tradeDate As Date, ByVal tradingSymbol As String, ByVal tradeStatus As Strategy.TradeExecutionStatus) As List(Of Trade)
        Dim ret As List(Of Trade) = Nothing
        If TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
            ret = TradesTaken(tradeDate)(tradingSymbol).FindAll(Function(x)
                                                                    Return x.TradingStatus = tradeStatus
                                                                End Function)
        End If

        Return ret
    End Function
    Public Function GetLastTrade(ByVal tradeDate As Date, ByVal tradingSymbol As String, ByVal tradeStatus As Strategy.TradeExecutionStatus) As Trade
        Dim ret As Trade = Nothing
        If tradeStatus <> TradeExecutionStatus.Close And tradeStatus <> TradeExecutionStatus.Inprogress Then Throw New ApplicationException("Potential duplicate status could be retrieved, so will not proceed")
        If TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
            ret = TradesTaken(tradeDate)(tradingSymbol).FindAll(Function(x)
                                                                    Return x.TradingStatus = tradeStatus
                                                                End Function).LastOrDefault
        End If

        Return ret
    End Function
    Public Function GetEndOfDayTrade(ByVal tradeDate As Date, ByVal tradingSymbol As String, ByVal tradeStatus As Strategy.TradeExecutionStatus) As Trade
        Dim ret As Trade = Nothing
        If tradeStatus <> TradeExecutionStatus.Close Then Throw New ApplicationException("Potential duplicate status could be retrieved, so will not proceed")
        If TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
            ret = TradesTaken(tradeDate)(tradingSymbol).FindAll(Function(x)
                                                                    Return x.TradingStatus = tradeStatus AndAlso x.ExitCondition = TradeExitCondition.EndOfDay
                                                                End Function).LastOrDefault
        End If

        Return ret
    End Function
    Public Function EnterTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal typeOfStock As TypeOfStock) As Boolean
        Dim ret As Boolean = False
        Dim stockName As String = Nothing
        If currentPayload.TradingSymbol.Contains("FUT") Then
            stockName = currentPayload.TradingSymbol.Remove(currentTrade.TradingSymbol.Count - 8)
        Else
            stockName = currentPayload.TradingSymbol
        End If
        Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(currentPayload.PayloadDate.Date, stockName, TradeExecutionStatus.Inprogress)
        If currentTrade.EntryDirection = TradeExecutionDirection.Buy Then
            If currentPayload.High >= currentTrade.EntryPrice Then
                If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                    For Each item In itemSpecificTrade
                        MoveStopLoss(item, currentTrade.EntryPrice)
                        ExitTradeIfPossible(item, currentPayload, "00:00:00", typeOfStock, True)
                    Next
                End If
                With currentTrade
                    .TradingStatus = TradeExecutionStatus.Inprogress
                    .EntryTime = currentPayload.PayloadDate
                End With
                ret = True
            End If
        ElseIf currentTrade.EntryDirection = TradeExecutionDirection.Sell Then
            If currentPayload.Low <= currentTrade.EntryPrice Then
                If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                    For Each item In itemSpecificTrade
                        MoveStopLoss(item, currentTrade.EntryPrice)
                        ExitTradeIfPossible(item, currentPayload, "00:00:00", typeOfStock, True)
                    Next
                End If
                With currentTrade
                    .TradingStatus = TradeExecutionStatus.Inprogress
                    .EntryTime = currentPayload.PayloadDate
                End With
                ret = True
            End If
        End If
        Return ret
    End Function
    Public Function EnterTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal typeOfStock As TypeOfStock, ByRef tickCount As Integer) As Boolean
        Dim ret As Boolean = False
        Dim stockName As String = Nothing
        If currentPayload.TradingSymbol.Contains("FUT") Then
            stockName = currentPayload.TradingSymbol.Remove(currentTrade.TradingSymbol.Count - 8)
        Else
            stockName = currentPayload.TradingSymbol
        End If

        For i = tickCount To currentPayload.Ticks.Count - 1
            Dim itemSpecificTrade As List(Of Trade) = GetSpecificTrades(currentPayload.PayloadDate.Date, stockName, TradeExecutionStatus.Inprogress)
            If currentTrade.EntryDirection = TradeExecutionDirection.Buy Then
                If currentPayload.Ticks(i) >= currentTrade.EntryPrice Then
                    If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                        For Each item In itemSpecificTrade
                            MoveStopLoss(item, currentTrade.EntryPrice)
                            ExitTradeIfPossible(item, currentPayload, "00:00:00", typeOfStock, True, tickCount)
                        Next
                    End If
                    With currentTrade
                        .TradingStatus = TradeExecutionStatus.Inprogress
                        .EntryTime = currentPayload.PayloadDate
                    End With
                    ret = True
                    tickCount = i
                    Exit For
                End If
            ElseIf currentTrade.EntryDirection = TradeExecutionDirection.Sell Then
                If currentPayload.Ticks(i) <= currentTrade.EntryPrice Then
                    If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                        For Each item In itemSpecificTrade
                            MoveStopLoss(item, currentTrade.EntryPrice)
                            ExitTradeIfPossible(item, currentPayload, "00:00:00", typeOfStock, True, tickCount)
                        Next
                    End If
                    With currentTrade
                        .TradingStatus = TradeExecutionStatus.Inprogress
                        .EntryTime = currentPayload.PayloadDate
                    End With
                    ret = True
                    tickCount = i
                    Exit For
                End If
            End If
            tickCount = i
        Next
        Return ret
    End Function
    Public Sub CancelTrade(ByVal currentTrade As Trade, ByVal currentPayload As Payload)
        With currentTrade
            .TradingStatus = TradeExecutionStatus.Cancel
            .ExitTime = currentPayload.PayloadDate
            .ExitCondition = TradeExitCondition.Cancelled
        End With
    End Sub
    Public Function ExitTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal endOfDay As DateTime, ByVal typeOfStock As Strategy.TypeOfStock, ByVal forceExit As Boolean) As Boolean
        Dim ret As Boolean = False
        If Not forceExit Then
            If currentTrade.EntryDirection = TradeExecutionDirection.Buy Then
                If currentPayload.Low <= currentTrade.PotentialSL Then
                    Dim exitPrice As Double = currentTrade.PotentialSL
                    If currentPayload.Open <= currentTrade.PotentialSL Then
                        exitPrice = currentPayload.Open
                    End If
                    With currentTrade
                        .TradingStatus = TradeExecutionStatus.Close
                        .ExitTime = currentPayload.PayloadDate
                        .ExitPrice = exitPrice
                        .ExitCondition = TradeExitCondition.StopLoss
                    End With
                    ret = True
                ElseIf currentPayload.High >= currentTrade.PotentialTP Then
                    Dim exitPrice As Double = currentTrade.PotentialTP
                    If currentPayload.Open >= currentTrade.PotentialTP Then
                        exitPrice = currentPayload.Open
                    End If
                    With currentTrade
                        .TradingStatus = TradeExecutionStatus.Close
                        .ExitTime = currentPayload.PayloadDate
                        .ExitPrice = exitPrice
                        .ExitCondition = TradeExitCondition.Target
                    End With
                    ret = True
                End If
                If Not ret Then
                    currentTrade.MaximumDrawDown = Math.Min(currentPayload.Low, currentTrade.MaximumDrawDown)
                    currentTrade.MaximumDrawUp = Math.Max(currentPayload.High, currentTrade.MaximumDrawUp)
                End If
            ElseIf currentTrade.EntryDirection = TradeExecutionDirection.Sell Then
                If currentPayload.High >= currentTrade.PotentialSL Then
                    Dim exitPrice As Double = currentTrade.PotentialSL
                    If currentPayload.Open >= currentTrade.PotentialSL Then
                        exitPrice = currentPayload.Open
                    End If
                    With currentTrade
                        .TradingStatus = TradeExecutionStatus.Close
                        .ExitTime = currentPayload.PayloadDate
                        .ExitPrice = exitPrice
                        .ExitCondition = TradeExitCondition.StopLoss
                    End With
                    ret = True
                ElseIf currentPayload.Low <= currentTrade.PotentialTP Then
                    Dim exitPrice As Double = currentTrade.PotentialTP
                    If currentPayload.Open <= currentTrade.PotentialTP Then
                        exitPrice = currentPayload.Open
                    End If
                    With currentTrade
                        .TradingStatus = TradeExecutionStatus.Close
                        .ExitTime = currentPayload.PayloadDate
                        .ExitPrice = exitPrice
                        .ExitCondition = TradeExitCondition.Target
                    End With
                    ret = True
                End If
                If Not ret Then
                    currentTrade.MaximumDrawUp = Math.Min(currentPayload.Low, currentTrade.MaximumDrawUp)
                    currentTrade.MaximumDrawDown = Math.Max(currentPayload.High, currentTrade.MaximumDrawDown)
                End If
            End If
        Else
            With currentTrade
                .TradingStatus = TradeExecutionStatus.Close
                .ExitTime = currentPayload.PayloadDate
                .ExitPrice = currentTrade.PotentialSL
                .ExitCondition = TradeExitCondition.ForceExit
            End With
            ret = True
        End If

        If Not ret And currentPayload.PayloadDate.TimeOfDay.Hours = endOfDay.TimeOfDay.Hours And currentPayload.PayloadDate.TimeOfDay.Minutes = endOfDay.TimeOfDay.Minutes Then
            With currentTrade
                .TradingStatus = TradeExecutionStatus.Close
                .ExitTime = currentPayload.PayloadDate
                .ExitPrice = currentPayload.Open
                .ExitCondition = TradeExitCondition.EndOfDay
            End With
            ret = True
        End If
        If ret = True Then
            If currentTrade.EntryDirection = TradeExecutionDirection.Buy Then
                With currentTrade
                    .ProfitLoss = CalculateProfitLoss(.TradingSymbol, .EntryPrice, .ExitPrice, .Quantity, typeOfStock)
                End With
            ElseIf currentTrade.EntryDirection = TradeExecutionDirection.Sell Then
                With currentTrade
                    .ProfitLoss = CalculateProfitLoss(.TradingSymbol, .ExitPrice, .EntryPrice, .Quantity, typeOfStock)
                End With
            End If
        End If

        'GetFirstProfitableTradeOfDay(currentTrade.TradingDate, currentTrade.TradingSymbol, currentTrade)

        Return ret
    End Function
    Public Function ExitTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal endOfDay As DateTime, ByVal typeOfStock As Strategy.TypeOfStock, ByVal forceExit As Boolean, ByRef tickCount As Integer) As Boolean
        Dim ret As Boolean = False

        For i = tickCount To currentPayload.Ticks.Count - 1
            If Not forceExit Then
                If currentTrade.EntryDirection = TradeExecutionDirection.Buy Then
                    If currentPayload.Ticks(i) <= currentTrade.PotentialSL Then
                        Dim exitPrice As Double = currentPayload.Ticks(i)
                        With currentTrade
                            .TradingStatus = TradeExecutionStatus.Close
                            .ExitTime = currentPayload.PayloadDate
                            .ExitPrice = exitPrice
                            .ExitCondition = TradeExitCondition.StopLoss
                        End With
                        ret = True
                        tickCount = i
                        Exit For
                    ElseIf currentPayload.Ticks(i) >= currentTrade.PotentialTP Then
                        Dim exitPrice As Double = currentPayload.Ticks(i)
                        With currentTrade
                            .TradingStatus = TradeExecutionStatus.Close
                            .ExitTime = currentPayload.PayloadDate
                            .ExitPrice = exitPrice
                            .ExitCondition = TradeExitCondition.Target
                        End With
                        ret = True
                        tickCount = i
                        Exit For
                    End If
                    If Not ret Then
                        currentTrade.MaximumDrawDown = If(currentPayload.Low < currentTrade.MaximumDrawDown, currentPayload.Low, currentTrade.MaximumDrawDown)
                        currentTrade.MaximumDrawUp = If(currentPayload.High > currentTrade.MaximumDrawUp, currentPayload.High, currentTrade.MaximumDrawUp)
                    End If
                ElseIf currentTrade.EntryDirection = TradeExecutionDirection.Sell Then
                    If currentPayload.Ticks(i) >= currentTrade.PotentialSL Then
                        Dim exitPrice As Double = currentPayload.Ticks(i)
                        With currentTrade
                            .TradingStatus = TradeExecutionStatus.Close
                            .ExitTime = currentPayload.PayloadDate
                            .ExitPrice = exitPrice
                            .ExitCondition = TradeExitCondition.StopLoss
                        End With
                        ret = True
                        tickCount = i
                        Exit For
                    ElseIf currentPayload.Ticks(i) <= currentTrade.PotentialTP Then
                        Dim exitPrice As Double = currentPayload.Ticks(i)
                        With currentTrade
                            .TradingStatus = TradeExecutionStatus.Close
                            .ExitTime = currentPayload.PayloadDate
                            .ExitPrice = exitPrice
                            .ExitCondition = TradeExitCondition.Target
                        End With
                        ret = True
                        tickCount = i
                        Exit For
                    End If
                    If Not ret Then
                        currentTrade.MaximumDrawUp = If(currentPayload.Low < currentTrade.MaximumDrawUp, currentPayload.Low, currentTrade.MaximumDrawUp)
                        currentTrade.MaximumDrawDown = If(currentPayload.High > currentTrade.MaximumDrawDown, currentPayload.High, currentTrade.MaximumDrawDown)
                    End If
                End If
            Else
                With currentTrade
                    .TradingStatus = TradeExecutionStatus.Close
                    .ExitTime = currentPayload.PayloadDate
                    .ExitPrice = currentPayload.Ticks(i)
                    .ExitCondition = TradeExitCondition.ForceExit
                End With
                ret = True
                tickCount = i
                Exit For
            End If
            tickCount = i
        Next
        If Not ret And currentPayload.PayloadDate.TimeOfDay.Hours = endOfDay.TimeOfDay.Hours And currentPayload.PayloadDate.TimeOfDay.Minutes = endOfDay.TimeOfDay.Minutes Then
            With currentTrade
                .TradingStatus = TradeExecutionStatus.Close
                .ExitTime = currentPayload.PayloadDate
                .ExitPrice = currentPayload.Open
                .ExitCondition = TradeExitCondition.EndOfDay
            End With
            ret = True
        End If
        If ret = True Then
            If currentTrade.EntryDirection = TradeExecutionDirection.Buy Then
                With currentTrade
                    .ProfitLoss = CalculateProfitLoss(.TradingSymbol, .EntryPrice, .ExitPrice, .Quantity, typeOfStock)
                End With
            ElseIf currentTrade.EntryDirection = TradeExecutionDirection.Sell Then
                With currentTrade
                    .ProfitLoss = CalculateProfitLoss(.TradingSymbol, .ExitPrice, .EntryPrice, .Quantity, typeOfStock)
                End With
            End If
        End If

        Return ret
    End Function
    Public Function CalculateProfitLoss(ByVal stockName As String, ByVal buy As Double, ByVal sell As Double, ByVal quantity As Integer, ByVal typeOfStock As Strategy.TypeOfStock) As Double
        Dim outputCalculator As New Calculator.Output_Brokerage_Calculator
        Dim calculate As New Calculator.Brokerage_Calculator(_canceller)

        Select Case typeOfStock
            Case TypeOfStock.Cash
                calculate.Intraday_Equity(buy, sell, quantity, outputCalculator)
            Case TypeOfStock.Commodity
                stockName = stockName.Remove(stockName.Count - 8)
                calculate.Commodity_MCX(stockName, buy, sell, quantity, outputCalculator)
            Case TypeOfStock.Currency
                Throw New ApplicationException("Not Implemented")
            Case TypeOfStock.Futures
                calculate.FO_Futures(buy, sell, quantity, outputCalculator)
        End Select

        Return outputCalculator.NetProfitLoss
    End Function
    Public Function CalculateQuantity(ByVal stockName As String, ByVal buy As Double, ByVal sell As Double, ByVal NetProfitLossOfTrade As Double, ByVal typeOfStock As TypeOfStock) As Integer
        Dim outputCalculator As Calculator.Output_Brokerage_Calculator = Nothing
        Dim calculate As New Calculator.Brokerage_Calculator(_canceller)

        Dim quantity As Integer = 1
        Dim previousQuantity As Integer = 1
        For quantity = 1 To Integer.MaxValue
            outputCalculator = New Calculator.Output_Brokerage_Calculator
            Select Case typeOfStock
                Case TypeOfStock.Cash
                    calculate.Intraday_Equity(buy, sell, quantity, outputCalculator)
                Case TypeOfStock.Commodity
                    stockName = stockName.Remove(stockName.Count - 8)
                    calculate.Commodity_MCX(stockName, buy, sell, quantity, outputCalculator)
                Case TypeOfStock.Currency
                    Throw New ApplicationException("Not Implemented")
                Case TypeOfStock.Futures
                    calculate.FO_Futures(buy, sell, quantity, outputCalculator)
            End Select
            'calculate.Intraday_Equity(buy, sell, quantity, outputCalculator)
            If NetProfitLossOfTrade > 0 Then
                If outputCalculator.NetProfitLoss > NetProfitLossOfTrade Then
                    Exit For
                Else
                    previousQuantity = quantity
                End If
            ElseIf NetProfitLossOfTrade < 0 Then
                If outputCalculator.NetProfitLoss < NetProfitLossOfTrade Then
                    Exit For
                Else
                    previousQuantity = quantity
                End If
            End If
        Next
        Return previousQuantity
    End Function
    Public Function CalculateTargetOrStoploss(ByVal stockName As String, ByVal entryPrice As Double, ByVal quantity As Integer, ByVal NetProfitLossOfTrade As Double, ByVal tradeDirection As TradeExecutionDirection, ByVal typeOfStock As TypeOfStock) As Double
        Dim outputCalculator As Calculator.Output_Brokerage_Calculator = Nothing
        Dim calculate As New Calculator.Brokerage_Calculator(_canceller)

        Dim exitPrice As Double = entryPrice
        Dim previousExitPrice As Double = 0
        outputCalculator = New Calculator.Output_Brokerage_Calculator

        If NetProfitLossOfTrade > 0 Then
            While Not outputCalculator.NetProfitLoss > NetProfitLossOfTrade
                If tradeDirection = TradeExecutionDirection.Buy Then
                    Select Case typeOfStock
                        Case TypeOfStock.Cash
                            calculate.Intraday_Equity(entryPrice, exitPrice, quantity, outputCalculator)
                        Case TypeOfStock.Commodity
                            stockName = stockName.Remove(stockName.Count - 8)
                            calculate.Commodity_MCX(stockName, entryPrice, exitPrice, quantity, outputCalculator)
                        Case TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case TypeOfStock.Futures
                            calculate.FO_Futures(entryPrice, exitPrice, quantity, outputCalculator)
                    End Select
                    If outputCalculator.NetProfitLoss > NetProfitLossOfTrade Then
                        Exit While
                    Else
                        previousExitPrice = exitPrice
                    End If
                    exitPrice += 0.05
                ElseIf tradeDirection = TradeExecutionDirection.Sell Then
                    Select Case typeOfStock
                        Case TypeOfStock.Cash
                            calculate.Intraday_Equity(exitPrice, entryPrice, quantity, outputCalculator)
                        Case TypeOfStock.Commodity
                            stockName = stockName.Remove(stockName.Count - 8)
                            calculate.Commodity_MCX(stockName, exitPrice, entryPrice, quantity, outputCalculator)
                        Case TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case TypeOfStock.Futures
                            calculate.FO_Futures(exitPrice, entryPrice, quantity, outputCalculator)
                    End Select
                    If outputCalculator.NetProfitLoss > NetProfitLossOfTrade Then
                        Exit While
                    Else
                        previousExitPrice = exitPrice
                    End If
                    exitPrice -= 0.05
                End If
            End While
        ElseIf NetProfitLossOfTrade < 0 Then
            While Not outputCalculator.NetProfitLoss < NetProfitLossOfTrade
                If tradeDirection = TradeExecutionDirection.Buy Then
                    Select Case typeOfStock
                        Case TypeOfStock.Cash
                            calculate.Intraday_Equity(entryPrice, exitPrice, quantity, outputCalculator)
                        Case TypeOfStock.Commodity
                            stockName = stockName.Remove(stockName.Count - 8)
                            calculate.Commodity_MCX(stockName, entryPrice, exitPrice, quantity, outputCalculator)
                        Case TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case TypeOfStock.Futures
                            calculate.FO_Futures(entryPrice, exitPrice, quantity, outputCalculator)
                    End Select
                    If outputCalculator.NetProfitLoss < NetProfitLossOfTrade Then
                        Exit While
                    Else
                        previousExitPrice = exitPrice
                    End If
                    exitPrice -= 0.05
                ElseIf tradeDirection = TradeExecutionDirection.Sell Then
                    Select Case typeOfStock
                        Case TypeOfStock.Cash
                            calculate.Intraday_Equity(exitPrice, entryPrice, quantity, outputCalculator)
                        Case TypeOfStock.Commodity
                            stockName = stockName.Remove(stockName.Count - 8)
                            calculate.Commodity_MCX(stockName, exitPrice, entryPrice, quantity, outputCalculator)
                        Case TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case TypeOfStock.Futures
                            calculate.FO_Futures(exitPrice, entryPrice, quantity, outputCalculator)
                    End Select
                    If outputCalculator.NetProfitLoss < NetProfitLossOfTrade Then
                        Exit While
                    Else
                        previousExitPrice = exitPrice
                    End If
                    exitPrice += 0.05
                End If
            End While
        End If

        If Math.Abs(previousExitPrice - entryPrice) >= 0.1 Then
            Return Math.Round(previousExitPrice, 2)
        Else
            Return Math.Round(exitPrice, 2)
        End If
    End Function
    Public Sub MoveStopLoss(ByVal currentTrade As Trade, ByVal price As Double)
        currentTrade.PotentialSL = price
        currentTrade.StopLossRemark = "Stoploss Moved"
    End Sub
    Public Sub ModifyOrder(ByVal currentTrade As Trade, ByVal entryPrice As Double, ByVal targetPrice As Double, ByVal stopLossPrice As Double, ByVal quantity As Integer, ByVal signalCandle As Payload, Optional absoluteATR As Double = Nothing, Optional indicatorTime As Date = Nothing)
        With currentTrade
            .EntryPrice = entryPrice
            .PotentialTP = ConvertFloorCeling(targetPrice, 0.05, RoundOfType.Floor)
            .PotentialSL = If(.EntryDirection = TradeExecutionDirection.Buy, ConvertFloorCeling(stopLossPrice, 0.05, RoundOfType.Floor), ConvertFloorCeling(stopLossPrice, 0.05, RoundOfType.Celing))
            .Quantity = quantity
            '.EntryTime = signalCandle.PayloadDate
            '.SignalCandle = signalCandle
            '.TradingSymbol = signalCandle.TradingSymbol
            '.TradingDate = signalCandle.PayloadDate.Date
            .AbsoluteATR = absoluteATR
            .IndicatorCandleTime = indicatorTime
        End With
    End Sub
    Public Function GetPLForDay(ByVal currentDate As DateTime, stockName As String) As Double
        Dim ret As Double = 0
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso
            TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockName) Then
            ret = TradesTaken(currentDate)(stockName).Sum(Function(x)
                                                              If x.TradingStatus = TradeExecutionStatus.Close Then
                                                                  If x.EntryDirection = TradeExecutionDirection.Buy Then
                                                                      Return (x.ExitPrice - x.EntryPrice)
                                                                  ElseIf x.EntryDirection = TradeExecutionDirection.Sell Then
                                                                      Return (x.EntryPrice - x.ExitPrice)
                                                                  Else
                                                                      Return 0
                                                                  End If
                                                              Else
                                                                  Return 0
                                                              End If
                                                          End Function)
        End If
        Return ret
    End Function
    Public Function GetProfitLossForStock(ByVal stock As String, ByVal fromDate As Date, ByVal toDate As Date) As Double
        Dim ret As Double = 0
        Dim stockList As Dictionary(Of String, List(Of Trade)) = Nothing
        Dim currentDate As Date = fromDate
        While currentDate.Date <= toDate.Date
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate) Then
                stockList = TradesTaken(currentDate)
                If stockList.ContainsKey(stock) Then
                    ret += stockList(stock).Sum(Function(x)
                                                    Return x.ProfitLoss
                                                End Function)
                End If
            End If
            currentDate = currentDate.AddDays(1)
        End While
        Return ret
    End Function
    Public Function GetOverallProfitLoss(ByVal fromDate As Date, ByVal toDate As Date) As Double
        Dim ret As Double = 0
        Dim stockList As Dictionary(Of String, List(Of Trade)) = Nothing
        Dim currentDate As Date = fromDate
        While currentDate.Date <= toDate.Date
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate) Then
                stockList = TradesTaken(currentDate)
                For Each stock In stockList.Keys
                    ret += stockList(stock).Sum(Function(x)
                                                    Return x.ProfitLoss
                                                End Function)
                Next
            End If
            currentDate = currentDate.AddDays(1)
        End While
        Return ret
    End Function
    Public Function CountTradesForDay(ByVal currentDate As Date, ByVal stockName As String) As Integer
        Dim ret As Integer = 0
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso
            TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockName) Then
            ret = TradesTaken(currentDate)(stockName).LongCount(Function(x)
                                                                    Return x.ExitCondition = TradeExitCondition.StopLoss OrElse
                                                                            x.ExitCondition = TradeExitCondition.Target OrElse
                                                                            x.ExitCondition = TradeExitCondition.ForceExit
                                                                End Function)
        End If
        Return ret
    End Function
    Public Function GetHeikenAshiStatus(ByVal heikenAshiCandle As Payload) As String
        Dim status As String = Nothing
        If heikenAshiCandle.CandleColor = Color.Green Then
            If heikenAshiCandle.CandleWicks.Bottom = 0 Then
                status = "LL"
            Else
                status = "L"
            End If
        ElseIf heikenAshiCandle.CandleColor = Color.Red Then
            If heikenAshiCandle.CandleWicks.Top = 0 Then
                status = "SS"
            Else
                status = "S"
            End If
        Else
            status = "None"
        End If
        Return status
    End Function
    Public Function GetSpecificTime(ByVal signalTime As Date, ByVal specificPayload As Dictionary(Of Date, Payload), ByVal position As Integer) As Date
        Dim ret As Date = Date.MinValue
        Dim requiredPayloads As IEnumerable(Of KeyValuePair(Of Date, Payload)) = Nothing

        requiredPayloads = specificPayload.Where(Function(x)
                                                     Return x.Key <= signalTime
                                                 End Function)

        If requiredPayloads IsNot Nothing AndAlso requiredPayloads.Count >= position Then
            ret = requiredPayloads(requiredPayloads.Count - position).Key
        End If

        Return ret
    End Function
    Public Function GetPreviousXMinuteCandleTime(ByVal currentTime As Date, ByVal totalXMinutePayload As Dictionary(Of Date, Payload), ByVal timeFrame As Integer) As Date
        Dim ret As Date = Nothing

        If totalXMinutePayload IsNot Nothing AndAlso totalXMinutePayload.Count > 0 Then
            ret = totalXMinutePayload.Keys.LastOrDefault(Function(x)
                                                             Return x <= currentTime.AddMinutes(-timeFrame)
                                                         End Function)
        End If

        Return ret
    End Function
#End Region

#Region "Private Functions"
    Private Sub GetFirstProfitableTradeOfDay(tradingDate As Date, tradingSymbol As String, currentTrade As Trade)
        Dim ret As Boolean = False

        If tradingSymbol.Contains("FUT") Then
            tradingSymbol = tradingSymbol.Remove(currentTrade.TradingSymbol.Count - 8)
        Else
            tradingSymbol = currentTrade.TradingSymbol
        End If

        If TradesTaken Is Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradingDate) Then
            Dim stockTrade As Dictionary(Of String, List(Of Trade)) = TradesTaken(tradingDate)
            If stockTrade IsNot Nothing AndAlso stockTrade.Count > 0 AndAlso stockTrade.ContainsKey(tradingSymbol) Then
                Dim tradeList As List(Of Trade) = stockTrade(tradingSymbol).FindAll(Function(x)
                                                                                        Return x.TradingStatus <> TradeExecutionStatus.Cancel
                                                                                    End Function)
                If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                    Dim modifiedTradeList As List(Of Trade) = tradeList.FindAll(Function(y)
                                                                                    Return y.ExitTime <= currentTrade.ExitTime
                                                                                End Function)
                    If modifiedTradeList IsNot Nothing AndAlso modifiedTradeList.Count > 0 Then
                        For Each item In modifiedTradeList
                            If item.ProfitLoss > 0 Then
                                ret = True
                                Exit For
                            End If
                        Next
                    Else
                        currentTrade.FirstProfitableTradeOfDay = False
                    End If
                Else
                    currentTrade.FirstProfitableTradeOfDay = False
                End If
                If ret Then
                    currentTrade.FirstProfitableTradeOfDay = True
                End If
            End If
        End If
    End Sub
#End Region

#Region "Print To Excel"
    Public Overridable Sub PrintTrades(filename As String, stock As String)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
            Dim cts As New CancellationTokenSource

            Dim filepath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, filename)
            Dim openStatus As ExcelHelper.ExcelOpenStatus = ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite
            If System.IO.File.Exists(filepath) Then
                openStatus = ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite
            End If
            Using excelWriter As New ExcelHelper(System.IO.Path.Combine(My.Application.Info.DirectoryPath, filepath), openStatus, ExcelHelper.ExcelSaveType.XLS_XLSX, cts)
                Dim rowCtr As Integer = 0
                Dim colCtr As Integer = 0

                If openStatus = ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite Then
                    rowCtr += 1
                    'Headers
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Date")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Symbol")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 20)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Capital Required With Margin")
                    excelWriter.SetCellWrap(rowCtr, colCtr, True)
                    excelWriter.SetCellHeight(rowCtr, colCtr, 30)
                    excelWriter.SetCellWidth(rowCtr, colCtr, 25)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Direction")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 10)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Buy Price")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 10)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Sell Price")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 10)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Quantity")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 10)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Entry Time")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 10)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Exit Time")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 10)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Exit Condition")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 12)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "StopLoss Remark")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 12)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Profit & Loss")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Candle Range")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Candle Range %")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Absolute ATR")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "ATR %")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Signal Candle Volume")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Signal Candle Volume SMA")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "PrevioustoSignal Candle Volume")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                    colCtr += 1
                    excelWriter.SetData(rowCtr, colCtr, "Duration Of Trade")
                    excelWriter.SetCellWidth(rowCtr, colCtr, 15)
                    excelWriter.SetCellBackColor(rowCtr, colCtr, ExcelHelper.XLColor.BlueGray)
                Else
                    rowCtr = excelWriter.GetLastRow()
                End If

                Dim dateCtr As Integer = 0
                For Each tempkeys In TradesTaken.Keys
                    dateCtr += 1
                    OnHeartbeat(String.Format("Excel printing for Date: {0} [{1} of {2}]", tempkeys.Date.ToShortDateString, dateCtr, TradesTaken.Count))
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(tempkeys)
                    If stockTrades.Count > 0 Then
                        Dim stockCtr As Integer = 0
                        'For Each stock In stockTrades.Keys
                        stockCtr += 1
                        OnHeartbeat(String.Format("Excel printing for Stock: {0} [{1} of {2}]", stock, stockCtr, stockTrades.Count))
                        Dim tradeList As List(Of Trade) = stockTrades(stock).FindAll(Function(x)
                                                                                         Return x.TradingStatus <> TradeExecutionStatus.Cancel
                                                                                     End Function)
                        Dim tradeCtr As Integer = 0
                        If tradeList IsNot Nothing Then
                            For Each tradeTaken In tradeList
                                tradeCtr += 1
                                OnHeartbeat(String.Format("Excel printing: {0} of {1}", tradeCtr, tradeList.Count))
                                'Data
                                rowCtr += 1
                                colCtr = 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.TradingDate.ToString("dd-MMM-yyyy"))
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.TradingSymbol)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.CapitalRequiredWithMargin, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.EntryDirection.ToString)
                                colCtr += 1
                                If tradeTaken.EntryDirection = TradeExecutionDirection.Buy Then
                                    excelWriter.SetData(rowCtr, colCtr, tradeTaken.EntryPrice, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                    colCtr += 1
                                    excelWriter.SetData(rowCtr, colCtr, tradeTaken.ExitPrice, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                    colCtr += 1
                                ElseIf tradeTaken.EntryDirection = TradeExecutionDirection.Sell Then
                                    excelWriter.SetData(rowCtr, colCtr, tradeTaken.ExitPrice, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                    colCtr += 1
                                    excelWriter.SetData(rowCtr, colCtr, tradeTaken.EntryPrice, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                    colCtr += 1
                                End If
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.Quantity, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.EntryTime.ToString("HH:mm:ss"))
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.ExitTime.ToString("HH:mm:ss"))
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.ExitCondition)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.StopLossRemark)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.ProfitLoss, "##,##,###.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.SignalCandle.CandleRange, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.SignalCandle.CandleRangePercentage, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.AbsoluteATR, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.ATRPercentage, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.SignalCandle.Volume, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.SignalCandleVolumeSMA, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.SignalCandle.PreviousCandlePayload.Volume, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                                excelWriter.SetData(rowCtr, colCtr, tradeTaken.DurationOfTrade.TotalMinutes, "#,##0.00", ExcelHelper.XLAlign.Right)
                                colCtr += 1
                            Next

                        End If
                        'Next
                    End If
                Next

                excelWriter.SaveExcel()
            End Using
        End If
    End Sub
    Public Overridable Sub PrintArrayToExcel(filename As String, stock As String)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
            Dim cts As New CancellationTokenSource
            OnHeartbeat("Opening Excel.....")
            Dim filepath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, filename)
            Dim openStatus As ExcelHelper.ExcelOpenStatus = ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite
            If System.IO.File.Exists(filepath) Then
                openStatus = ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite
            End If
            Using excelWriter As New ExcelHelper(System.IO.Path.Combine(My.Application.Info.DirectoryPath, filepath), openStatus, ExcelHelper.ExcelSaveType.XLS_XLSX, cts)
                Dim rowCtr As Integer = 0
                Dim colCtr As Integer = 0

                Dim rowCount As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    rowCount = TradesTaken.Sum(Function(x)
                                                   Dim stockTrades = x.Value
                                                   Return stockTrades.Sum(Function(y)
                                                                              If y.Key = stock Then
                                                                                  Dim trades = y.Value.FindAll((Function(z)
                                                                                                                    Return z.TradingStatus <> TradeExecutionStatus.Cancel
                                                                                                                End Function))
                                                                                  Return trades.Count
                                                                              Else
                                                                                  Return 0
                                                                              End If
                                                                          End Function)
                                               End Function)
                    'rowCount = TradesTaken.Sum(Function(x)
                    '                               Dim stockTrades = x.Value
                    '                               Return stockTrades.Sum(Function(y)
                    '                                                          If y.Key = stock Then
                    '                                                              Return y.Value.Count
                    '                                                          Else
                    '                                                              Return 0
                    '                                                          End If
                    '                                                      End Function)
                    '                           End Function)
                End If

                Dim rawData(rowCount, 0) As Object

                If rowCtr = 0 Then
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Trading Date"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Trading Symbol"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Capital Required With Margin"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Type"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Direction"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Buy Price"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Sell Price"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Quantity"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Exit Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Duration Of Trade"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Condition"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Exit Condition"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Stoploss Remark"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "P & L Point"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Profit & Loss"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Maximum Draw Up"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Maximum Draw Down"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Signal Candle Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Absolute ATR"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "ATR Candle Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Month"

                    rowCtr += 1
                End If
                Dim dateCtr As Integer = 0
                For Each tempkeys In TradesTaken.Keys
                    dateCtr += 1
                    OnHeartbeat(String.Format("Excel printing for Date: {0} [{1} of {2}]", tempkeys.Date.ToShortDateString, dateCtr, TradesTaken.Count))
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(tempkeys)
                    If stockTrades.Count > 0 Then
                        Dim stockCtr As Integer = 0
                        'For Each stock In stockTrades.Keys
                        stockCtr += 1
                        OnHeartbeat(String.Format("Excel printing for Stock: {0} [{1} of {2}]", stock, stockCtr, stockTrades.Count))
                        If stockTrades.ContainsKey(stock) Then
                            Dim tradeList As List(Of Trade) = stockTrades(stock).FindAll(Function(x)
                                                                                             Return x.TradingStatus <> TradeExecutionStatus.Cancel
                                                                                         End Function)
                            'Dim tradeList As List(Of Trade) = stockTrades(stock)
                            Dim tradeCtr As Integer = 0
                            If tradeList IsNot Nothing Then
                                For Each tradeTaken In tradeList
                                    tradeCtr += 1
                                    OnHeartbeat(String.Format("Excel printing: {0} of {1}", tradeCtr, tradeList.Count))
                                    'Data
                                    colCtr = 0

                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.TradingDate.ToString("dd-MMM-yyyy")
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.TradingSymbol
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.CapitalRequiredWithMargin
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.EntryType
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.EntryDirection.ToString
                                    colCtr += 1
                                    If tradeTaken.EntryDirection = TradeExecutionDirection.Buy Then
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 2)
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 2)
                                        colCtr += 1
                                    ElseIf tradeTaken.EntryDirection = TradeExecutionDirection.Sell Then
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 2)
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 2)
                                        colCtr += 1
                                    End If
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.Quantity
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.EntryTime.ToString("HH:mm:ss")
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.ExitTime.ToString("HH:mm:ss")
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.DurationOfTrade.TotalMinutes
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.EntryCondition
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.ExitCondition.ToString
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.StopLossRemark
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.PLPoint
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.ProfitLoss
                                    colCtr += 1
                                    If tradeTaken.EntryDirection = TradeExecutionDirection.Buy Then
                                        If tradeTaken.ExitCondition = TradeExitCondition.StopLoss Or tradeTaken.ExitCondition = TradeExitCondition.ForceExit Then
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawUp - tradeTaken.EntryPrice), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawUp, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Nothing
                                            colCtr += 1
                                        ElseIf tradeTaken.ExitCondition = TradeExitCondition.Target Then
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Nothing
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawDown), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawDown, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                        Else
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawUp - tradeTaken.EntryPrice), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawUp, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawDown), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawDown, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                        End If
                                    ElseIf tradeTaken.EntryDirection = TradeExecutionDirection.Sell Then
                                        If tradeTaken.ExitCondition = TradeExitCondition.StopLoss Or tradeTaken.ExitCondition = TradeExitCondition.ForceExit Then
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawUp), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawUp, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Nothing
                                            colCtr += 1
                                        ElseIf tradeTaken.ExitCondition = TradeExitCondition.Target Then
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Nothing
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawDown - tradeTaken.EntryPrice), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawDown, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                        Else
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawUp), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawUp, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawDown - tradeTaken.EntryPrice), 2)
                                            rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawDown, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                            colCtr += 1
                                        End If
                                    End If
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.SignalCandle.PayloadDate.ToString("HH:mm:ss")
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.AbsoluteATR
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = tradeTaken.IndicatorCandleTime.ToString("HH:mm:ss")
                                    colCtr += 1
                                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                    rawData(rowCtr, colCtr) = String.Format("{0}-{1}", tradeTaken.TradingDate.ToString("yyyy"), tradeTaken.TradingDate.ToString("MM"))

                                    rowCtr += 1
                                Next
                            End If
                        End If
                        'Next
                    End If
                Next

                Dim range As String = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                RaiseEvent Heartbeat("Writing from memory to excel...")
                excelWriter.WriteArrayToExcel(rawData, range)
                Erase rawData
                rawData = Nothing
                RaiseEvent Heartbeat("Saving excel...")
                excelWriter.SaveExcel()
            End Using
        End If
    End Sub
    Public Overridable Sub PrintArrayToExcel(filename As String)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
            Dim cts As New CancellationTokenSource

            Dim filepath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, filename)
            Dim openStatus As ExcelHelper.ExcelOpenStatus = ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite
            If System.IO.File.Exists(filepath) Then
                openStatus = ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite
            End If
            Using excelWriter As New ExcelHelper(System.IO.Path.Combine(My.Application.Info.DirectoryPath, filepath), openStatus, ExcelHelper.ExcelSaveType.XLS_XLSX, cts)
                Dim rowCtr As Integer = 0
                Dim colCtr As Integer = 0

                Dim rowCount As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    rowCount = TradesTaken.Sum(Function(x)
                                                   Dim stockTrades = x.Value
                                                   Return stockTrades.Sum(Function(y)
                                                                              Dim trades = y.Value.FindAll((Function(z)
                                                                                                                Return z.TradingStatus <> TradeExecutionStatus.Cancel
                                                                                                            End Function))
                                                                              Return trades.Count
                                                                          End Function)
                                               End Function)
                End If

                Dim rawData(rowCount, 0) As Object

                If rowCtr = 0 Then
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Trading Date"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Trading Symbol"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Capital Required With Margin"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Type"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Direction"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Buy Price"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Sell Price"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Quantity"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Exit Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Duration Of Trade"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Entry Condition"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Exit Condition"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Stoploss Remark"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "P & L Point"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Profit & Loss"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Maximum Draw Up"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Maximum Draw Down"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Signal Candle Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Absolute ATR"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "ATR Candle Time"
                    colCtr += 1
                    If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                    rawData(rowCtr, colCtr) = "Month"

                    rowCtr += 1
                End If
                Dim dateCtr As Integer = 0
                For Each tempkeys In TradesTaken.Keys
                    dateCtr += 1
                    OnHeartbeat(String.Format("Excel printing for Date: {0} [{1} of {2}]", tempkeys.Date.ToShortDateString, dateCtr, TradesTaken.Count))
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(tempkeys)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        Dim stockCtr As Integer = 0
                        For Each stock In stockTrades.Keys
                            stockCtr += 1
                            If stockTrades.ContainsKey(stock) Then
                                Dim tradeList As List(Of Trade) = stockTrades(stock).FindAll(Function(x)
                                                                                                 Return x.TradingStatus <> TradeExecutionStatus.Cancel
                                                                                             End Function)
                                'Dim tradeList As List(Of Trade) = stockTrades(stock)
                                Dim tradeCtr As Integer = 0
                                If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                                    For Each tradeTaken In tradeList
                                        tradeCtr += 1
                                        OnHeartbeat(String.Format("Excel printing: {0} of {1}", tradeCtr, tradeList.Count))
                                        'Data
                                        colCtr = 0

                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.TradingDate.ToString("dd-MMM-yyyy")
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.TradingSymbol
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.CapitalRequiredWithMargin
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.EntryType
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.EntryDirection.ToString
                                        colCtr += 1
                                        If tradeTaken.EntryDirection = TradeExecutionDirection.Buy Then
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 2)
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 2)
                                            colCtr += 1
                                        ElseIf tradeTaken.EntryDirection = TradeExecutionDirection.Sell Then
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 2)
                                            colCtr += 1
                                            If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                            rawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 2)
                                            colCtr += 1
                                        End If
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.Quantity
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.EntryTime.ToString("HH:mm:ss")
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.ExitTime.ToString("HH:mm:ss")
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.DurationOfTrade.TotalMinutes
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.EntryCondition
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.ExitCondition.ToString
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.StopLossRemark
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.PLPoint
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.ProfitLoss
                                        colCtr += 1
                                        If tradeTaken.EntryDirection = TradeExecutionDirection.Buy Then
                                            If tradeTaken.ExitCondition = TradeExitCondition.StopLoss Or tradeTaken.ExitCondition = TradeExitCondition.ForceExit Then
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawUp - tradeTaken.EntryPrice), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawUp, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                rawData(rowCtr, colCtr) = Nothing
                                                colCtr += 1
                                            ElseIf tradeTaken.ExitCondition = TradeExitCondition.Target Then
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                rawData(rowCtr, colCtr) = Nothing
                                                colCtr += 1
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawDown), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawDown, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                            Else
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawUp - tradeTaken.EntryPrice), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawUp, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawDown), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.EntryPrice, tradeTaken.MaximumDrawDown, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                            End If
                                        ElseIf tradeTaken.EntryDirection = TradeExecutionDirection.Sell Then
                                            If tradeTaken.ExitCondition = TradeExitCondition.StopLoss Or tradeTaken.ExitCondition = TradeExitCondition.ForceExit Then
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawUp), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawUp, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                rawData(rowCtr, colCtr) = Nothing
                                                colCtr += 1
                                            ElseIf tradeTaken.ExitCondition = TradeExitCondition.Target Then
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                rawData(rowCtr, colCtr) = Nothing
                                                colCtr += 1
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawDown - tradeTaken.EntryPrice), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawDown, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                            Else
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.EntryPrice - tradeTaken.MaximumDrawUp), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawUp, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                                If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                                'rawData(rowCtr, colCtr) = Math.Round((tradeTaken.MaximumDrawDown - tradeTaken.EntryPrice), 2)
                                                rawData(rowCtr, colCtr) = CalculateProfitLoss(tradeTaken.TradingSymbol, tradeTaken.MaximumDrawDown, tradeTaken.EntryPrice, tradeTaken.Quantity, tradeTaken.TypeOfStock)
                                                colCtr += 1
                                            End If
                                        End If
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.SignalCandle.PayloadDate.ToString("HH:mm:ss")
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.AbsoluteATR
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = tradeTaken.IndicatorCandleTime.ToString("HH:mm:ss")
                                        colCtr += 1
                                        If colCtr > UBound(rawData, 2) Then ReDim Preserve rawData(UBound(rawData, 1), 0 To UBound(rawData, 2) + 1)
                                        rawData(rowCtr, colCtr) = String.Format("{0}-{1}", tradeTaken.TradingDate.ToString("yyyy"), tradeTaken.TradingDate.ToString("MM"))

                                        rowCtr += 1
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next

                Dim range As String = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                RaiseEvent Heartbeat("Writing from memory to excel...")
                excelWriter.WriteArrayToExcel(rawData, range)
                Erase rawData
                rawData = Nothing
                RaiseEvent Heartbeat("Saving excel...")
                excelWriter.SaveExcel()
            End Using
        End If
    End Sub
#End Region
End Class
