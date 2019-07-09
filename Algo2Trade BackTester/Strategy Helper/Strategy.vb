Imports System.Threading
Imports Utilities.DAL
Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.IO

Public MustInherit Class Strategy

#Region "Constructor"
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan)
        Me.Canceller = canceller
        Strategy.TickSize = tickSize
        Me.EODExitTime = eodExitTime
        Me.LastTradeEntryTime = lastTradeEntryTime
        Me.ExchangeStartTime = exchangeStartTime
        Me.ExchangeEndTime = exchangeEndTime
        AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
        AddHandler Cmn.WaitingFor, AddressOf OnWaitingFor
        AddHandler Cmn.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
        AddHandler Cmn.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
    End Sub
#End Region

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

#Region "Constants"
    Const MARGIN_MULTIPLIER As Integer = 30
    Const DEFAULT_TICK_SIZE As Double = 0.05
    Const EOD_EXIT_TIME As String = "15:00:00"
    Const LAST_TRADE_ENTRY_TIME As String = "15:00:00"
    Const EXCHANGE_START_TIME As String = "09:15:00"
    Const EXCHANGE_END_TIME As String = "15:30:00"
#End Region

#Region "Protected Variables"
    Protected Canceller As CancellationTokenSource
    Protected Cmn As Common = New Common(Canceller)
    Protected TradesTaken As Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
    Protected Shared TickSize As Double = DEFAULT_TICK_SIZE
    Protected EODExitTime As TimeSpan = TimeSpan.Parse(EOD_EXIT_TIME)
    Protected LastTradeEntryTime As TimeSpan = TimeSpan.Parse(LAST_TRADE_ENTRY_TIME)
    Protected ExchangeStartTime As TimeSpan = TimeSpan.Parse(EXCHANGE_START_TIME)
    Protected ExchangeEndTime As TimeSpan = TimeSpan.Parse(EXCHANGE_END_TIME)
    Public Shared MarginMultiplier As Integer = MARGIN_MULTIPLIER
#End Region

#Region "Property"
    Public Property InitialCapital As Decimal = Decimal.MaxValue
    Public Property CapitalForPumpIn As Decimal = Decimal.MaxValue
    Public Property MinimumEarnedCapitalToWithdraw As Decimal = Decimal.MaxValue
    Public Property AmountToBeWithdrawn As Decimal = Decimal.MaxValue

    Public Property NumberOfTradeableStockPerDay As Integer = Integer.MaxValue
    Public Property PerTradeMaxProfitPercentage As Decimal = Decimal.MaxValue
    Public Property PerTradeMaxLossPercentage As Decimal = Decimal.MinValue
    Public Property PerStockMaxProfitPercentage As Decimal = Decimal.MaxValue
    Public Property PerStockMaxLossPercentage As Decimal = Decimal.MinValue
    Public Property OverAllProfitPerDay As Double = Double.MaxValue
    Public Property OverAllLossPerDay As Double = Double.MinValue
    Protected ReadOnly Property AllPLAfterBrokerage(ByVal currentDate As Date) As Double
        Get
            Dim ret As Double = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        ret += StockPLAfterBrokerage(currentDate, stock)
                    Next
                End If
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property AllPLBeforeBrokerage(ByVal currentDate As Date) As Double
        Get
            Dim ret As Double = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        ret += StockPLBeforeBrokerage(currentDate, stock)
                    Next
                End If
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property StockPLPoint(ByVal currentDate As Date, ByVal currentStockName As String) As Double
        Get
            Dim ret As Double = Nothing
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(currentStockName) Then
                ret = TradesTaken(currentDate.Date)(currentStockName).Sum(Function(x)
                                                                              If x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                                                                                  Return x.PLPoint
                                                                              Else
                                                                                  Return 0
                                                                              End If
                                                                          End Function)
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property StockPLAfterBrokerage(ByVal currentDate As Date, ByVal currentStockName As String) As Double
        Get
            Dim ret As Double = Nothing
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(currentStockName) Then
                ret = TradesTaken(currentDate.Date)(currentStockName).Sum(Function(x)
                                                                              If x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                                                                                  Return x.PLAfterBrokerage
                                                                              Else
                                                                                  Return 0
                                                                              End If
                                                                          End Function)
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property StockPLBeforeBrokerage(ByVal currentDate As Date, ByVal currentStockName As String) As Double
        Get
            Dim ret As Double = Nothing
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(currentStockName) Then
                ret = TradesTaken(currentDate.Date)(currentStockName).Sum(Function(x)
                                                                              If x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                                                                                  Return x.PLBeforeBrokerage
                                                                              Else
                                                                                  Return 0
                                                                              End If
                                                                          End Function)
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property NumberOfTradesPerStockPerDay(ByVal currentDate As Date, ByVal currentStockName As String) As Integer
        Get
            Dim ret As Integer = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(currentStockName) Then
                Dim tradeList As List(Of Trade) = TradesTaken(currentDate.Date)(currentStockName).FindAll(Function(x)
                                                                                                              Return x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso
                                                                                                              x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open
                                                                                                          End Function)
                If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                    Dim artnrGroups = From a In tradeList
                                      Group a By Key = a.Tag Into Group
                                      Select artnr = Key, numbersCount = Group.Count()

                    If artnrGroups IsNot Nothing AndAlso artnrGroups.Count > 0 Then
                        ret = artnrGroups.Count
                    End If
                End If
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property NumberOfTradesPerDay(ByVal currentDate As Date) As Integer
        Get
            Dim ret As Integer = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        ret += NumberOfTradesPerStockPerDay(currentDate, stock)
                    Next
                End If
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property NumberOfTradesPerStockPerDayWithoutBreakevenExit(ByVal currentDate As Date, ByVal currentStockName As String) As Integer
        Get
            Dim ret As Integer = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(currentStockName) Then
                ret = Me.NumberOfTradesPerStockPerDay(currentDate, currentStockName)
                Dim beakevenTradeList As List(Of Trade) = TradesTaken(currentDate.Date)(currentStockName).FindAll(Function(x)
                                                                                                                      Return x.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso x.PLPoint > 0
                                                                                                                  End Function)
                If beakevenTradeList IsNot Nothing AndAlso beakevenTradeList.Count > 0 Then
                    Dim artnrGroups = From a In beakevenTradeList
                                      Group a By Key = a.Tag Into Group
                                      Select artnr = Key, numbersCount = Group.Count()

                    If artnrGroups IsNot Nothing AndAlso artnrGroups.Count > 0 Then
                        ret = ret - artnrGroups.Count
                    End If
                End If
            End If
            Return ret
        End Get
    End Property
    Protected ReadOnly Property NumberOfTradesPerDayWithoutBreakevenExit(ByVal currentDate As Date) As Integer
        Get
            Dim ret As Integer = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        ret += NumberOfTradesPerStockPerDayWithoutBreakevenExit(currentDate, stock)
                    Next
                End If
            End If
            Return ret
        End Get
    End Property

    Private _AllMaxDrawDownPLAfterBrokerage As Double = Double.MaxValue
    Public ReadOnly Property AllMaxDrawDownPLAfterBrokerage(ByVal currentDate As Date) As Double
        Get
            Dim pl As Double = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        pl += StockPLAfterBrokerage(currentDate, stock)
                    Next
                End If
                _AllMaxDrawDownPLAfterBrokerage = Math.Min(_AllMaxDrawDownPLAfterBrokerage, pl)
            End If
            Return _AllMaxDrawDownPLAfterBrokerage
        End Get
    End Property

    Private _AllMaxDrawUpPLAfterBrokerage As Double = Double.MinValue
    Public ReadOnly Property AllMaxDrawUpPLAfterBrokerage(ByVal currentDate As Date) As Double
        Get
            Dim pl As Double = 0
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                    For Each stock In stockTrades.Keys
                        pl += StockPLAfterBrokerage(currentDate, stock)
                    Next
                End If
                _AllMaxDrawUpPLAfterBrokerage = Math.Max(_AllMaxDrawUpPLAfterBrokerage, pl)
            End If
            Return _AllMaxDrawUpPLAfterBrokerage
        End Get
    End Property
    Protected Property CapitalMovement As Dictionary(Of Date, List(Of Capital))
#End Region

#Region "Public Functions"
    Public Function IsTradeActive(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType, Optional ByVal tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None) As Boolean
        Dim ret As Boolean = False
        Dim tradeDate As Date = currentMinutePayload.PayloadDate.Date

        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
            Dim tradeList As List(Of Trade) = Nothing
            If tradeDirection = Trade.TradeExecutionDirection.None Then
                tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                   Return x.SquareOffType = tradeType AndAlso
                                                                                                   x.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress
                                                                                               End Function)
            Else
                tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                   Return x.SquareOffType = tradeType AndAlso
                                                                                                   x.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso
                                                                                                   x.EntryDirection = tradeDirection
                                                                                               End Function)
            End If
            ret = tradeList IsNot Nothing AndAlso tradeList.Count > 0
        End If

        Return ret
    End Function
    Public Function IsTradeOpen(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType, Optional ByVal tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None) As Boolean
        Dim ret As Boolean = False
        Dim tradeDate As Date = currentMinutePayload.PayloadDate.Date

        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
            Dim tradeList As List(Of Trade) = Nothing
            If tradeDirection = Trade.TradeExecutionDirection.None Then
                tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                   Return x.SquareOffType = tradeType AndAlso x.TradeCurrentStatus = Trade.TradeExecutionStatus.Open
                                                                                               End Function)
            Else
                tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                   Return x.SquareOffType = tradeType AndAlso x.TradeCurrentStatus = Trade.TradeExecutionStatus.Open AndAlso x.EntryDirection = tradeDirection
                                                                                               End Function)
            End If
            ret = tradeList IsNot Nothing AndAlso tradeList.Count > 0
        End If

        Return ret
    End Function
    Public Function CalculateBuffer(ByVal price As Double, ByVal floorOrCeiling As RoundOfType) As Double
        Dim bufferPrice As Double = Nothing
        'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
        bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, TickSize, floorOrCeiling)
        Return bufferPrice
    End Function
    Public Sub PlaceOrModifyOrder(ByVal currentTrade As Trade, ByVal modifyTrade As Trade)
        Dim usableTrade As Trade = Nothing
        Dim capitalToBeAdded As Double = Nothing
        Dim capitalToBeReleased As Double = Nothing
        If modifyTrade Is Nothing Then  'Addition
            usableTrade = currentTrade
            capitalToBeAdded = currentTrade.CapitalRequiredWithMargin
            capitalToBeReleased = 0
        Else    'Modification
            If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then Throw New ApplicationException("Supplied trade is not open, cannot modify")
            usableTrade = modifyTrade
            capitalToBeAdded = modifyTrade.CapitalRequiredWithMargin
            capitalToBeReleased = currentTrade.CapitalRequiredWithMargin
        End If
        Dim tradeDate As Date = currentTrade.TradingDate.Date
        Dim tradingSymbol As String = currentTrade.TradingSymbol

        Dim lastTradingTime As Date = New Date(usableTrade.TradingDate.Year, usableTrade.TradingDate.Month, usableTrade.TradingDate.Day, LastTradeEntryTime.Hours, LastTradeEntryTime.Minutes, LastTradeEntryTime.Seconds)

        If usableTrade.EntryTime < lastTradingTime Then
            currentTrade.UpdateTrade(usableTrade)

            With usableTrade
                If .EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    currentTrade.UpdateTrade(EntryPrice:=ConvertFloorCeling(.EntryPrice, TickSize, RoundOfType.Celing),
                                         PotentialTarget:=ConvertFloorCeling(.PotentialTarget, TickSize, RoundOfType.Celing),
                                         PotentialStopLoss:=ConvertFloorCeling(.PotentialStopLoss, TickSize, RoundOfType.Floor))
                ElseIf .EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    currentTrade.UpdateTrade(EntryPrice:=ConvertFloorCeling(.EntryPrice, TickSize, RoundOfType.Floor),
                                         PotentialTarget:=ConvertFloorCeling(.PotentialTarget, TickSize, RoundOfType.Floor),
                                         PotentialStopLoss:=ConvertFloorCeling(.PotentialStopLoss, TickSize, RoundOfType.Floor))
                End If
                currentTrade.UpdateTrade(EntryTime:= .EntryTime, Quantity:= .Quantity, TradeCurrentStatus:=Trade.TradeExecutionStatus.Open)
            End With

            If modifyTrade Is Nothing Then
                If currentTrade.EntryTime < lastTradingTime Then
                    If TradesTaken.ContainsKey(tradeDate) Then
                        If TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
                            TradesTaken(tradeDate)(tradingSymbol).Add(currentTrade)
                        Else
                            TradesTaken(tradeDate).Add(tradingSymbol, New List(Of Trade) From {currentTrade})
                        End If
                    Else
                        TradesTaken.Add(tradeDate, New Dictionary(Of String, List(Of Trade)) From {{tradingSymbol, New List(Of Trade) From {currentTrade}}})
                    End If
                End If
            End If

            InsertCapitalRequired(currentTrade.EntryTime, capitalToBeAdded, capitalToBeReleased, "Placed Or Modify Order")
        End If

        ''Indibar
        ''If usableTrade.EntryTime < lastTradingTime Then
        'currentTrade.UpdateTrade(usableTrade)

        'With usableTrade
        '    If .EntryDirection = Trade.TradeExecutionDirection.Buy Then
        '        currentTrade.UpdateTrade(EntryPrice:=ConvertFloorCeling(.EntryPrice, TickSize, RoundOfType.Celing),
        '                             PotentialTarget:=ConvertFloorCeling(.PotentialTarget, TickSize, RoundOfType.Celing),
        '                             PotentialStopLoss:=ConvertFloorCeling(.PotentialStopLoss, TickSize, RoundOfType.Floor))
        '    ElseIf .EntryDirection = Trade.TradeExecutionDirection.Sell Then
        '        currentTrade.UpdateTrade(EntryPrice:=ConvertFloorCeling(.EntryPrice, TickSize, RoundOfType.Floor),
        '                             PotentialTarget:=ConvertFloorCeling(.PotentialTarget, TickSize, RoundOfType.Floor),
        '                             PotentialStopLoss:=ConvertFloorCeling(.PotentialStopLoss, TickSize, RoundOfType.Floor))
        '    End If
        '    currentTrade.UpdateTrade(EntryTime:= .EntryTime, Quantity:= .Quantity, TradeCurrentStatus:=Trade.TradeExecutionStatus.Open)
        'End With

        'If modifyTrade Is Nothing Then
        '    'If currentTrade.EntryTime < lastTradingTime Then
        '    If TradesTaken.ContainsKey(tradeDate) Then
        '        If TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
        '            TradesTaken(tradeDate)(tradingSymbol).Add(currentTrade)
        '        Else
        '            TradesTaken(tradeDate).Add(tradingSymbol, New List(Of Trade) From {currentTrade})
        '        End If
        '    Else
        '        If currentTrade.EntryTime < lastTradingTime Then
        '            TradesTaken.Add(tradeDate, New Dictionary(Of String, List(Of Trade)) From {{tradingSymbol, New List(Of Trade) From {currentTrade}}})
        '        End If
        '    End If
        '    '    End If
        'End If

        'InsertCapitalRequired(currentTrade.EntryTime, capitalToBeAdded, capitalToBeReleased, "Placed Or Modify Order")
        ''End If
        ''End Indibar
    End Sub
    Public Function GetSpecificTrades(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType, ByVal tradeStatus As Trade.TradeExecutionStatus) As List(Of Trade)
        Dim ret As List(Of Trade) = Nothing
        If currentMinutePayload IsNot Nothing Then
            Dim tradeDate As Date = currentMinutePayload.PayloadDate.Date
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                ret = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                             Return x.SquareOffType = tradeType AndAlso x.TradeCurrentStatus = tradeStatus
                                                                                         End Function)
            End If
        End If
        Return ret
    End Function
    Public Function GetLastSpecificTrades(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType, ByVal tradeStatus As Trade.TradeExecutionStatus) As Trade
        Dim ret As Trade = Nothing
        Dim specificTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, tradeStatus)
        If specificTrades IsNot Nothing AndAlso specificTrades.Count > 0 Then
            ret = specificTrades.LastOrDefault
        End If
        Return ret
    End Function
    Public Sub ExitTradeByForce(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal exitRemark As String)
        If currentTrade Is Nothing Then Throw New ApplicationException("Supplied trade is nothing, cannot exit")

        If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            CancelTrade(currentTrade, currentPayload, exitRemark)
        Else
            If currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Inprogress Then Throw New ApplicationException("Supplied trade is not active, cannot exit")
            currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                         ExitPrice:=currentPayload.Open,
                                         ExitCondition:=Trade.TradeExitCondition.ForceExit,
                                         ExitRemark:=exitRemark,
                                         TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
            InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin + currentTrade.PLAfterBrokerage, "Exit Trade By Force")
            'Console.WriteLine(currentTrade.ToString)
        End If
    End Sub
    Public Sub ExitStockTradesByForce(ByVal currentPayload As Payload, ByVal tradeType As Trade.TradeType, ByVal exitRemark As String)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentPayload.PayloadDate.Date) AndAlso TradesTaken(currentPayload.PayloadDate.Date).ContainsKey(currentPayload.TradingSymbol) Then
            Dim forceExitTrades As List(Of Trade) = New List(Of Trade)
            Dim inprogessTrades As List(Of Trade) = GetSpecificTrades(currentPayload, tradeType, Trade.TradeExecutionStatus.Inprogress)
            Dim openTrades As List(Of Trade) = GetSpecificTrades(currentPayload, tradeType, Trade.TradeExecutionStatus.Open)
            If inprogessTrades IsNot Nothing Then
                forceExitTrades.AddRange(inprogessTrades)
            End If
            If openTrades IsNot Nothing Then
                forceExitTrades.AddRange(openTrades)
            End If

            If forceExitTrades IsNot Nothing AndAlso forceExitTrades.Count > 0 Then
                For Each forceExitTrade In forceExitTrades
                    ExitTradeByForce(forceExitTrade, currentPayload, exitRemark)
                Next
            End If
        End If
    End Sub
    Public Sub ExitAllTradeByForce(ByVal currentTimeOfExit As Date, ByVal allOneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)), ByVal tradeType As Trade.TradeType, ByVal exitRemark As String)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentTimeOfExit.Date) Then
            Dim allStockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentTimeOfExit.Date)

            'Dim trades As List(Of Trade) = TradesTaken(currentTimeOfExit.Date).Values.Where(Function(x)
            '                                                                                    If x.Find(Function(y)
            '                                                                                                  Return y.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Or y.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress
            '                                                                                              End Function) IsNot Nothing Then
            '                                                                                        Return True
            '                                                                                    Else
            '                                                                                        Return False
            '                                                                                    End If
            '                                                                                End Function)
            If allStockTrades IsNot Nothing AndAlso allStockTrades.Count > 0 Then
                For Each stockName In allStockTrades.Keys
                    Dim candleTime As Date = New Date(currentTimeOfExit.Year, currentTimeOfExit.Month, currentTimeOfExit.Day, currentTimeOfExit.Hour, currentTimeOfExit.Minute, 0)
                    Dim stock As String = Nothing
                    'Indibar
                    If stockName.Contains("FUT") Then
                        stock = stockName.Remove(stockName.Count - 8)
                    Else
                        stock = stockName
                    End If
                    'stock = stockName
                    'End Indibar
                    Dim currentPayload As Payload = Nothing
                    If allOneMinutePayload.ContainsKey(stock) AndAlso allOneMinutePayload(stock).ContainsKey(candleTime) Then
                        currentPayload = allOneMinutePayload(stock)(candleTime).Ticks.FindAll(Function(x)
                                                                                                  Return x.PayloadDate >= currentTimeOfExit
                                                                                              End Function).FirstOrDefault
                        'ElseIf allOneMinutePayload.ContainsKey(stock) Then
                        '    currentPayload = allOneMinutePayload(stock).Where(Function(x)
                        '                                                          Return x.Key >= currentTimeOfExit
                        '                                                      End Function).FirstOrDefault.Value
                    End If
                    If currentPayload Is Nothing Then           'If the current time is more than last available Tick then move to the next available minute
                        currentPayload = allOneMinutePayload(stock).Where(Function(x)
                                                                              Return x.Key >= currentTimeOfExit
                                                                          End Function).FirstOrDefault.Value
                    End If
                    'Start Indibar
                    'ExitStockTradesByForce(currentPayload, tradeType, exitRemark)
                    If currentPayload Is Nothing Then           'If current time payload is not available then pick last available payload
                        Dim lastPayloadTime As Date = allOneMinutePayload(stock).Keys.LastOrDefault
                        currentPayload = allOneMinutePayload(stock)(lastPayloadTime.AddMinutes(-5))
                    End If
                    If currentPayload IsNot Nothing Then
                        ExitStockTradesByForce(currentPayload, tradeType, exitRemark)
                    Else
                        Throw New ApplicationException("Current Payload is NULL in 'ExitAllTradeByForce'")
                    End If
                    'End Indibar
                Next
            End If
        End If

    End Sub
    'Indibar
    'Public Function EnterTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload) As Boolean
    '    Dim ret As Boolean = False
    '    If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then Throw New ApplicationException("Supplied trade is not open, cannot enter")
    '    Dim lastTradingTime As Date = New Date(currentTrade.TradingDate.Year, currentTrade.TradingDate.Month, currentTrade.TradingDate.Day, LastTradeEntryTime.Hours, LastTradeEntryTime.Minutes, LastTradeEntryTime.Seconds)
    '    If currentPayload.PayloadDate < lastTradingTime Then
    '        Dim previousRunningTrades As List(Of Trade) = GetSpecificTrades(currentPayload, currentTrade.SquareOffType, Trade.TradeExecutionStatus.Inprogress)
    '        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
    '            If currentPayload.High >= currentTrade.EntryPrice Then
    '                If previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
    '                    For Each previousRunningTrade In previousRunningTrades
    '                        If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
    '                            ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
    '                        End If
    '                    Next
    '                End If
    '                currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
    '                ret = True
    '            End If
    '        ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
    '            If currentPayload.Low <= currentTrade.EntryPrice Then
    '                If previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
    '                    For Each previousRunningTrade In previousRunningTrades
    '                        If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
    '                            ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
    '                        End If
    '                    Next
    '                End If
    '                currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
    '                ret = True
    '            End If
    '        End If
    '    End If
    '    'If ret Then Console.WriteLine(currentTrade.ToString)
    '    If ret Then SetCurrentLTPForStock(currentPayload, currentPayload, Trade.TradeType.MIS)
    '    Return ret
    'End Function
    Public Function EnterTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload, Optional ByVal reverseSignalExitOnly As Boolean = False) As Boolean
        Dim ret As Boolean = False
        Dim reverseSignalExit As Boolean = False
        If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then Throw New ApplicationException("Supplied trade is not open, cannot enter")
        Dim lastTradingTime As Date = New Date(currentTrade.TradingDate.Year, currentTrade.TradingDate.Month, currentTrade.TradingDate.Day, LastTradeEntryTime.Hours, LastTradeEntryTime.Minutes, LastTradeEntryTime.Seconds)
        If currentPayload.PayloadDate < lastTradingTime Then
            Dim previousRunningTrades As List(Of Trade) = GetSpecificTrades(currentPayload, currentTrade.SquareOffType, Trade.TradeExecutionStatus.Inprogress)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentPayload.High >= currentTrade.EntryPrice Then
                    If previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
                        For Each previousRunningTrade In previousRunningTrades
                            If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
                                reverseSignalExit = True
                            End If
                        Next
                    End If
                    If reverseSignalExitOnly Then
                        If reverseSignalExit OrElse NumberOfTradesPerStockPerDay(currentPayload.PayloadDate, currentPayload.TradingSymbol) >= 1 Then
                            ExitTradeByForce(currentTrade, currentPayload, "Opposite direction trade exited")
                        Else
                            currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                        End If
                    Else
                        currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                    End If
                    ret = True
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentPayload.Low <= currentTrade.EntryPrice Then
                    If previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
                        For Each previousRunningTrade In previousRunningTrades
                            If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
                                reverseSignalExit = True
                            End If
                        Next
                    End If
                    If reverseSignalExitOnly Then
                        If reverseSignalExit OrElse NumberOfTradesPerStockPerDay(currentPayload.PayloadDate, currentPayload.TradingSymbol) >= 1 Then
                            ExitTradeByForce(currentTrade, currentPayload, "Opposite direction trade exited")
                        Else
                            currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                        End If
                    Else
                        currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                    End If
                    ret = True
                End If
            End If
        Else
            'Indibar
            reverseSignalExitOnly = True
            Dim previousRunningTrades As List(Of Trade) = GetSpecificTrades(currentPayload, currentTrade.SquareOffType, Trade.TradeExecutionStatus.Inprogress)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentPayload.High >= currentTrade.EntryPrice Then
                    If previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
                        For Each previousRunningTrade In previousRunningTrades
                            If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
                                reverseSignalExit = True
                            End If
                        Next
                    End If
                    If reverseSignalExitOnly Then
                        If reverseSignalExit OrElse NumberOfTradesPerStockPerDay(currentPayload.PayloadDate, currentPayload.TradingSymbol) >= 1 Then
                            ExitTradeByForce(currentTrade, currentPayload, "Opposite direction trade exited")
                        Else
                            currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                        End If
                    Else
                        currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                    End If
                    ret = True
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentPayload.Low <= currentTrade.EntryPrice Then
                    If previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
                        For Each previousRunningTrade In previousRunningTrades
                            If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
                                reverseSignalExit = True
                            End If
                        Next
                    End If
                    If reverseSignalExitOnly Then
                        If reverseSignalExit OrElse NumberOfTradesPerStockPerDay(currentPayload.PayloadDate, currentPayload.TradingSymbol) >= 1 Then
                            ExitTradeByForce(currentTrade, currentPayload, "Opposite direction trade exited")
                        Else
                            currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                        End If
                    Else
                        currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                    End If
                    ret = True
                End If
            End If
            'end Indibar
        End If
        'If ret Then Console.WriteLine(currentTrade.ToString)
        If ret Then SetCurrentLTPForStock(currentPayload, currentPayload, Trade.TradeType.MIS)
        Return ret
    End Function
    'End Indibar
    Public Sub CancelTrade(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal exitRemark As String)
        If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
            Throw New ApplicationException("Supplied trade is not open, cannot cancel")
        End If
        currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                 ExitCondition:=Trade.TradeExitCondition.Cancelled,
                                 ExitRemark:=exitRemark,
                                 TradeCurrentStatus:=Trade.TradeExecutionStatus.Cancel)

        InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin, "Cancel Trade")
        'Console.WriteLine(currentTrade.ToString)
    End Sub
    Public Function ExitTradeIfPossible(ByVal currentTrade As Trade, ByVal currentPayload As Payload) As Boolean
        Dim ret As Boolean = False
        If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Inprogress Then Throw New ApplicationException("Supplied trade is not active, cannot exit")
        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
            If currentPayload.Low <= currentTrade.PotentialStopLoss Then
                currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                         ExitPrice:=currentPayload.Open,                'Assuming this is tick and OHLC=tickprice
                                         ExitCondition:=Trade.TradeExitCondition.StopLoss,
                                         ExitRemark:="SL hit under normal condition",
                                         TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
                ret = True
            ElseIf currentPayload.High >= currentTrade.PotentialTarget Then
                currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                         ExitPrice:=currentPayload.Open,                'Assuming this is tick and OHLC=tickprice
                                         ExitCondition:=Trade.TradeExitCondition.Target,
                                         ExitRemark:="Target hit under normal condition",
                                         TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
                ret = True
            End If
        ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
            If currentPayload.High >= currentTrade.PotentialStopLoss Then
                currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                         ExitPrice:=currentPayload.Open,            'Assuming this is tick and OHLC=tickprice
                                         ExitCondition:=Trade.TradeExitCondition.StopLoss,
                                         ExitRemark:="SL hit under normal condition",
                                         TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
                ret = True
            ElseIf currentPayload.Low <= currentTrade.PotentialTarget Then
                currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                         ExitPrice:=currentPayload.Open,            'Assuming this is tick and OHLC=tickprice
                                         ExitCondition:=Trade.TradeExitCondition.Target,
                                         ExitRemark:="Target hit under normal condition",
                                         TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
                ret = True
            End If
        End If

        If Not ret And currentPayload.PayloadDate.TimeOfDay.Hours = EODExitTime.Hours And currentPayload.PayloadDate.TimeOfDay.Minutes = EODExitTime.Minutes Then
            currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                         ExitPrice:=currentPayload.Open,            'Assuming this is tick and OHLC=tickprice
                                         ExitCondition:=Trade.TradeExitCondition.EndOfDay,
                                         ExitRemark:="EOD Exit",
                                         TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
            ret = True
        End If

        If ret Then
            InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin + currentTrade.PLAfterBrokerage, "Exit Trade If Possible")
            'Console.WriteLine(currentTrade.ToString)
        End If

        Return ret
    End Function
    Public Shared Function CalculatePL(ByVal stockName As String, ByVal buyPrice As Double, ByVal sellPrice As Double, ByVal quantity As Integer, ByVal lotSize As Integer, ByVal typeOfStock As Trade.TypeOfStock) As Double
        Dim potentialBrokerage As New Calculator.BrokerageAttributes
        Dim calculator As New Calculator.BrokerageCalculator(Nothing)

        Select Case typeOfStock
            Case Trade.TypeOfStock.Cash
                calculator.Intraday_Equity(buyPrice, sellPrice, quantity, potentialBrokerage)
            Case Trade.TypeOfStock.Currency
                calculator.Currency_Futures(buyPrice, sellPrice, quantity / lotSize, potentialBrokerage)
            Case Trade.TypeOfStock.Commodity
                calculator.Commodity_MCX(stockName, buyPrice, sellPrice, quantity / lotSize, potentialBrokerage)
            Case Trade.TypeOfStock.Currency
                Throw New ApplicationException("Not Implemented")
            Case Trade.TypeOfStock.Futures
                calculator.FO_Futures(buyPrice, sellPrice, quantity, potentialBrokerage)
        End Select

        Return potentialBrokerage.NetProfitLoss
    End Function
    Public Shared Function CalculateQuantityFromSL(ByVal stockName As String, ByVal buyPrice As Double, ByVal sellPrice As Double, ByVal NetProfitLossOfTrade As Double, ByVal typeOfStock As Trade.TypeOfStock) As Integer
        Dim potentialBrokerage As Calculator.BrokerageAttributes = Nothing
        Dim calculator As New Calculator.BrokerageCalculator(Nothing)

        Dim quantity As Integer = 1
        Dim previousQuantity As Integer = 1
        For quantity = 1 To Integer.MaxValue
            potentialBrokerage = New Calculator.BrokerageAttributes
            Select Case typeOfStock
                Case Trade.TypeOfStock.Cash
                    calculator.Intraday_Equity(buyPrice, sellPrice, quantity, potentialBrokerage)
                Case Trade.TypeOfStock.Commodity
                    stockName = stockName.Remove(stockName.Count - 8)
                    calculator.Commodity_MCX(stockName, buyPrice, sellPrice, quantity, potentialBrokerage)
                Case Trade.TypeOfStock.Currency
                    Throw New ApplicationException("Not Implemented")
                Case Trade.TypeOfStock.Futures
                    calculator.FO_Futures(buyPrice, sellPrice, quantity, potentialBrokerage)
            End Select

            If NetProfitLossOfTrade > 0 Then
                If potentialBrokerage.NetProfitLoss > NetProfitLossOfTrade Then
                    Exit For
                Else
                    previousQuantity = quantity
                End If
            ElseIf NetProfitLossOfTrade < 0 Then
                If potentialBrokerage.NetProfitLoss < NetProfitLossOfTrade Then
                    Exit For
                Else
                    previousQuantity = quantity
                End If
            End If
        Next
        Return previousQuantity
    End Function
    Public Shared Function CalculateQuantityFromInvestment(ByVal lotSize As Integer, ByVal totalInvestment As Double, ByVal stockPrice As Double, ByVal typeOfStock As Trade.TypeOfStock) As Integer
        Dim quantity As Integer = lotSize
        Dim quantityMultiplier As Double = totalInvestment / (quantity * stockPrice / MarginMultiplier)
        Select Case typeOfStock
            Case Trade.TypeOfStock.Cash
                quantityMultiplier = quantityMultiplier
            Case Trade.TypeOfStock.Commodity
                quantityMultiplier = Math.Floor(quantityMultiplier)
            Case Trade.TypeOfStock.Currency
                Throw New ApplicationException("Not Implemented")
            Case Trade.TypeOfStock.Futures
                quantityMultiplier = Math.Floor(quantityMultiplier)
        End Select
        quantity = Math.Floor(quantity * quantityMultiplier)
        Return quantity
    End Function
    Public Shared Function CalculatorTargetOrStoploss(ByVal coreStockName As String, ByVal entryPrice As Double, ByVal quantity As Integer, ByVal desiredProfitLossOfTrade As Double, ByVal tradeDirection As Trade.TradeExecutionDirection, ByVal typeOfStock As Trade.TypeOfStock) As Double
        Dim potentialBrokerage As Calculator.BrokerageAttributes = Nothing
        Dim calculator As New Calculator.BrokerageCalculator(Nothing)

        Dim exitPrice As Double = entryPrice
        potentialBrokerage = New Calculator.BrokerageAttributes

        If desiredProfitLossOfTrade > 0 Then 'To check target
            While Not potentialBrokerage.NetProfitLoss > desiredProfitLossOfTrade
                If tradeDirection = Trade.TradeExecutionDirection.Buy Then
                    Select Case typeOfStock
                        Case Trade.TypeOfStock.Cash
                            calculator.Intraday_Equity(entryPrice, exitPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Commodity
                            calculator.Commodity_MCX(coreStockName, entryPrice, exitPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case Trade.TypeOfStock.Futures
                            calculator.FO_Futures(entryPrice, exitPrice, quantity, potentialBrokerage)
                    End Select
                    If potentialBrokerage.NetProfitLoss > desiredProfitLossOfTrade Then Exit While
                    exitPrice += TickSize
                ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
                    Select Case typeOfStock
                        Case Trade.TypeOfStock.Cash
                            calculator.Intraday_Equity(exitPrice, entryPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Commodity
                            calculator.Commodity_MCX(coreStockName, exitPrice, entryPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case Trade.TypeOfStock.Futures
                            calculator.FO_Futures(exitPrice, entryPrice, quantity, potentialBrokerage)
                    End Select
                    If potentialBrokerage.NetProfitLoss > desiredProfitLossOfTrade Then Exit While
                    exitPrice -= TickSize
                End If
            End While
        ElseIf desiredProfitLossOfTrade < 0 Then 'To check SL
            While Not potentialBrokerage.NetProfitLoss < desiredProfitLossOfTrade
                If tradeDirection = Trade.TradeExecutionDirection.Buy Then
                    Select Case typeOfStock
                        Case Trade.TypeOfStock.Cash
                            calculator.Intraday_Equity(entryPrice, exitPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Commodity
                            calculator.Commodity_MCX(coreStockName, entryPrice, exitPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case Trade.TypeOfStock.Futures
                            calculator.FO_Futures(entryPrice, exitPrice, quantity, potentialBrokerage)
                    End Select
                    If potentialBrokerage.NetProfitLoss < desiredProfitLossOfTrade Then Exit While
                    exitPrice -= TickSize
                ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
                    Select Case typeOfStock
                        Case Trade.TypeOfStock.Cash
                            calculator.Intraday_Equity(exitPrice, entryPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Commodity
                            calculator.Commodity_MCX(coreStockName, exitPrice, entryPrice, quantity, potentialBrokerage)
                        Case Trade.TypeOfStock.Currency
                            Throw New ApplicationException("Not Implemented")
                        Case Trade.TypeOfStock.Futures
                            calculator.FO_Futures(exitPrice, entryPrice, quantity, potentialBrokerage)
                    End Select
                    If potentialBrokerage.NetProfitLoss < desiredProfitLossOfTrade Then Exit While
                    exitPrice += TickSize
                End If
            End While
        End If

        Return Math.Round(exitPrice, 2)
    End Function
    Public Sub MoveStopLoss(ByVal currentTime As Date, ByVal currentTrade As Trade, ByVal price As Double, ByVal slRemark As String)
        If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Inprogress Then Throw New ApplicationException("Supplied trade is not active, cannot move SL")
        currentTrade.UpdateTrade(PotentialStopLoss:=price, SLRemark:=slRemark, StoplossSetTime:=currentTime)
        'Console.WriteLine(currentTrade.ToString)
    End Sub

    ''' <summary>
    ''' If in a lower timeframe, we want to know the previous higher timeframe time, then we need to use this
    ''' It will take the lower timeframe
    ''' </summary>
    Public Function GetPreviousXMinuteCandleTime(ByVal lowerTFTime As Date, ByVal higherTFPayload As Dictionary(Of Date, Payload), ByVal higherTF As Integer) As Date
        Dim ret As Date = Nothing

        If higherTFPayload IsNot Nothing AndAlso higherTFPayload.Count > 0 Then
            ret = higherTFPayload.Keys.LastOrDefault(Function(x)
                                                         Return x <= lowerTFTime.AddMinutes(-higherTF)
                                                     End Function)
        End If
        Return ret
    End Function
    Public Function GetCurrentXMinuteCandleTime(ByVal lowerTFTime As Date, ByVal higherTFPayload As Dictionary(Of Date, Payload)) As Date
        Dim ret As Date = Nothing

        If higherTFPayload IsNot Nothing AndAlso higherTFPayload.Count > 0 Then
            ret = higherTFPayload.Keys.LastOrDefault(Function(x)
                                                         Return x <= lowerTFTime
                                                     End Function)
        End If
        Return ret
    End Function
    Public Function GetDaysPLAfterBrokerageForThePortfolio(ByVal currentDate As Date) As Double
        Dim ret As Double = Nothing
        currentDate = currentDate.Date
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate) Then
            ret = TradesTaken(currentDate).Sum(Function(x)
                                                   Return x.Value.Sum(Function(y)
                                                                          If y.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Or y.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                                                              Return y.PLAfterBrokerage
                                                                          Else
                                                                              Return 0
                                                                          End If
                                                                      End Function)
                                               End Function)
        End If
        Return ret
    End Function

    ''' <summary>
    ''' Calculate SL for fixed percentage
    ''' </summary>
    ''' <param name="entryPrice"></param>
    ''' <param name="currentLTP"></param>
    ''' <param name="tradeDirection"></param>
    ''' <param name="slab"></param>
    ''' <param name="breakevenPoint"></param>
    ''' <param name="movementRemarks"></param>
    ''' <returns></returns>
    Public Function CalculateMathematicalTrailingSL(ByVal entryPrice As Double, ByVal currentLTP As Double, ByVal tradeDirection As Trade.TradeExecutionDirection, ByVal slab As Double, ByVal breakevenPoint As Double, ByRef movementRemarks As String) As Double
        Dim ret As Double = Nothing
        Dim movement As Double = Nothing
        Dim gain As Double = ((currentLTP / entryPrice) - 1)

        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
            ret = Double.MinValue
            If gain >= slab And gain < slab * 1.5 Then
                ret = entryPrice + breakevenPoint * entryPrice
                movementRemarks = "Move to Breakeven"
            ElseIf gain >= slab * 1.5 Then
                movement = (CInt(gain / (slab / 2)) - 2) * (slab / 2)
                ret = entryPrice + entryPrice * movement
                movementRemarks = String.Format("Onward Move({0})", movement / (slab / 2))
            End If
        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
            ret = Double.MaxValue
            slab = slab * -1
            If gain <= slab And gain > slab * 1.5 Then
                ret = entryPrice - breakevenPoint * entryPrice
                movementRemarks = "Move to Breakeven"
            ElseIf gain <= slab * 1.5 Then
                movement = (CInt(gain / (slab / 2)) - 2) * (slab / 2)
                ret = entryPrice - entryPrice * (movement * -1)
                movementRemarks = String.Format("Onward Move({0})", movement * -1 / (slab / 2))
            End If
        End If
        Return ret
    End Function
    'Start Indibar
    ''' <summary>
    ''' Calculate SL for fixed point
    ''' </summary>
    ''' <param name="currentTrade"></param>
    ''' <param name="currentLTP"></param>
    ''' <param name="breakevenPoint"></param>
    ''' <param name="movementRemarks"></param>
    ''' <returns></returns>
    Public Function CalculateMathematicalTrailingSL(ByVal currentTrade As Trade, ByVal currentLTP As Double, ByVal breakevenPoint As Double, ByRef movementRemarks As String) As Double
        ' Original
        Dim ret As Double = Nothing
        Dim movement As Double = Nothing
        Dim movementMultiplier As Integer = 0
        Dim entryPrice As Decimal = currentTrade.EntryPrice
        Dim squareOffValue As Decimal = currentTrade.SquareOffValue
        Dim tradeDirection As Trade.TradeExecutionDirection = currentTrade.EntryDirection
        Dim gain As Double = currentLTP - entryPrice
        'Dim slab As Double = squareOffValue / 2
        Dim slab As Double = squareOffValue

        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
            ret = entryPrice
            If gain >= squareOffValue And gain < (squareOffValue + slab) Then
                ret = entryPrice + breakevenPoint
                movementRemarks = "Move to Breakeven"
            ElseIf gain >= (squareOffValue + slab) Then
                movementMultiplier = Math.Floor(gain / slab)
                movement = (movementMultiplier * slab) - squareOffValue
                ret = entryPrice + movement
                movementRemarks = String.Format("Onward Move({0})", movement / slab)
            End If
        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
            ret = entryPrice
            If gain <= -squareOffValue And gain > -(squareOffValue + slab) Then
                ret = entryPrice - breakevenPoint
                movementRemarks = "Move to Breakeven"
            ElseIf gain <= -(squareOffValue + slab) Then
                movementMultiplier = Math.Floor(Math.Abs(gain / slab)) * -1
                movement = (movementMultiplier * slab) + squareOffValue
                ret = entryPrice - (movement * -1)
                movementRemarks = String.Format("Onward Move({0})", movement * -1 / slab)
            End If
        End If
        Return ret
    End Function
    'Public Function CalculateMathematicalTrailingSL(ByVal currentTrade As Trade, ByVal currentLTP As Double, ByVal breakevenPoint As Double, ByRef movementRemarks As String) As Double
    '    'Duplicate
    '    Dim ret As Double = Nothing
    '    Dim movement As Double = Nothing
    '    Dim movementMultiplier As Integer = 0
    '    Dim entryPrice As Decimal = currentTrade.EntryPrice
    '    Dim squareOffValue As Decimal = currentTrade.SquareOffValue
    '    Dim tradeDirection As Trade.TradeExecutionDirection = currentTrade.EntryDirection
    '    Dim gain As Double = Math.Abs(currentLTP - entryPrice)
    '    Dim gainPercentage As Double = Math.Round(gain * 100 / entryPrice, 2)
    '    'Dim slab As Double = squareOffValue / 2
    '    Dim slab As Double = squareOffValue / 2

    '    ret = entryPrice
    '    If gainPercentage >= 1.5 Then
    '        movementMultiplier = Math.Floor(gainPercentage * 10 / 5)
    '        movement = (movementMultiplier - 2) * 0.5
    '        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
    '            ret = entryPrice + ConvertFloorCeling(entryPrice * movement / 100, TickSize, RoundOfType.Floor)
    '        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
    '            ret = entryPrice - ConvertFloorCeling(entryPrice * movement / 100, TickSize, RoundOfType.Floor)
    '        End If
    '        movementRemarks = String.Format("Onward Move({0})", movement / 0.5)
    '    End If
    '    Return ret
    'End Function

    'Private _SLmultiplier As Integer = 1
    'Private _lastTradeEntryTime As Date = Date.MinValue
    'Public Function CalculateLogicalTrailingSL(ByVal currentTrade As Trade, ByVal currentLTP As Double, ByVal time As Date, ByVal swingHigh As Double, ByVal swingLow As Double, ByRef movementRemarks As String) As Double
    '    Dim ret As Double = Nothing
    '    If _lastTradeEntryTime <> Date.MinValue AndAlso _lastTradeEntryTime <> currentTrade.EntryTime Then
    '        _SLmultiplier = 1
    '    End If
    '    _lastTradeEntryTime = currentTrade.EntryTime
    '    Dim entryPrice As Decimal = currentTrade.EntryPrice
    '    Dim squareOffValue As Decimal = currentTrade.SquareOffValue
    '    Dim tradeDirection As Trade.TradeExecutionDirection = currentTrade.EntryDirection
    '    Dim gain As Double = Math.Abs(currentLTP - entryPrice)

    '    ret = entryPrice
    '    If gain >= squareOffValue * _SLmultiplier Then
    '        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
    '            ret = swingLow - CalculateBuffer(swingLow, RoundOfType.Floor)
    '        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
    '            ret = swingHigh + CalculateBuffer(swingHigh, RoundOfType.Floor)
    '        End If
    '        movementRemarks = String.Format("Onward Move ({0},{1},{2})", _SLmultiplier, time, ret)
    '        If currentTrade.PotentialStopLoss <> ret Then
    '            _SLmultiplier += 1
    '        End If
    '    End If
    '    Return ret
    'End Function
    Public Function CalculateLogicalTrailingSL(ByVal currentTrade As Trade, ByVal currentLTP As Double, ByVal currentLTPTime As Date, ByRef movementRemarks As String) As Double
        Dim ret As Double = Nothing
        Dim entryPrice As Decimal = currentTrade.EntryPrice
        Dim squareOffValue As Decimal = currentTrade.SquareOffValue
        Dim tradeDirection As Trade.TradeExecutionDirection = currentTrade.EntryDirection
        Dim gain As Double = Math.Abs(currentLTP - entryPrice)
        Dim targetToAchive As Decimal = Decimal.MinValue
        ret = currentTrade.PotentialStopLoss
        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
            targetToAchive = (currentTrade.PotentialTarget - currentTrade.EntryPrice) * 2 / 3
            If gain >= targetToAchive Then
                Dim breakevenPoints As Decimal = GetBreakevenPoints(currentTrade.TradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection)
                ret = entryPrice + breakevenPoints
                movementRemarks = String.Format("Move to breakeven point {0} at {1}", breakevenPoints, currentLTPTime.ToString("HH:mm:ss"))
            End If
        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
            targetToAchive = (currentTrade.EntryPrice - currentTrade.PotentialTarget) * 2 / 3
            If gain >= targetToAchive Then
                Dim breakevenPoints As Decimal = GetBreakevenPoints(currentTrade.TradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection)
                ret = entryPrice - breakevenPoints
                movementRemarks = String.Format("Move to breakeven point {0} at {1}", breakevenPoints, currentLTPTime.ToString("HH:mm:ss"))
            End If
        End If
        Return ret
    End Function
    'End Indibar
    Public Function CalculateRevisedMathematicalTrailingSL(ByVal entryPrice As Double, ByVal currentLTP As Double, ByVal tradeDirection As Trade.TradeExecutionDirection, ByVal slab As Double, ByVal breakevenPoint As Double, ByRef movementRemarks As String) As Double
        Dim ret As Double = Nothing
        Dim movement As Double = Nothing
        Dim gain As Double = ((currentLTP / entryPrice) - 1)

        If tradeDirection = Trade.TradeExecutionDirection.Buy Then
            ret = Double.MinValue
            If gain >= slab And gain < slab * 1.1 Then
                ret = entryPrice + breakevenPoint * entryPrice
                movementRemarks = "Move to Breakeven"
            ElseIf gain >= slab * 1.1 Then
                gain = gain * 100
                movement = (gain - (slab * 100) - If((gain * 100) Mod 10 = 10, 0, (gain * 100) Mod 10) / 100) / 100
                ret = entryPrice + entryPrice * movement
                movementRemarks = String.Format("Onward Move({0})", movement / (slab / 10))
            End If
        ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
            ret = Double.MaxValue
            slab = slab * -1
            If gain <= slab And gain > slab * 1.1 Then
                ret = entryPrice - breakevenPoint * entryPrice
                movementRemarks = "Move to Breakeven"
            ElseIf gain <= slab * 1.1 Then
                gain = gain * 100
                movement = (gain - (slab * 100) - If((gain * 100) Mod -10 = -10, 0, (gain * 100) Mod -10) / 100) / 100
                ret = entryPrice - entryPrice * (movement * -1)
                movementRemarks = String.Format("Onward Move({0})", movement * -1 / (slab / 10))
            End If
        End If
        Return ret
    End Function

    Public Shared Function GetBreakevenPoints(ByVal tradingSymbol As String, ByVal entryPrice As Decimal, ByVal quantity As Integer, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = TickSize
        If direction = Trade.TradeExecutionDirection.Buy Then
            For exitPrice As Decimal = entryPrice To Decimal.MaxValue Step ret
                Dim pl As Decimal = CalculatePL(tradingSymbol, entryPrice, exitPrice, quantity, 1, Trade.TypeOfStock.Futures)
                If pl >= 0 Then
                    ret = Math.Round(exitPrice - entryPrice, 2)
                    Exit For
                End If
            Next
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            For exitPrice As Decimal = entryPrice To Decimal.MinValue Step ret * -1
                Dim pl As Decimal = CalculatePL(tradingSymbol, exitPrice, entryPrice, quantity, 1, Trade.TypeOfStock.Futures)
                If pl >= 0 Then
                    ret = Math.Round(entryPrice - exitPrice, 2)
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Public Function IsStockAnyPLReachedForTheDay(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType) As Boolean
        Dim ret As Boolean = False
        Dim lastTrade As Trade = GetLastSpecificTrades(currentMinutePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Close)
        Dim currentStockMTM As Double = Nothing
        Dim probableExitPrice As Decimal = Nothing
        If lastTrade IsNot Nothing Then
            currentStockMTM = StockPLBeforeBrokerage(currentMinutePayload.PayloadDate, currentMinutePayload.TradingSymbol)
            probableExitPrice = currentStockMTM / lastTrade.Quantity + lastTrade.EntryPrice
            If Math.Round((((probableExitPrice / lastTrade.EntryPrice) - 1) * 100), 1) >= PerStockMaxProfitPercentage OrElse
            Math.Round((((probableExitPrice / lastTrade.EntryPrice) - 1) * 100), 1) <= PerStockMaxLossPercentage Then
                ret = True
            End If
        End If
        Return ret
    End Function
    Public Sub SetCurrentLTPForStock(ByVal currentMinutePayload As Payload, ByVal currentTickPayload As Payload, ByVal tradeType As Trade.TradeType)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentMinutePayload.PayloadDate.Date) AndAlso TradesTaken(currentMinutePayload.PayloadDate.Date).ContainsKey(currentMinutePayload.TradingSymbol) Then
            Dim ltpUpdateTrades As List(Of Trade) = New List(Of Trade)
            Dim inprogessTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Inprogress)
            Dim openTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Open)
            If inprogessTrades IsNot Nothing Then
                ltpUpdateTrades.AddRange(inprogessTrades)
            End If
            If openTrades IsNot Nothing Then
                ltpUpdateTrades.AddRange(openTrades)
            End If

            If ltpUpdateTrades IsNot Nothing AndAlso ltpUpdateTrades.Count > 0 Then
                For Each ltpUpdateTrade In ltpUpdateTrades
                    ltpUpdateTrade.CurrentLTP = currentTickPayload.Open  'Assuming OHCL of tick is same
                Next
            End If
        End If
    End Sub
    Public Function IsStockTradeUpdatedInCurrentMinute(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType) As Boolean
        Dim ret As Boolean = False
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentMinutePayload.PayloadDate.Date) AndAlso TradesTaken(currentMinutePayload.PayloadDate.Date).ContainsKey(currentMinutePayload.TradingSymbol) Then
            Dim updateTrades As List(Of Trade) = New List(Of Trade)
            Dim inprogessTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Inprogress)
            Dim openTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Open)
            Dim closeTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Close)
            If inprogessTrades IsNot Nothing Then
                updateTrades.AddRange(inprogessTrades)
            End If
            If openTrades IsNot Nothing Then
                updateTrades.AddRange(openTrades)
            End If
            If closeTrades IsNot Nothing Then
                updateTrades.AddRange(closeTrades)
            End If

            If updateTrades IsNot Nothing AndAlso updateTrades.Count > 0 Then
                For Each updateTrade In updateTrades
                    Dim chkeckingTime As Date = Nothing
                    chkeckingTime = New Date(updateTrade.TradeUpdateTimeStamp.Year, updateTrade.TradeUpdateTimeStamp.Month, updateTrade.TradeUpdateTimeStamp.Day, updateTrade.TradeUpdateTimeStamp.Hour, updateTrade.TradeUpdateTimeStamp.Minute, 0)
                    If chkeckingTime = currentMinutePayload.PayloadDate Then
                        ret = True
                        Exit For
                    End If
                Next
            End If
        End If
        Return ret
    End Function
    Public Sub SetOverallDrawUpDrawDownForTheDay(ByVal currentDate As Date)
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
            Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
            If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                For Each stock In stockTrades.Keys
                    Dim trades As List(Of Trade) = TradesTaken(currentDate.Date)(stock)
                    For Each runningTrade In trades
                        runningTrade.OverAllMaxDrawDownPL = Me.AllMaxDrawDownPLAfterBrokerage(currentDate)
                        runningTrade.OverAllMaxDrawUpPL = Me.AllMaxDrawUpPLAfterBrokerage(currentDate)
                    Next
                Next
            End If
        End If
        _AllMaxDrawDownPLAfterBrokerage = Double.MaxValue
        _AllMaxDrawUpPLAfterBrokerage = Double.MinValue
    End Sub
    Public Function IsAnyTradeOfTheStockTargetReached(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType) As Boolean
        Dim ret As Boolean = False
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentMinutePayload.PayloadDate.Date) AndAlso TradesTaken(currentMinutePayload.PayloadDate.Date).ContainsKey(currentMinutePayload.TradingSymbol) Then
            Dim completeTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Close)
            If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
                Dim targetTrades As List(Of Trade) = completeTrades.FindAll(Function(x)
                                                                                Return x.ExitCondition = Trade.TradeExitCondition.Target
                                                                            End Function)
                If targetTrades IsNot Nothing AndAlso targetTrades.Count > 0 Then
                    ret = True
                End If
            End If
        End If
        Return ret
    End Function
    Public Function GetLastExitTradeOfTheStock(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TradeType) As Trade
        Dim ret As Trade = Nothing
        If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentMinutePayload.PayloadDate.Date) AndAlso TradesTaken(currentMinutePayload.PayloadDate.Date).ContainsKey(currentMinutePayload.TradingSymbol) Then
            Dim completeTrades As List(Of Trade) = GetSpecificTrades(currentMinutePayload, tradeType, Trade.TradeExecutionStatus.Close)
            If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
                ret = completeTrades.OrderBy(Function(x)
                                                 Return x.ExitTime
                                             End Function).LastOrDefault
            End If
        End If
        Return ret
    End Function
    Public Function IsAnyCandleClosesAboveOrBelow(ByVal currentMinute As Date, ByVal lastExitTime As Date, ByVal totalXMinutePayload As Dictionary(Of Date, Payload), ByVal potentialEntryTrade As Trade) As Boolean
        Dim ret As Boolean = False
        If totalXMinutePayload IsNot Nothing AndAlso totalXMinutePayload.Count > 0 Then
            Dim payloadsToCheck As IEnumerable(Of KeyValuePair(Of Date, Payload)) = totalXMinutePayload.Where(Function(x)
                                                                                                                  Return x.Key >= lastExitTime AndAlso
                                                                                                                  x.Key < currentMinute
                                                                                                              End Function)
            If payloadsToCheck IsNot Nothing AndAlso payloadsToCheck.Count > 0 Then
                If potentialEntryTrade IsNot Nothing Then
                    If potentialEntryTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        For Each runningPayload In payloadsToCheck
                            If runningPayload.Value.Close <= potentialEntryTrade.EntryPrice Then
                                ret = True
                                Exit For
                            End If
                        Next
                    ElseIf potentialEntryTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        For Each runningPayload In payloadsToCheck
                            If runningPayload.Value.Close >= potentialEntryTrade.EntryPrice Then
                                ret = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If
        Return ret
    End Function
    Public MustOverride Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
#End Region

#Region "Private Functions"
    Private Sub InsertCapitalRequired(ByVal currentDate As Date, ByVal capitalToBeInserted As Double, ByVal capitalToBeReleased As Double, ByVal remarks As String)
        Dim capitalRequired As Capital = New Capital
        Dim capitalToBeAdded As Double = 0

        If CapitalMovement IsNot Nothing AndAlso CapitalMovement.Count > 0 AndAlso CapitalMovement.ContainsKey(currentDate.Date) Then
            Dim capitalList As List(Of Capital) = CapitalMovement(currentDate.Date)
            capitalToBeAdded = capitalList.LastOrDefault.RunningCapital
        Else
            If CapitalMovement Is Nothing Then CapitalMovement = New Dictionary(Of Date, List(Of Capital))
            CapitalMovement.Add(currentDate.Date, New List(Of Capital))
        End If

        With capitalRequired
            .TradingDate = currentDate.Date
            .CapitalExhaustedDateTime = currentDate
            .RunningCapital = If(capitalToBeInserted = 0, 0, capitalToBeInserted + capitalToBeAdded - capitalToBeReleased)
            .CapitalReleased = capitalToBeReleased
            .Remarks = remarks
        End With
        CapitalMovement(currentDate.Date).Add(capitalRequired)
    End Sub
#End Region

#Region "Deserialize"
    Public Function Deserialize(ByVal inputFilePath As String, ByVal retryCounter As Integer) As Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
        Dim ret As Dictionary(Of Date, Dictionary(Of String, List(Of Trade))) = Nothing
        Using stream As New FileStream(inputFilePath, FileMode.Open)
            Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            Dim counter As Integer = 0
            Dim totalSize As Long = 0
            While stream.Position <> stream.Length
                If totalSize <> 0 Then OnHeartbeat(String.Format("Deserializing Trades collection {0}/{1} Retry Counter:{2}", counter, totalSize, retryCounter))
                Dim temp As KeyValuePair(Of Date, Dictionary(Of String, List(Of Trade))) = Nothing
                Dim tempData As Dictionary(Of Date, Dictionary(Of String, List(Of Trade))) = binaryFormatter.Deserialize(stream)
                For Each runningDate In tempData.Keys
                    Dim stockData As Dictionary(Of String, List(Of Trade)) = Nothing
                    For Each stock In tempData(runningDate).Keys
                        Dim tradeList As List(Of Trade) = tempData(runningDate)(stock).FindAll(Function(x)
                                                                                                   Return x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                               End Function)
                        If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                            If stockData Is Nothing Then stockData = New Dictionary(Of String, List(Of Trade))
                            stockData.Add(stock, tradeList)
                        End If
                    Next
                    If stockData IsNot Nothing AndAlso stockData.Count > 0 Then
                        temp = New KeyValuePair(Of Date, Dictionary(Of String, List(Of Trade)))(runningDate, stockData)
                        If ret Is Nothing Then
                            ret = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
                            totalSize = Math.Ceiling(stream.Length / stream.Position)
                        End If
                        ret.Add(temp.Key, temp.Value)
                    End If
                Next
                counter += 1
            End While
            Return ret
        End Using
    End Function
#End Region

#Region "Print to excel direct"
    Public Overridable Sub PrintArrayToExcel(ByVal fileName As String, Optional ByVal tradesFilename As String = Nothing, Optional ByVal capitalFileName As String = Nothing)
        For retryCounter As Integer = 1 To 20 Step 1
            Try
                Dim allTradesData As Dictionary(Of Date, Dictionary(Of String, List(Of Trade))) = Nothing
                Dim allCapitalData As Dictionary(Of Date, List(Of Capital)) = Nothing
                If tradesFilename IsNot Nothing AndAlso capitalFileName IsNot Nothing Then
                    OnHeartbeat("Deserializing Trades collections")
                    allTradesData = Deserialize(tradesFilename, retryCounter)
                    OnHeartbeat("Deserializing Capital collections")
                    allCapitalData = Utilities.Strings.DeserializeToCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName)
                Else
                    allTradesData = TradesTaken
                    allCapitalData = CapitalMovement
                End If
                If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 Then
                    Dim cts As New CancellationTokenSource
                    Dim totalTrades As Integer = allTradesData.Values.Sum(Function(x)
                                                                              Return x.Values.Sum(Function(y)
                                                                                                      Return y.FindAll(Function(z)
                                                                                                                           Return z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                                       End Function).Count
                                                                                                  End Function)
                                                                          End Function)

                    Dim totalPositiveTrades As Integer = allTradesData.Values.Sum(Function(x)
                                                                                      Return x.Values.Sum(Function(y)
                                                                                                              Return y.FindAll(Function(z)
                                                                                                                                   Return z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                  z.PLAfterBrokerage > 0
                                                                                                                               End Function).Count
                                                                                                          End Function)
                                                                                  End Function)

                    Dim sumOfPositiveTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                      Return x.Values.Sum(Function(y)
                                                                                                              Return y.Sum(Function(z)
                                                                                                                               If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                z.PLAfterBrokerage > 0 Then
                                                                                                                                   Return z.PLAfterBrokerage
                                                                                                                               Else
                                                                                                                                   Return 0
                                                                                                                               End If
                                                                                                                           End Function)
                                                                                                          End Function)
                                                                                  End Function)

                    Dim sumOfNegativeTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                      Return x.Values.Sum(Function(y)
                                                                                                              Return y.Sum(Function(z)
                                                                                                                               If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                z.PLAfterBrokerage <= 0 Then
                                                                                                                                   Return z.PLAfterBrokerage
                                                                                                                               Else
                                                                                                                                   Return 0
                                                                                                                               End If
                                                                                                                           End Function)
                                                                                                          End Function)
                                                                                  End Function)

                    Dim totalDurationInTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                        Return x.Values.Sum(Function(y)
                                                                                                                Return y.Sum(Function(z)
                                                                                                                                 If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                                                                                                     Return z.DurationOfTrade.TotalMinutes
                                                                                                                                 Else
                                                                                                                                     Return 0
                                                                                                                                 End If
                                                                                                                             End Function)
                                                                                                            End Function)
                                                                                    End Function)

                    Dim totalDurationInPositiveTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                                Return x.Values.Sum(Function(y)
                                                                                                                        Return y.Sum(Function(z)
                                                                                                                                         If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                          z.PLAfterBrokerage > 0 Then
                                                                                                                                             Return z.DurationOfTrade.TotalMinutes
                                                                                                                                         Else
                                                                                                                                             Return 0
                                                                                                                                         End If
                                                                                                                                     End Function)
                                                                                                                    End Function)
                                                                                            End Function)

                    Dim totalDurationInNegativeTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                                Return x.Values.Sum(Function(y)
                                                                                                                        Return y.Sum(Function(z)
                                                                                                                                         If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                          z.PLAfterBrokerage <= 0 Then
                                                                                                                                             Return z.DurationOfTrade.TotalMinutes
                                                                                                                                         Else
                                                                                                                                             Return 0
                                                                                                                                         End If
                                                                                                                                     End Function)
                                                                                                                    End Function)
                                                                                            End Function)

                    Dim largestWinningTrade As Decimal = allTradesData.Values.Max(Function(x)
                                                                                      Return x.Values.Max(Function(y)
                                                                                                              Return y.Max(Function(z)
                                                                                                                               If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                                                                                                   Return z.PLAfterBrokerage
                                                                                                                               Else
                                                                                                                                   Return Decimal.MinValue
                                                                                                                               End If
                                                                                                                           End Function)
                                                                                                          End Function)
                                                                                  End Function)

                    Dim largestLosingTrade As Decimal = allTradesData.Values.Min(Function(x)
                                                                                     Return x.Values.Min(Function(y)
                                                                                                             Return y.Min(Function(z)
                                                                                                                              If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                                                                                                  Return z.PLAfterBrokerage
                                                                                                                              Else
                                                                                                                                  Return Decimal.MaxValue
                                                                                                                              End If
                                                                                                                          End Function)
                                                                                                         End Function)
                                                                                 End Function)

                    Dim winRatio As Decimal = Math.Round((totalPositiveTrades / totalTrades) * 100, 2)
                    Dim riskReward As Decimal = 0
                    If totalPositiveTrades <> 0 AndAlso (totalTrades - totalPositiveTrades) <> 0 Then
                        riskReward = Math.Round(Math.Abs((sumOfPositiveTrades / totalPositiveTrades) / (sumOfNegativeTrades / (totalTrades - totalPositiveTrades))), 2)
                    End If

                    Dim strategyOutputData As StrategyOutput = New StrategyOutput
                    With strategyOutputData
                        .WinRatio = winRatio
                        .NetProfit = sumOfPositiveTrades + sumOfNegativeTrades
                        .GrossProfit = sumOfPositiveTrades
                        .GrossLoss = sumOfNegativeTrades
                        .TotalTrades = totalTrades
                        .TotalWinningTrades = totalPositiveTrades
                        .TotalLosingTrades = totalTrades - totalPositiveTrades
                        .AverageTrades = (sumOfPositiveTrades + sumOfNegativeTrades) / totalTrades
                        .AverageWinningTrades = If(totalPositiveTrades <> 0, sumOfPositiveTrades / totalPositiveTrades, 0)
                        .AverageLosingTrades = If((totalTrades - totalPositiveTrades) <> 0, sumOfNegativeTrades / (totalTrades - totalPositiveTrades), 0)
                        .RiskReward = riskReward
                        .LargestWinningTrade = largestWinningTrade
                        .LargestLosingTrade = largestLosingTrade
                        .AverageDurationInTrades = totalDurationInTrades / totalTrades
                        .AverageDurationInWinningTrades = If(totalPositiveTrades <> 0, totalDurationInPositiveTrades / totalPositiveTrades, 0)
                        .AverageDurationInLosingTrades = If((totalTrades - totalPositiveTrades) <> 0, totalDurationInNegativeTrades / (totalTrades - totalPositiveTrades), 0)
                    End With

                    fileName = String.Format("WR {0}%,RR {1},PL {2},{3}.xlsx", winRatio, riskReward, Math.Round(strategyOutputData.NetProfit, 0), fileName)
                    Dim filepath As String = Path.Combine(My.Application.Info.DirectoryPath, "BackTest Output", fileName)
                    If File.Exists(filepath) Then File.Delete(filepath)

                    OnHeartbeat("Opening Excel.....")
                    Using excelWriter As New ExcelHelper(filepath, ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, cts)
                        excelWriter.CreateNewSheet("Capital")
                        excelWriter.CreateNewSheet("Data")
                        excelWriter.SetActiveSheet("Data")

                        Dim rowCtr As Integer = 0
                        Dim colCtr As Integer = 0

                        Dim rowCount As Integer = 0
                        If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 Then
                            rowCount = allTradesData.Sum(Function(x)
                                                             Dim stockTrades = x.Value
                                                             Return stockTrades.Sum(Function(y)
                                                                                        Dim trades = y.Value.FindAll((Function(z)
                                                                                                                          Return z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                                      End Function))
                                                                                        Return trades.Count
                                                                                    End Function)
                                                         End Function)
                        End If

                        Dim mainRawData(rowCount, 0) As Object

                        If rowCtr = 0 Then
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Trading Date"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Trading Symbol"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Capital Required With Margin"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Square Off Type"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Entry Direction"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Buy Price"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Sell Price"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Quantity"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Entry Time"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Exit Time"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Duration Of Trade"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Entry Condition"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Entry Remark"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Exit Condition"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Exit Remark"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "SL Remark"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Target Remark"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "PL Point"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "PL Before Brokerage"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "PL After Brokerage"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Maximum DrawUp Point"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Maximum DrawDown Point"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Maximum Draw Up PL"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Maximum Draw Down PL"
                            colCtr += 1
                            'If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            'mainRawData(rowCtr, colCtr) = "Signal Candle Time"
                            'colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Month"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Entry Buffer"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "SL Buffer"
                            colCtr += 1
                            'If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            'mainRawData(rowCtr, colCtr) = "SquareOffValue"
                            'colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Exit Before PL"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Overall Draw Up PL for the day"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Overall Draw Down PL for the day"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Supporting1"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Supporting2"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Supporting3"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Supporting4"
                            colCtr += 1
                            If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                            mainRawData(rowCtr, colCtr) = "Supporting5"

                            rowCtr += 1
                        End If
                        Dim dateCtr As Integer = 0
                        For Each tempkeys In allTradesData.Keys
                            dateCtr += 1
                            OnHeartbeat(String.Format("Excel printing for Date: {0} [{1} of {2}]", tempkeys.Date.ToShortDateString, dateCtr, allTradesData.Count))
                            Dim stockTrades As Dictionary(Of String, List(Of Trade)) = allTradesData(tempkeys)
                            If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                                Dim stockCtr As Integer = 0
                                For Each stock In stockTrades.Keys
                                    stockCtr += 1
                                    If stockTrades.ContainsKey(stock) Then
                                        Dim tradeList As List(Of Trade) = stockTrades(stock).FindAll(Function(x)
                                                                                                         Return x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                     End Function)
                                        Dim tradeCtr As Integer = 0
                                        If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                                            For Each tradeTaken In tradeList.OrderBy(Function(x)
                                                                                         Return x.EntryTime
                                                                                     End Function)
                                                tradeCtr += 1
                                                OnHeartbeat(String.Format("Excel printing: {0} of {1}", tradeCtr, tradeList.Count))
                                                colCtr = 0

                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.TradingDate.ToString("dd-MMM-yyyy")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.TradingSymbol
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.CapitalRequiredWithMargin
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.SquareOffType.ToString
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryDirection.ToString
                                                colCtr += 1
                                                If tradeTaken.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 4)
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 4)
                                                    colCtr += 1
                                                ElseIf tradeTaken.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 4)
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 4)
                                                    colCtr += 1
                                                End If
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Quantity
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryTime.ToString("HH:mm:ss")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ExitTime.ToString("HH:mm:ss")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.DurationOfTrade.TotalMinutes, 4)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryCondition.ToString
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryRemark
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ExitCondition.ToString
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.ExitRemark
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.SLRemark
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.TargetRemark
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.PLPoint, 4)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.PLBeforeBrokerage, 4)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.PLAfterBrokerage
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = If(tradeTaken.EntryDirection = Trade.TradeExecutionDirection.Buy, tradeTaken.MaximumDrawUp - tradeTaken.EntryPrice, tradeTaken.EntryPrice - tradeTaken.MaximumDrawUp)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = If(tradeTaken.EntryDirection = Trade.TradeExecutionDirection.Buy, tradeTaken.MaximumDrawDown - tradeTaken.EntryPrice, tradeTaken.EntryPrice - tradeTaken.MaximumDrawDown)
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.MaximumDrawUpPL
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.MaximumDrawDownPL
                                                colCtr += 1
                                                'If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                'mainRawData(rowCtr, colCtr) = tradeTaken.SignalCandle.PayloadDate.ToString("HH:mm:ss")
                                                'colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = String.Format("{0}-{1}", tradeTaken.TradingDate.ToString("yyyy"), tradeTaken.TradingDate.ToString("MM"))
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.EntryBuffer
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.StoplossBuffer
                                                colCtr += 1
                                                'If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                'mainRawData(rowCtr, colCtr) = tradeTaken.SquareOffValue
                                                'colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = If(Math.Round(tradeTaken.PLPoint, 4) = Math.Round(tradeTaken.WarningPLPoint, 4), "FALSE", "TRUE")
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.OverAllMaxDrawUpPL
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.OverAllMaxDrawDownPL
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Supporting1
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Supporting2
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Supporting3
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Supporting4
                                                colCtr += 1
                                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                mainRawData(rowCtr, colCtr) = tradeTaken.Supporting5

                                                rowCtr += 1
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        Next

                        Dim range As String = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                        RaiseEvent Heartbeat("Writing from memory to excel...")
                        excelWriter.WriteArrayToExcel(mainRawData, range)
                        Erase mainRawData
                        mainRawData = Nothing

                        excelWriter.SetActiveSheet("Capital")
                        rowCtr = 0
                        colCtr = 0

                        rowCount = 0
                        If allCapitalData IsNot Nothing AndAlso allCapitalData.Count > 0 Then
                            rowCount = allCapitalData.Values.Sum(Function(x)
                                                                     Return x.Count
                                                                 End Function)
                        End If

                        Dim capitalRawData(rowCount, 0) As Object

                        If rowCtr = 0 Then
                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                            capitalRawData(rowCtr, colCtr) = "Trading Date"
                            colCtr += 1
                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                            capitalRawData(rowCtr, colCtr) = "Capital Exhausted Date Time"
                            colCtr += 1
                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                            capitalRawData(rowCtr, colCtr) = "Running Capital"
                            colCtr += 1
                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                            capitalRawData(rowCtr, colCtr) = "Capital Released"
                            colCtr += 1
                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                            capitalRawData(rowCtr, colCtr) = "Remarks"

                            rowCtr += 1
                        End If

                        If allCapitalData IsNot Nothing AndAlso allCapitalData.Count > 0 Then
                            Dim stockCtr As Integer = 0
                            For Each item In allCapitalData.Keys
                                stockCtr += 1
                                If allCapitalData.ContainsKey(item) Then
                                    Dim capitalList As List(Of Capital) = allCapitalData(item)
                                    Dim itemCtr As Integer = 0
                                    If capitalList IsNot Nothing AndAlso capitalList.Count > 0 Then
                                        For Each data In capitalList
                                            itemCtr += 1
                                            OnHeartbeat(String.Format("Excel printing: {0} of {1}", itemCtr, capitalList.Count))
                                            colCtr = 0

                                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                            capitalRawData(rowCtr, colCtr) = data.TradingDate.ToString("dd-MMM-yyyy")
                                            colCtr += 1
                                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                            capitalRawData(rowCtr, colCtr) = data.CapitalExhaustedDateTime.ToString("HH:mm:ss")
                                            colCtr += 1
                                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(mainRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                            capitalRawData(rowCtr, colCtr) = data.RunningCapital
                                            colCtr += 1
                                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                            capitalRawData(rowCtr, colCtr) = data.CapitalReleased
                                            colCtr += 1
                                            If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                            capitalRawData(rowCtr, colCtr) = data.Remarks

                                            rowCtr += 1
                                        Next
                                    End If
                                End If
                            Next
                        End If

                        range = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                        OnHeartbeat("Writing from memory to excel...")
                        excelWriter.WriteArrayToExcel(capitalRawData, range)
                        Erase capitalRawData
                        capitalRawData = Nothing

                        excelWriter.CreateNewSheet("Summary")
                        excelWriter.SetActiveSheet("Summary")
                        excelWriter.SetCellWidth(1, 6, 33)
                        excelWriter.SetCellWidth(1, 7, 15)
                        Dim n As Integer = 1
                        excelWriter.SetData(n + 10, 6, "Trade Win Ratio")
                        excelWriter.SetData(n + 10, 7, strategyOutputData.WinRatio, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        n = 2
                        excelWriter.SetData(n + 11, 6, "Net Profit")
                        excelWriter.SetData(n + 11, 7, strategyOutputData.NetProfit, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 12, 6, "Gross Profit")
                        excelWriter.SetData(n + 12, 7, strategyOutputData.GrossProfit, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 13, 6, "Gross Loss")
                        excelWriter.SetData(n + 13, 7, strategyOutputData.GrossLoss, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        n = 6
                        excelWriter.SetData(n + 15, 6, "Total Trades")
                        excelWriter.SetData(n + 15, 7, strategyOutputData.TotalTrades, "##,##,##0", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 16, 6, "Total Winning Trades")
                        excelWriter.SetData(n + 16, 7, strategyOutputData.TotalWinningTrades, "##,##,##0", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 17, 6, "Total Losing Trades")
                        excelWriter.SetData(n + 17, 7, strategyOutputData.TotalLosingTrades, "##,##,##0", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 18, 6, "Average Trades")
                        excelWriter.SetData(n + 18, 7, strategyOutputData.AverageTrades, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 19, 6, "Average Winning Trades")
                        excelWriter.SetData(n + 19, 7, strategyOutputData.AverageWinningTrades, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 20, 6, "Average Losing Trades")
                        excelWriter.SetData(n + 20, 7, strategyOutputData.AverageLosingTrades, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 21, 6, "Risk Reward")
                        excelWriter.SetData(n + 21, 7, strategyOutputData.RiskReward, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 22, 6, "Largest Winning Trade")
                        excelWriter.SetData(n + 22, 7, strategyOutputData.LargestWinningTrade, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 23, 6, "Largest Losing Trade")
                        excelWriter.SetData(n + 23, 7, strategyOutputData.LargestLosingTrade, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 24, 6, "Average Duration In Trades")
                        excelWriter.SetData(n + 24, 7, strategyOutputData.AverageDurationInTrades, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 25, 6, "Average Duration In Winning Trades")
                        excelWriter.SetData(n + 25, 7, strategyOutputData.AverageDurationInWinningTrades, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                        excelWriter.SetData(n + 26, 6, "Average Duration In Losing Trades")
                        excelWriter.SetData(n + 26, 7, strategyOutputData.AverageDurationInLosingTrades, "##,##,##0.00", ExcelHelper.XLAlign.Right)

                        excelWriter.SetActiveSheet("Data")
                        OnHeartbeat("Saving excel...")
                        excelWriter.SaveExcel()
                    End Using

                    Dim p As Process = New Process()
                    Dim pi As ProcessStartInfo = New ProcessStartInfo()
                    pi.Arguments = String.Format(" ""{0}"" {1} {2} {3} {4}", filepath, InitialCapital, CapitalForPumpIn, MinimumEarnedCapitalToWithdraw, AmountToBeWithdrawn)
                    pi.FileName = "BacktesterExcelModifier.exe"
                    p.StartInfo = pi
                    p.Start()

                    If tradesFilename IsNot Nothing AndAlso File.Exists(tradesFilename) Then File.Delete(tradesFilename)
                    If capitalFileName IsNot Nothing AndAlso File.Exists(capitalFileName) Then File.Delete(capitalFileName)
                    Exit For
                End If
            Catch ex As System.OutOfMemoryException
                retryCounter += 1
                GC.Collect()
            End Try
            Thread.Sleep(30000)
        Next
    End Sub
#End Region

End Class
