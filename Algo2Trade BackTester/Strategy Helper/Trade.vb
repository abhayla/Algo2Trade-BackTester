Imports Algo2TradeBLL
<Serializable>
Public Class Trade

#Region "Enum"
    Public Enum TypeOfStock
        Cash = 1
        Currency
        Commodity
        Futures
        None
    End Enum
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
        None
    End Enum
    Public Enum TradeEntryCondition
        Original = 1
        Reversal
        Onward
        None
    End Enum
    Public Enum TradeExitCondition
        Target = 1
        StopLoss
        EndOfDay
        Cancelled
        ForceExit
        None
    End Enum
    Public Enum TradeType
        MIS = 1
        CNC
        None
    End Enum
#End Region

#Region "Constructor"
    Public Sub New(ByVal originatingStrategy As Strategy,
                       ByVal tradingSymbol As String,
                       ByVal stockType As TypeOfStock,
                       ByVal tradingDate As Date,
                       ByVal entryDirection As TradeExecutionDirection,
                       ByVal entryPrice As Double,
                       ByVal entryBuffer As Decimal,
                       ByVal squareOffType As TradeType,
                       ByVal entryCondition As TradeEntryCondition,
                       ByVal entryRemark As String,
                       ByVal quantity As Integer,
                       ByVal potentialTarget As Double,
                       ByVal targetRemark As String,
                       ByVal potentialStopLoss As Double,
                       ByVal stoplossBuffer As Decimal,
                       ByVal slRemark As String,
                       ByVal signalCandle As Payload)
        Me._OriginatingStrategy = originatingStrategy
        Me._TradingSymbol = tradingSymbol
        Me._StockType = stockType
        Me._EntryTime = tradingDate
        Me._TradingDate = tradingDate.Date
        Me._EntryDirection = entryDirection
        Me._EntryPrice = Math.Round(entryPrice, 4)
        Me._EntryBuffer = entryBuffer
        Me._SquareOffType = squareOffType
        Me._EntryCondition = entryCondition
        Me._EntryRemark = entryRemark
        Me._Quantity = quantity
        Me._PotentialTarget = Math.Round(potentialTarget, 4)
        Me._TargetRemark = targetRemark
        Me._PotentialStopLoss = Math.Round(potentialStopLoss, 4)
        Me._StoplossBuffer = stoplossBuffer
        Me._SLRemark = slRemark
        Me._StoplossSetTime = Me._EntryTime
        Me._SignalCandle = signalCandle
        Me._TradeUpdateTimeStamp = Me._EntryTime
        Me._TradeCurrentStatus = TradeExecutionStatus.None
        Me._ExitTime = Date.MinValue
        Me._ExitPrice = Double.MinValue
        Me._ExitCondition = TradeExitCondition.None
        Me._ExitRemark = Nothing
    End Sub
    'Start Indibar
    Public Sub New(ByVal originatingStrategy As Strategy,
                       ByVal tradingSymbol As String,
                       ByVal stockType As TypeOfStock,
                       ByVal tradingDate As Date,
                       ByVal entryDirection As TradeExecutionDirection,
                       ByVal entryPrice As Double,
                       ByVal entryBuffer As Decimal,
                       ByVal squareOffType As TradeType,
                       ByVal entryCondition As TradeEntryCondition,
                       ByVal entryRemark As String,
                       ByVal quantity As Integer,
                       ByVal lotSize As Integer,
                       ByVal potentialTarget As Double,
                       ByVal targetRemark As String,
                       ByVal potentialStopLoss As Double,
                       ByVal stoplossBuffer As Decimal,
                       ByVal slRemark As String,
                       ByVal signalCandle As Payload)
        Me._OriginatingStrategy = originatingStrategy
        Me._TradingSymbol = tradingSymbol
        Me._StockType = stockType
        Me._EntryTime = tradingDate
        Me._TradingDate = tradingDate.Date
        Me._EntryDirection = entryDirection
        Me._EntryPrice = entryPrice
        Me._EntryBuffer = entryBuffer
        Me._SquareOffType = squareOffType
        Me._EntryCondition = entryCondition
        Me._EntryRemark = entryRemark
        Me._Quantity = quantity
        Me._LotSize = lotSize
        Me._PotentialTarget = potentialTarget
        Me._TargetRemark = targetRemark
        Me._PotentialStopLoss = potentialStopLoss
        Me._StoplossBuffer = stoplossBuffer
        Me._SLRemark = slRemark
        Me._StoplossSetTime = Me._EntryTime
        Me._SignalCandle = signalCandle
        Me._TradeUpdateTimeStamp = Me._EntryTime
        Me._TradeCurrentStatus = TradeExecutionStatus.None
        Me._ExitTime = Date.MinValue
        Me._ExitPrice = Double.MinValue
        Me._ExitCondition = TradeExitCondition.None
        Me._ExitRemark = Nothing
    End Sub
    Public Sub New(ByVal originatingStrategy As Strategy,
                    ByVal tradingSymbol As String,
                    ByVal stockType As TypeOfStock,
                    ByVal tradingDate As Date,
                    ByVal entryDirection As TradeExecutionDirection,
                    ByVal entryPrice As Double,
                    ByVal entryBuffer As Decimal,
                    ByVal squareOffType As TradeType,
                    ByVal entryCondition As TradeEntryCondition,
                    ByVal entryRemark As String,
                    ByVal quantity As Integer,
                    ByVal lotSize As Integer,
                    ByVal potentialTarget As Double,
                    ByVal targetRemark As String,
                    ByVal potentialStopLoss As Double,
                    ByVal stoplossBuffer As Decimal,
                    ByVal slRemark As String,
                    ByVal signalCandle As Payload,
                    ByVal firstEntryDirection As TradeExecutionDirection,
                    ByVal additionalTrade As Boolean)
        Me._OriginatingStrategy = originatingStrategy
        Me._TradingSymbol = tradingSymbol
        Me._StockType = stockType
        Me._EntryTime = tradingDate
        Me._TradingDate = tradingDate.Date
        Me._EntryDirection = entryDirection
        Me._FirstEntryDirection = firstEntryDirection
        Me._AdditionalTrade = additionalTrade
        Me._EntryPrice = entryPrice
        Me._EntryBuffer = entryBuffer
        Me._SquareOffType = squareOffType
        Me._EntryCondition = entryCondition
        Me._EntryRemark = entryRemark
        Me._Quantity = quantity
        Me._LotSize = lotSize
        Me._PotentialTarget = potentialTarget
        Me._TargetRemark = targetRemark
        Me._PotentialStopLoss = potentialStopLoss
        Me._StoplossBuffer = stoplossBuffer
        Me._SLRemark = slRemark
        Me._StoplossSetTime = Me._EntryTime
        Me._SignalCandle = signalCandle
        Me._TradeUpdateTimeStamp = Me._EntryTime
        Me._TradeCurrentStatus = TradeExecutionStatus.None
        Me._ExitTime = Date.MinValue
        Me._ExitPrice = Double.MinValue
        Me._ExitCondition = TradeExitCondition.None
        Me._ExitRemark = Nothing
    End Sub
    'End Indibar
#End Region

#Region "Variables"
    <NonSerialized>
    Private _OriginatingStrategy As Strategy
    Public ReadOnly Property TradingSymbol As String
    Public ReadOnly Property CoreTradingSymbol As String
        Get
            If TradingSymbol.Contains("FUT") Then
                Return TradingSymbol.Remove(TradingSymbol.Count - 8)
            Else
                Return TradingSymbol
            End If
        End Get
    End Property
    Public ReadOnly Property StockType As TypeOfStock
    Public ReadOnly Property EntryTime As DateTime
    Public ReadOnly Property TradingDate As Date '''''
    Public ReadOnly Property EntryDirection As TradeExecutionDirection
    Public ReadOnly Property FirstEntryDirection As TradeExecutionDirection
    Public ReadOnly Property AdditionalTrade As Boolean
    Public ReadOnly Property EntryPrice As Double
    Public ReadOnly Property EntryBuffer As Decimal
    Public ReadOnly Property SquareOffType As TradeType
    Public ReadOnly Property EntryCondition As TradeEntryCondition
    Public ReadOnly Property EntryRemark As String
    Public ReadOnly Property Quantity As Integer
    Public ReadOnly Property LotSize As Integer
    Public ReadOnly Property PotentialTarget As Double
    Public ReadOnly Property TargetRemark As String
    Public ReadOnly Property PotentialStopLoss As Double
    Public ReadOnly Property StoplossBuffer As Decimal
    Public ReadOnly Property SLRemark As String
    Public ReadOnly Property StoplossSetTime As Date
    Public ReadOnly Property SignalCandle As Payload
    Public ReadOnly Property TradeUpdateTimeStamp As Date
    Public ReadOnly Property TradeCurrentStatus As TradeExecutionStatus
    Public ReadOnly Property ExitTime As DateTime
    Public ReadOnly Property ExitPrice As Double
    Public ReadOnly Property ExitCondition As TradeExitCondition
    Public ReadOnly Property ExitRemark As String
    Public ReadOnly Property Tag As String
    Public ReadOnly Property SquareOffValue As Double
    Public ReadOnly Property Supporting1 As String
    Public ReadOnly Property Supporting2 As String
    Public ReadOnly Property Supporting3 As String
    Public ReadOnly Property Supporting4 As String
    Public ReadOnly Property Supporting5 As String

    Private _CurrentLTP As Double
    Public Property CurrentLTP As Double
        Get
            Return _CurrentLTP
        End Get
        Set(value As Double)
            _CurrentLTP = value
            'Trigger lazy properties
            Dim x = Me.MaximumDrawDown
            Dim y = Me.MaximumDrawUp
            If Me.MaximumDrawDown = 0 Or Me.MaximumDrawUp = 0 Then
                Throw New ApplicationException(String.Format("{0} Value: 0", If(Me.MaximumDrawUp = 0, Me.MaximumDrawUp.ToString, Me.MaximumDrawDown.ToString)))
            End If
            Dim a = _OriginatingStrategy.AllMaxDrawDownPLAfterBrokerage(Me.TradingDate)
            Dim b = _OriginatingStrategy.AllMaxDrawUpPLAfterBrokerage(Me.TradingDate)
        End Set
    End Property
    Public ReadOnly Property PLAfterBrokerage As Double
        Get
            If EntryDirection = TradeExecutionDirection.Buy Then
                Return Strategy.CalculatePL(CoreTradingSymbol, EntryPrice, CurrentLTP, Quantity, LotSize, StockType)
            ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                Return Strategy.CalculatePL(CoreTradingSymbol, CurrentLTP, EntryPrice, Quantity, LotSize, StockType)
            Else
                Return Double.MinValue
            End If
        End Get
    End Property
    Public ReadOnly Property PLBeforeBrokerage As Double
        Get
            If Me.EntryDirection = TradeExecutionDirection.Buy Then
                Return ((Me.CurrentLTP - Me.EntryPrice) * Me.Quantity)
            ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                Return ((Me.EntryPrice - Me.CurrentLTP) * Me.Quantity)
            Else
                Return Double.MinValue
            End If
        End Get
    End Property
    Public ReadOnly Property CapitalRequiredWithMargin As Double
        Get
            Return Me.EntryPrice * Me.Quantity / Strategy.MarginMultiplier
        End Get
    End Property

    Private _MaximumDrawUp As Double = Double.MinValue
    Public ReadOnly Property MaximumDrawUp As Double
        Get
            If TradeCurrentStatus = TradeExecutionStatus.Inprogress Then
                If _MaximumDrawUp = Double.MinValue Then
                    _MaximumDrawUp = Me.EntryPrice
                End If
                If EntryDirection = TradeExecutionDirection.Buy Then
                    _MaximumDrawUp = Math.Max(_MaximumDrawUp, CurrentLTP)
                ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                    _MaximumDrawUp = Math.Min(_MaximumDrawUp, CurrentLTP)
                End If
            End If
            Return _MaximumDrawUp
        End Get
    End Property

    Private _MaximumDrawDown As Double = Double.MinValue
    Public ReadOnly Property MaximumDrawDown As Double
        Get
            If TradeCurrentStatus = TradeExecutionStatus.Inprogress Then
                If _MaximumDrawDown = Double.MinValue Then
                    _MaximumDrawDown = Me.EntryPrice
                End If
                If EntryDirection = TradeExecutionDirection.Buy Then
                    _MaximumDrawDown = Math.Min(_MaximumDrawDown, CurrentLTP)
                ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                    _MaximumDrawDown = Math.Max(_MaximumDrawDown, CurrentLTP)
                End If
            End If
            Return _MaximumDrawDown
        End Get
    End Property
    Public ReadOnly Property MaximumDrawUpPL As Double
        Get
            If EntryDirection = TradeExecutionDirection.Buy Then
                Return Strategy.CalculatePL(CoreTradingSymbol, EntryPrice, MaximumDrawUp, Quantity, LotSize, StockType)
            ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                Return Strategy.CalculatePL(CoreTradingSymbol, MaximumDrawUp, EntryPrice, Quantity, LotSize, StockType)
            Else
                Return Double.MinValue
            End If
        End Get
    End Property
    Public ReadOnly Property MaximumDrawDownPL As Double
        Get
            'Try
            If EntryDirection = TradeExecutionDirection.Buy Then
                Return Strategy.CalculatePL(CoreTradingSymbol, EntryPrice, MaximumDrawDown, Quantity, LotSize, StockType)
            ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                Return Strategy.CalculatePL(CoreTradingSymbol, MaximumDrawDown, EntryPrice, Quantity, LotSize, StockType)
            Else
                Return Double.MinValue
            End If
            'Catch ex As Exception
            '    Throw ex
            'End Try
        End Get
    End Property
    Public ReadOnly Property DurationOfTrade As TimeSpan
        Get
            Return Me.ExitTime - Me.EntryTime
        End Get
    End Property
    Public ReadOnly Property PLPoint As Double
        Get
            If Me.EntryDirection = TradeExecutionDirection.Buy Then
                Return (Me.CurrentLTP - Me.EntryPrice)
            ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                Return (Me.EntryPrice - Me.CurrentLTP)
            Else
                Return Double.MinValue
            End If
        End Get
    End Property
    'Start Indibar
    Public ReadOnly Property WarningPLPoint As Double
        Get
            If Me.EntryDirection = TradeExecutionDirection.Buy Then
                Return (Me.ExitPrice - Me.EntryPrice)
            ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                Return (Me.EntryPrice - Me.ExitPrice)
            Else
                Return Double.MinValue
            End If
        End Get
    End Property
    Public Property OverAllMaxDrawUpPL As Double
    Public Property OverAllMaxDrawDownPL As Double
    'End Indibar
#End Region

#Region "Public Fuction"
    'Public Overrides Function ToString() As String
    '    Return String.Format("TradingSymbol:{0},Direction:{1},EntryTime:{2},EntryPrice:{3},EntryRemark:{4},TP:{5},TPRemark:{6},SL:{7},SLRemark:{8},ExitTime:{9},ExitPrice:{10},ExitRemark:{11},Status:{12}",
    '                                                                                                                                            Me.TradingSymbol,
    '                                                                                                                                            Me.EntryDirection,
    '                                                                                                                                            Me.EntryTime.ToString("yyyy-MM-dd HH:mm:ss"),
    '                                                                                                                                            Me.EntryPrice,
    '                                                                                                                                            Me.EntryRemark,
    '                                                                                                                                            Me.PotentialTarget,
    '                                                                                                                                            Me.TargetRemark,
    '                                                                                                                                            Me.PotentialStopLoss,
    '                                                                                                                                            Me.SLRemark,
    '                                                                                                                                            Me.ExitTime.ToString("yyyy-MM-dd HH:mm:ss"),
    '                                                                                                                                            Me.ExitPrice,
    '                                                                                                                                            Me.ExitRemark,
    '                                                                                                                                            Me.TradeCurrentStatus)
    'End Function
    Public Sub UpdateTrade(Optional ByVal TradingSymbol As String = Nothing,
                            Optional ByVal StockType As TypeOfStock = TypeOfStock.None,
                            Optional ByVal EntryTime As Date = Nothing,
                            Optional ByVal TradingDate As Date = Nothing,
                            Optional ByVal EntryDirection As TradeExecutionDirection = TradeExecutionDirection.None,
                            Optional ByVal EntryPrice As Double = Double.MinValue,
                            Optional ByVal EntryBuffer As Decimal = Decimal.MinValue,
                            Optional ByVal SquareOffType As TradeType = TradeType.None,
                            Optional ByVal EntryCondition As TradeEntryCondition = TradeEntryCondition.None,
                            Optional ByVal EntryRemark As String = Nothing,
                            Optional ByVal Quantity As Integer = Integer.MinValue,
                            Optional ByVal PotentialTarget As Double = Double.MinValue,
                            Optional ByVal TargetRemark As String = Nothing,
                            Optional ByVal PotentialStopLoss As Double = Double.MinValue,
                            Optional ByVal StoplossBuffer As Decimal = Decimal.MinValue,
                            Optional ByVal SLRemark As String = Nothing,
                            Optional ByVal StoplossSetTime As Date = Nothing,
                            Optional ByVal SignalCandle As Payload = Nothing,
                            Optional ByVal TradeCurrentStatus As TradeExecutionStatus = TradeExecutionStatus.None,
                            Optional ByVal ExitTime As Date = Nothing,
                            Optional ByVal ExitPrice As Double = Double.MinValue,
                            Optional ByVal ExitCondition As TradeExitCondition = TradeExitCondition.None,
                            Optional ByVal ExitRemark As String = Nothing,
                            Optional ByVal Tag As String = Nothing,
                            Optional ByVal SquareOffValue As Double = Double.MinValue,
                            Optional ByVal AdditionalTrade As Boolean = False,
                            Optional ByVal Supporting1 As String = Nothing,
                            Optional ByVal Supporting2 As String = Nothing,
                            Optional ByVal Supporting3 As String = Nothing,
                            Optional ByVal Supporting4 As String = Nothing,
                            Optional ByVal Supporting5 As String = Nothing)


        If TradingSymbol IsNot Nothing Then _TradingSymbol = TradingSymbol
        If StockType <> TypeOfStock.None Then _StockType = StockType
        If EntryTime <> Nothing OrElse EntryTime <> Date.MinValue Then _EntryTime = EntryTime
        If TradingDate <> Nothing OrElse TradingDate <> Date.MinValue Then _TradingDate = TradingDate
        If EntryDirection <> TradeExecutionDirection.None Then _EntryDirection = EntryDirection
        If EntryPrice <> Double.MinValue Then _EntryPrice = Math.Round(EntryPrice, 4)
        If EntryBuffer <> Decimal.MinValue Then _EntryBuffer = EntryBuffer
        If SquareOffType <> TradeType.None Then _SquareOffType = SquareOffType
        If EntryCondition <> TradeEntryCondition.None Then _EntryCondition = EntryCondition
        If EntryRemark IsNot Nothing Then _EntryRemark = EntryRemark
        If Quantity <> Integer.MinValue Then _Quantity = Quantity
        If PotentialTarget <> Double.MinValue Then _PotentialTarget = Math.Round(PotentialTarget, 4)
        If TargetRemark IsNot Nothing Then _TargetRemark = TargetRemark
        If PotentialStopLoss <> Double.MinValue Then _PotentialStopLoss = Math.Round(PotentialStopLoss, 4)
        If StoplossBuffer <> Decimal.MinValue Then _StoplossBuffer = StoplossBuffer
        If SLRemark IsNot Nothing Then _SLRemark = SLRemark
        If StoplossSetTime <> Nothing OrElse StoplossSetTime <> Date.MinValue Then _StoplossSetTime = StoplossSetTime
        If SignalCandle IsNot Nothing Then _SignalCandle = SignalCandle
        If TradeCurrentStatus <> TradeExecutionStatus.None Then _TradeCurrentStatus = TradeCurrentStatus
        If ExitTime <> Nothing OrElse ExitTime <> Date.MinValue Then _ExitTime = ExitTime
        If ExitPrice <> Double.MinValue Then _ExitPrice = ExitPrice
        If ExitCondition <> TradeExitCondition.None Then _ExitCondition = ExitCondition
        If ExitRemark IsNot Nothing Then _ExitRemark = ExitRemark
        If Tag IsNot Nothing Then _Tag = Tag
        If SquareOffValue <> Double.MinValue Then _SquareOffValue = SquareOffValue
        If AdditionalTrade Then _AdditionalTrade = AdditionalTrade
        If Supporting1 IsNot Nothing Then _Supporting1 = Supporting1
        If Supporting2 IsNot Nothing Then _Supporting2 = Supporting2
        If Supporting3 IsNot Nothing Then _Supporting3 = Supporting3
        If Supporting4 IsNot Nothing Then _Supporting4 = Supporting4
        If Supporting5 IsNot Nothing Then _Supporting5 = Supporting5

        If Me._ExitTime <> Nothing OrElse Me._ExitTime <> Date.MinValue Then
            Me._TradeUpdateTimeStamp = Me._ExitTime
        ElseIf Me._StoplossSetTime <> Nothing OrElse Me._StoplossSetTime <> Date.MinValue Then
            Me._TradeUpdateTimeStamp = Me._StoplossSetTime
        ElseIf Me._EntryTime <> Nothing OrElse Me._EntryTime <> Date.MinValue Then
            Me._TradeUpdateTimeStamp = Me._EntryTime
        End If
    End Sub
    Public Sub UpdateTrade(ByVal tradeToBeUsed As Trade)
        If tradeToBeUsed Is Nothing Then Exit Sub
        With tradeToBeUsed
            If .TradingSymbol IsNot Nothing Then _TradingSymbol = .TradingSymbol
            If .StockType <> TypeOfStock.None Then _StockType = .StockType
            If .EntryTime <> Nothing OrElse .EntryTime <> Date.MinValue Then _EntryTime = .EntryTime
            If .TradingDate <> Nothing OrElse .TradingDate <> Date.MinValue Then _TradingDate = .TradingDate
            If .EntryDirection <> TradeExecutionDirection.None Then _EntryDirection = .EntryDirection
            If .EntryPrice <> Double.MinValue Then _EntryPrice = .EntryPrice
            If .EntryBuffer <> Decimal.MinValue Then _EntryBuffer = .EntryBuffer
            If .SquareOffType <> TradeType.None Then _SquareOffType = .SquareOffType
            If .EntryCondition <> TradeEntryCondition.None Then _EntryCondition = .EntryCondition
            If .EntryRemark IsNot Nothing Then _EntryRemark = .EntryRemark
            If .Quantity <> Integer.MinValue Then _Quantity = .Quantity
            If .PotentialTarget <> Double.MinValue Then _PotentialTarget = .PotentialTarget
            If .TargetRemark IsNot Nothing Then _TargetRemark = .TargetRemark
            If .PotentialStopLoss <> Double.MinValue Then _PotentialStopLoss = .PotentialStopLoss
            If .StoplossBuffer <> Decimal.MinValue Then _StoplossBuffer = .StoplossBuffer
            If .SLRemark IsNot Nothing Then _SLRemark = .SLRemark
            If .StoplossSetTime <> Nothing OrElse .StoplossSetTime <> Date.MinValue Then _StoplossSetTime = .StoplossSetTime
            If .SignalCandle IsNot Nothing Then _SignalCandle = .SignalCandle
            If .TradeCurrentStatus <> TradeExecutionStatus.None Then _TradeCurrentStatus = .TradeCurrentStatus
            If .ExitTime <> Nothing OrElse .ExitTime <> Date.MinValue Then _ExitTime = .ExitTime
            If .ExitPrice <> Double.MinValue Then _ExitPrice = .ExitPrice
            If .ExitCondition <> TradeExitCondition.None Then _ExitCondition = .ExitCondition
            If .ExitRemark IsNot Nothing Then _ExitRemark = .ExitRemark
            If .Tag IsNot Nothing Then _Tag = .Tag
            If .SquareOffValue <> Double.MinValue Then _SquareOffValue = .SquareOffValue
            If .Supporting1 IsNot Nothing Then _Supporting1 = .Supporting1
            If .Supporting2 IsNot Nothing Then _Supporting2 = .Supporting2
            If .Supporting3 IsNot Nothing Then _Supporting3 = .Supporting3
            If .Supporting4 IsNot Nothing Then _Supporting4 = .Supporting4
            If .Supporting5 IsNot Nothing Then _Supporting5 = .Supporting5
        End With

        If Me._ExitTime <> Nothing OrElse Me._ExitTime <> Date.MinValue Then
            Me._TradeUpdateTimeStamp = Me._ExitTime
        ElseIf Me._StoplossSetTime <> Nothing OrElse Me._StoplossSetTime <> Date.MinValue Then
            Me._TradeUpdateTimeStamp = Me._StoplossSetTime
        ElseIf Me._EntryTime <> Nothing OrElse Me._EntryTime <> Date.MinValue Then
            Me._TradeUpdateTimeStamp = Me._EntryTime
        End If
    End Sub
#End Region
End Class
