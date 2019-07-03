Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading

Public MustInherit Class StrategyRule
    Protected _inputPayload As Dictionary(Of Date, Payload)
    Protected _tickSize As Decimal
    Protected _quantity As Integer
    Protected _canceller As CancellationTokenSource
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal quantity As Integer, ByVal canceller As CancellationTokenSource)
        _inputPayload = inputPayload
        _tickSize = tickSize
        _canceller = canceller
        _quantity = quantity
    End Sub
    Public MustOverride Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
    Protected Function CalculateBuffer(ByVal numberOfBufferTicks As Integer) As Double
        Dim bufferPrice As Double = Nothing
        'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
        bufferPrice = _tickSize * numberOfBufferTicks
        Return bufferPrice
    End Function
    Protected Function CalculateBuffer(ByVal price As Double, ByVal floorOrCeiling As RoundOfType) As Double
        Dim bufferPrice As Double = Nothing
        'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
        bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, _tickSize, floorOrCeiling)
        Return bufferPrice
    End Function
    Protected Function CalculateQuantityFromSL(ByVal stockName As String, ByVal buyPrice As Double, ByVal sellPrice As Double, ByVal NetProfitLossOfTrade As Double, ByVal typeOfStock As Trade.TypeOfStock) As Integer
        Dim potentialBrokerage As Calculator.BrokerageAttributes = Nothing
        Dim calculator As New Calculator.BrokerageCalculator(_canceller)

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
    Protected Enum SignalType
        Buy = 1
        Sell
        None
    End Enum
End Class
