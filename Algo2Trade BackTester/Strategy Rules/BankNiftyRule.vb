Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Public Class BankNiftyRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly flag As StrategyRule.TypeOfSignal = TypeOfSignal.Sell
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal quantity As Integer, ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
    End Sub
    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing

            Dim cmn As Common = New Common(Nothing)
            Dim tradingDate As Date = _inputPayload.LastOrDefault.Value.PayloadDate.Date
            Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
            tradingSymbol = tradingSymbol.Remove(tradingSymbol.Count - 8)
            Dim eodPayload As Dictionary(Of Date, Payload) = cmn.GetRawPayload(Common.DataBaseTable.EOD_Futures, tradingSymbol, tradingDate.AddDays(-200), tradingDate)
            Dim ATRPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.ATR.CalculateATR(14, eodPayload, ATRPayload, True)

            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing
                Dim dummySignalPassTime As Date = runningPayload

                If _inputPayload(runningPayload).PayloadDate.Hour = 9 AndAlso
                    _inputPayload(runningPayload).PayloadDate.Minute = 16 Then
                    Dim previousTradingDay As Date = cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Futures, tradingSymbol, runningPayload.Date)
                    Dim signalPayload As Payload = _inputPayload(runningPayload)
                    'Dim randomNumber As Random = New Random
                    'Dim signalPoint As Integer = randomNumber.Next(0, signalPayload.Ticks.Count)
                    If eodPayload.ContainsKey(previousTradingDay) AndAlso
                        Math.Abs((eodPayload(runningPayload.Date).Open / eodPayload(previousTradingDay).Close) - 1) * 100 >= 0 Then
                        If flag = TypeOfSignal.Buy Then
                            signal = 1
                            entryPrice = signalPayload.Ticks.FindAll(Function(x)
                                                                         Return x.PayloadDate.Hour = 9 AndAlso x.PayloadDate.Minute = 16 AndAlso x.PayloadDate.Second >= 0
                                                                     End Function).FirstOrDefault.Open
                            'entryPrice += CalculateBuffer(entryPrice, RoundOfType.Floor)
                            Dim slPoint As Decimal = ATRPayload(previousTradingDay.Date) / 8
                            slPrice = entryPrice - slPoint
                            Dim lossAmount As Double = Strategy.CalculatePL(signalPayload.TradingSymbol, entryPrice, slPrice, quantity, 20, Trade.TypeOfStock.Futures)
                            targetPrice = Strategy.CalculatorTargetOrStoploss(signalPayload.TradingSymbol, entryPrice, quantity, If(lossAmount < 0, lossAmount * -2, lossAmount * 2), Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Futures)
                            Dim targetPoint As Decimal = targetPrice - entryPrice
                            Dim extraPoint As Decimal = targetPoint - slPoint * 2
                            Dim adjustablePoint As Decimal = extraPoint / 2
                            targetPoint = slPoint * 2
                            'targetPoint = slPoint * 2 + adjustablePoint
                            slPoint = slPoint - adjustablePoint
                            slPrice = entryPrice - slPoint
                            targetPrice = entryPrice + targetPoint
                        ElseIf flag = TypeOfSignal.Sell Then
                            signal = -1
                            entryPrice = signalPayload.Ticks.FindAll(Function(x)
                                                                         Return x.PayloadDate.Hour = 9 AndAlso x.PayloadDate.Minute = 16 AndAlso x.PayloadDate.Second >= 0
                                                                     End Function).FirstOrDefault.Open
                            'entryPrice += CalculateBuffer(entryPrice, RoundOfType.Floor)
                            Dim slPoint As Decimal = ATRPayload(previousTradingDay.Date) / 8
                            slPrice = entryPrice + slPoint
                            Dim lossAmount As Double = Strategy.CalculatePL(signalPayload.TradingSymbol, slPrice, entryPrice, quantity, 20, Trade.TypeOfStock.Futures)
                            targetPrice = Strategy.CalculatorTargetOrStoploss(signalPayload.TradingSymbol, entryPrice, quantity, If(lossAmount < 0, lossAmount * -2, lossAmount * 2), Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Futures)
                            Dim targetPoint As Decimal = entryPrice - targetPrice
                            Dim extraPoint As Decimal = targetPoint - slPoint * 2
                            Dim adjustablePoint As Decimal = extraPoint / 2
                            targetPoint = slPoint * 2
                            'targetPoint = slPoint * 2 + adjustablePoint
                            slPoint = slPoint - adjustablePoint
                            slPrice = entryPrice + slPoint
                            targetPrice = entryPrice - targetPoint
                        End If
                        supporting1 = Math.Round(((eodPayload(runningPayload.Date).Open / eodPayload(previousTradingDay).Close) - 1) * 100, 2)
                    End If
                    dummySignalPassTime = New Date(runningPayload.Year, runningPayload.Month, runningPayload.Day, 9, 15, 0)
                End If

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                    outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, signal)
                    If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
                    outputEntryPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryPrice)
                    If outputTargetPayload Is Nothing Then outputTargetPayload = New Dictionary(Of Date, Decimal)
                    outputTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, targetPrice)
                    If outputStoplossPayload Is Nothing Then outputStoplossPayload = New Dictionary(Of Date, Decimal)
                    outputStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, slPrice)
                    If outputQuantityPayload Is Nothing Then outputQuantityPayload = New Dictionary(Of Date, Integer)
                    outputQuantityPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, quantity)
                    If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                    outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
                End If
            Next
            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
            If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
        End If
    End Sub

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
