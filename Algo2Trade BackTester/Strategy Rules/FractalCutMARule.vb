Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FractalCutMARule
    Inherits StrategyRule
    Implements IDisposable
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal lotsize As Integer, ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, tickSize, lotsize, canceller)
    End Sub
    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing

            Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim SMA50Payload As Dictionary(Of Date, Decimal) = Nothing
            Dim SMA200Payload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.Fractals.CalculateFractal(5, _inputPayload, fractalHighPayload, fractalLowPayload)
            Indicator.SMA.CalculateSMA(50, Payload.PayloadFields.Close, _inputPayload, SMA50Payload)
            Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, _inputPayload, SMA200Payload)

            Dim dummySignal As SignalType = SignalType.None
            Dim dummySignalFractal As Decimal = 0
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = 0
                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    If _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                        dummySignal = SignalType.None
                    End If
                    If fractalHighPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) > SMA50Payload(runningPayload) AndAlso
                    fractalHighPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) > SMA200Payload(runningPayload) AndAlso
                    fractalHighPayload(runningPayload) < SMA50Payload(runningPayload) AndAlso
                    fractalHighPayload(runningPayload) < SMA200Payload(runningPayload) Then
                        dummySignal = SignalType.Buy
                        dummySignalFractal = fractalHighPayload(runningPayload)
                    ElseIf fractalLowPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) < SMA50Payload(runningPayload) AndAlso
                        fractalLowPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) < SMA200Payload(runningPayload) AndAlso
                        fractalLowPayload(runningPayload) > SMA50Payload(runningPayload) AndAlso
                        fractalLowPayload(runningPayload) > SMA200Payload(runningPayload) Then
                        dummySignal = SignalType.Sell
                        dummySignalFractal = fractalLowPayload(runningPayload)
                    End If
                    If dummySignal = SignalType.Buy Then
                        If _inputPayload(runningPayload).High > fractalHighPayload(runningPayload) AndAlso
                            fractalHighPayload(runningPayload) = dummySignalFractal Then
                            signal = 1
                            entryPrice = fractalHighPayload(runningPayload)
                            entryPrice += CalculateBuffer(entryPrice, RoundOfType.Floor)
                            slPrice = fractalLowPayload(runningPayload)
                            slPrice -= CalculateBuffer(slPrice, RoundOfType.Floor)
                            targetPrice = entryPrice + (entryPrice - slPrice) * 1
                            'quantity = _lotsize
                            quantity = CalculateQuantityFromSL(_inputPayload(runningPayload).TradingSymbol, entryPrice, slPrice, -2000, Trade.TypeOfStock.Cash)
                            dummySignal = SignalType.None
                        End If
                    ElseIf dummySignal = SignalType.Sell Then
                        If _inputPayload(runningPayload).Low < fractalLowPayload(runningPayload) AndAlso
                            fractalLowPayload(runningPayload) = dummySignalFractal Then
                            signal = -1
                            entryPrice = fractalLowPayload(runningPayload)
                            entryPrice -= CalculateBuffer(entryPrice, RoundOfType.Floor)
                            slPrice = fractalHighPayload(runningPayload)
                            slPrice += CalculateBuffer(slPrice, RoundOfType.Floor)
                            targetPrice = entryPrice - (slPrice - entryPrice) * 1
                            'quantity = _lotsize
                            quantity = CalculateQuantityFromSL(_inputPayload(runningPayload).TradingSymbol, slPrice, entryPrice, -2000, Trade.TypeOfStock.Cash)
                            dummySignal = SignalType.None
                        End If
                    End If
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
                End If
            Next
            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
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
