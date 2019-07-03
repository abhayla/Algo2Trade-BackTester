Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class HeikenAshiVolumeRule
    Inherits StrategyRule
    Implements IDisposable

    Private _maxCandleRange As Decimal
    Private _minCandleRangeForTarget As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal maxCandleRange As Decimal, ByVal minCandleRangeForTarget As Decimal, ByVal lotsize As Integer, ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, tickSize, lotsize, canceller)
        _maxCandleRange = maxCandleRange
        _minCandleRangeForTarget = minCandleRangeForTarget
    End Sub
    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            Dim emaPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim atrPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _inputPayload, emaPayload)
            Indicator.ATR.CalculateATR(14, _inputPayload, atrPayload)
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim potentialTargetPoint As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                    If NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.CandleRange, 2) <= NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) * _maxCandleRange Then
                        If _inputPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bullish AndAlso
                            NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Close, 2) >= NumberManipulation.RoundEX(emaPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Open, 2) <= NumberManipulation.RoundEX(emaPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(_inputPayload(runningPayload).Low, 2) < NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Low, 2) AndAlso
                            _inputPayload(runningPayload).PreviousCandlePayload.Volume > _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume * 2 Then
                            signal = -1
                            entryPrice = NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Low, 2)
                            slPrice = NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.High, 2)
                            potentialTargetPoint = slPrice - entryPrice
                            entryPrice -= CalculateBuffer(entryPrice, RoundOfType.Floor)
                            slPrice += CalculateBuffer(slPrice, RoundOfType.Floor)
                            If potentialTargetPoint <= NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) * _minCandleRangeForTarget Then
                                targetPrice = entryPrice - NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) * 2
                            Else
                                targetPrice = entryPrice - NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2)
                            End If
                            supporting1 = NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.CandleRange, 2) / NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2)
                        ElseIf _inputPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bearish AndAlso
                            NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Open, 2) >= NumberManipulation.RoundEX(emaPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Close, 2) <= NumberManipulation.RoundEX(emaPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(_inputPayload(runningPayload).High, 2) > NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.High, 2) AndAlso
                            _inputPayload(runningPayload).PreviousCandlePayload.Volume > _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume * 2 Then
                            signal = 1
                            entryPrice = NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.High, 2)
                            slPrice = NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.Low, 2)
                            potentialTargetPoint = entryPrice - slPrice
                            entryPrice += CalculateBuffer(entryPrice, RoundOfType.Floor)
                            slPrice -= CalculateBuffer(slPrice, RoundOfType.Floor)
                            If potentialTargetPoint <= NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) * _minCandleRangeForTarget Then
                                targetPrice = entryPrice + NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) * 2
                            Else
                                targetPrice = entryPrice + NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2)
                            End If
                            supporting1 = NumberManipulation.RoundEX(_inputPayload(runningPayload).PreviousCandlePayload.CandleRange, 2) / NumberManipulation.RoundEX(atrPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2)
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
