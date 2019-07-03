Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Public Class GapFillRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _SignalDirection As Integer
    Private ReadOnly _PreviousDayClose As Decimal
    Private ReadOnly _PreviousDayATR As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal lotsize As Integer, ByVal canceller As CancellationTokenSource, ByVal signalDirection As Integer, ByVal previousDayClose As Decimal, ByVal previousDayATR As Decimal)
        MyBase.New(inputPayload, tickSize, lotsize, canceller)
        _SignalDirection = signalDirection
        _PreviousDayClose = previousDayClose
        _PreviousDayATR = previousDayATR
    End Sub
    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            Dim firstCandleTime As Date = _inputPayload.FirstOrDefault.Key

            'Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
            'Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim ATRPayload As Dictionary(Of Date, Decimal) = Nothing
            'Indicator.Fractals.CalculateFractal(5, _inputPayload, fractalHighPayload, fractalLowPayload)
            Indicator.VWAP.CalculateVWAP(_inputPayload, VWAPPayload)
            Indicator.ATR.CalculateATR(14, _inputPayload, ATRPayload)

            Dim potentialEntry As Decimal = 0
            Dim oncePerDay As Boolean = True
            'Dim previousSL As Decimal = 0
            'Dim previousTarget As Decimal = 0
            'Dim potentialTargetPoint As Decimal = 0
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing
                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                    potentialEntry = 0
                    oncePerDay = True
                    firstCandleTime = runningPayload
                End If
                If oncePerDay AndAlso runningPayload.Hour <= 9 AndAlso runningPayload.Minute <= 45 Then
                    If _SignalDirection = 1 Then
                        If potentialEntry <> 0 AndAlso _inputPayload(runningPayload).High >= potentialEntry Then
                            If Math.Round(((potentialEntry / _PreviousDayClose) - 1) * 100, 2) < 0.5 Then
                                signal = 1
                                entryPrice = potentialEntry
                                'slPrice = entryPrice - entryPrice * 0.0025
                                'slPrice = entryPrice - ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 2
                                slPrice = _inputPayload.Values.Min(Function(x)
                                                                       If x.PayloadDate >= firstCandleTime AndAlso x.PayloadDate < runningPayload Then
                                                                           Return x.Low
                                                                       Else
                                                                           Return Integer.MaxValue
                                                                       End If
                                                                   End Function)
                                slPrice -= CalculateBuffer(slPrice, RoundOfType.Floor)
                                slPrice = Math.Max(slPrice, entryPrice - entryPrice * 0.01)
                                targetPrice = entryPrice + _PreviousDayATR / 3
                                targetPrice = Math.Min(targetPrice, entryPrice + entryPrice * 0.01)
                                supporting1 = Math.Round(((entryPrice / _PreviousDayClose) - 1) * 100, 2)
                                'potentialTargetPoint = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3
                                oncePerDay = False
                            Else
                                potentialEntry = 0
                                oncePerDay = False
                            End If
                        End If
                        If signal = 0 AndAlso _inputPayload(runningPayload).Low > VWAPPayload(runningPayload) Then
                            potentialEntry = _inputPayload(runningPayload).High
                            potentialEntry += CalculateBuffer(potentialEntry, RoundOfType.Floor)
                        End If
                    ElseIf _SignalDirection = -1 Then
                        If potentialEntry <> 0 AndAlso _inputPayload(runningPayload).Low <= potentialEntry Then
                            If Math.Round(((potentialEntry / _PreviousDayClose) - 1) * 100, 2) > 0.5 Then
                                signal = -1
                                entryPrice = potentialEntry
                                'slPrice = entryPrice + entryPrice * 0.0025
                                'slPrice = entryPrice + ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 2
                                slPrice = _inputPayload.Values.Max(Function(x)
                                                                       If x.PayloadDate >= firstCandleTime AndAlso x.PayloadDate < runningPayload Then
                                                                           Return x.High
                                                                       Else
                                                                           Return Integer.MinValue
                                                                       End If
                                                                   End Function)
                                slPrice += CalculateBuffer(slPrice, RoundOfType.Floor)
                                slPrice = Math.Min(slPrice, entryPrice + entryPrice * 0.01)
                                targetPrice = entryPrice - _PreviousDayATR / 3
                                targetPrice = Math.Max(targetPrice, entryPrice - entryPrice * 0.01)
                                supporting1 = Math.Round(((entryPrice / _PreviousDayClose) - 1) * 100, 2)
                                'potentialTargetPoint = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3
                                oncePerDay = False
                            Else
                                potentialEntry = 0
                                oncePerDay = False
                            End If
                        End If
                        If signal = 0 AndAlso _inputPayload(runningPayload).High < VWAPPayload(runningPayload) Then
                            potentialEntry = _inputPayload(runningPayload).Low
                            potentialEntry -= CalculateBuffer(potentialEntry, RoundOfType.Floor)
                        End If
                    End If
                End If
                'If potentialEntry <> 0 Then
                '    If _SignalDirection = 1 Then
                '        'If fractalLowPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) < VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                '        '    slPrice = fractalLowPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                '        '    slPrice -= CalculateBuffer(slPrice, RoundOfType.Floor)
                '        'Else
                '        '    slPrice = previousSL
                '        'End If
                '        If ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3 <= potentialTargetPoint Then
                '            targetPrice = potentialEntry + ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3
                '            potentialTargetPoint = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3
                '        Else
                '            targetPrice = previousTarget
                '        End If
                '    ElseIf _SignalDirection = -1 Then
                '        'If fractalHighPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) > VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                '        '    slPrice = fractalHighPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                '        '    slPrice += CalculateBuffer(slPrice, RoundOfType.Floor)
                '        'Else
                '        '    slPrice = previousSL
                '        'End If
                '        If ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3 <= potentialTargetPoint Then
                '            targetPrice = potentialEntry - ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3
                '            potentialTargetPoint = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * 3
                '        Else
                '            targetPrice = previousTarget
                '        End If
                '    End If
                '    previousTarget = targetPrice
                '    'previousSL = slPrice
                'End If
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
