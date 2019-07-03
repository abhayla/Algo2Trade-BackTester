Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PinBarWithFractalRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _nextSignalBreak As Boolean
    Private ReadOnly _breakevenMovement As Boolean
    Private ReadOnly _targetMultiplier As Decimal = 3
    Private ReadOnly _minimumCandleWick As Decimal = 0.15
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal quantity As Integer, ByVal canceller As CancellationTokenSource, ByVal nextSignalBreak As Boolean, ByVal breakevenMove As Boolean)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _nextSignalBreak = nextSignalBreak
        _breakevenMovement = breakevenMove
    End Sub
    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            'Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing

            Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.Fractals.CalculateFractal(5, _inputPayload, fractalHighPayload, fractalLowPayload)

            Dim potentialEntryPrice As Decimal = 0
            Dim potentialSLPrice As Decimal = 0
            Dim potentialEntryDirection As Integer = 0
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                'Dim supporting1 As String = Nothing
                'Dim supporting2 As String = Nothing
                'Dim supporting3 As String = Nothing
                'Dim supporting4 As String = Nothing
                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                    signal = 0
                    entryPrice = 0
                    targetPrice = 0
                    slPrice = 0
                    potentialEntryPrice = 0
                End If

                If potentialEntryDirection = 1 Then
                    If _inputPayload(runningPayload).High >= potentialEntryPrice AndAlso
                        _inputPayload(runningPayload).Low <= potentialSLPrice Then
                        If _inputPayload(runningPayload).CandleColor = Color.Green Then
                            potentialEntryDirection = 0
                            potentialEntryPrice = 0
                            potentialSLPrice = 0
                        Else
                            signal = 1
                            entryPrice = potentialEntryPrice
                            slPrice = potentialSLPrice
                            targetPrice = entryPrice + (entryPrice - slPrice) * _targetMultiplier
                            potentialEntryDirection = 0
                            potentialEntryPrice = 0
                            potentialSLPrice = 0
                        End If
                    ElseIf _inputPayload(runningPayload).Low <= potentialSLPrice Then
                        potentialEntryDirection = 0
                        potentialEntryPrice = 0
                        potentialSLPrice = 0
                    ElseIf _inputPayload(runningPayload).High >= potentialEntryPrice Then
                        signal = 1
                        entryPrice = potentialEntryPrice
                        slPrice = potentialSLPrice
                        targetPrice = entryPrice + (entryPrice - slPrice) * _targetMultiplier
                        potentialEntryDirection = 0
                        potentialEntryPrice = 0
                        potentialSLPrice = 0
                    ElseIf _nextSignalBreak Then
                        potentialEntryDirection = 0
                        potentialEntryPrice = 0
                        potentialSLPrice = 0
                    End If
                ElseIf potentialEntryDirection = -1 Then
                    If _inputPayload(runningPayload).Low <= potentialEntryPrice AndAlso
                        _inputPayload(runningPayload).High >= potentialSLPrice Then
                        If _inputPayload(runningPayload).CandleColor = Color.Red Then
                            potentialEntryDirection = 0
                            potentialEntryPrice = 0
                            potentialSLPrice = 0
                        Else
                            signal = -1
                            entryPrice = potentialEntryPrice
                            slPrice = potentialSLPrice
                            targetPrice = entryPrice - (slPrice - entryPrice) * _targetMultiplier
                            potentialEntryDirection = 0
                            potentialEntryPrice = 0
                            potentialSLPrice = 0
                        End If
                    ElseIf _inputPayload(runningPayload).High >= potentialSLPrice Then
                        potentialEntryDirection = 0
                        potentialEntryPrice = 0
                        potentialSLPrice = 0
                    ElseIf _inputPayload(runningPayload).Low <= potentialEntryPrice Then
                        signal = -1
                        entryPrice = potentialEntryPrice
                        slPrice = potentialSLPrice
                        targetPrice = entryPrice - (slPrice - entryPrice) * _targetMultiplier
                        potentialEntryDirection = 0
                        potentialEntryPrice = 0
                        potentialSLPrice = 0
                    ElseIf _nextSignalBreak Then
                        potentialEntryDirection = 0
                        potentialEntryPrice = 0
                        potentialSLPrice = 0
                    End If
                End If

                'If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                '_inputPayload(runningPayload).Volume <= _inputPayload(runningPayload).PreviousCandlePayload.Volume * 2 Then
                If _inputPayload(runningPayload).CandleWicks.Bottom >= _inputPayload(runningPayload).CandleRange * 0.5 AndAlso
                    _inputPayload(runningPayload).CandleWicks.Top <= _inputPayload(runningPayload).CandleRange * _minimumCandleWick AndAlso
                    _inputPayload(runningPayload).CandleRange <= _inputPayload(runningPayload).Low * 0.005 AndAlso
                    _inputPayload(runningPayload).CandleRange >= _inputPayload(runningPayload).Low * 0.002 AndAlso
                    _inputPayload(runningPayload).Low <= fractalLowPayload(runningPayload) AndAlso
                    _inputPayload(runningPayload).High >= fractalLowPayload(runningPayload) Then
                    Dim validSignal As Boolean = False
                    If _inputPayload(runningPayload).CandleColor = Color.Green Then
                        If _inputPayload(runningPayload).Open >= fractalLowPayload(runningPayload) Then validSignal = True
                    ElseIf _inputPayload(runningPayload).CandleColor = Color.Red Then
                        If _inputPayload(runningPayload).Close >= fractalLowPayload(runningPayload) Then validSignal = True
                    Else
                        If _inputPayload(runningPayload).Close >= fractalLowPayload(runningPayload) Then validSignal = True
                    End If
                    If validSignal Then
                        potentialEntryDirection = 1
                        potentialEntryPrice = _inputPayload(runningPayload).High
                        potentialEntryPrice += CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                        potentialSLPrice = _inputPayload(runningPayload).Low
                        potentialSLPrice -= CalculateBuffer(potentialSLPrice, RoundOfType.Floor)
                    End If
                ElseIf _inputPayload(runningPayload).CandleWicks.Top >= _inputPayload(runningPayload).CandleRange * 0.5 AndAlso
                    _inputPayload(runningPayload).CandleWicks.Bottom <= _inputPayload(runningPayload).CandleRange * _minimumCandleWick AndAlso
                    _inputPayload(runningPayload).CandleRange <= _inputPayload(runningPayload).High * 0.005 AndAlso
                    _inputPayload(runningPayload).CandleRange >= _inputPayload(runningPayload).High * 0.002 AndAlso
                    _inputPayload(runningPayload).High >= fractalHighPayload(runningPayload) AndAlso
                    _inputPayload(runningPayload).Low <= fractalHighPayload(runningPayload) Then
                    Dim validSignal As Boolean = False
                    If _inputPayload(runningPayload).CandleColor = Color.Green Then
                        If _inputPayload(runningPayload).Close <= fractalHighPayload(runningPayload) Then validSignal = True
                    ElseIf _inputPayload(runningPayload).CandleColor = Color.Red Then
                        If _inputPayload(runningPayload).Open <= fractalHighPayload(runningPayload) Then validSignal = True
                    Else
                        If _inputPayload(runningPayload).Close <= fractalHighPayload(runningPayload) Then validSignal = True
                    End If
                    If validSignal Then
                        potentialEntryDirection = -1
                        potentialEntryPrice = _inputPayload(runningPayload).Low
                        potentialEntryPrice -= CalculateBuffer(potentialEntryPrice, RoundOfType.Floor)
                        potentialSLPrice = _inputPayload(runningPayload).High
                        potentialSLPrice += CalculateBuffer(potentialSLPrice, RoundOfType.Floor)
                    End If
                End If
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
                    'If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                    'outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
                    'If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                    'outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
                    'If outputSupporting3Payload Is Nothing Then outputSupporting3Payload = New Dictionary(Of Date, String)
                    'outputSupporting3Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting3)
                    'If outputSupporting4Payload Is Nothing Then outputSupporting4Payload = New Dictionary(Of Date, String)
                    'outputSupporting4Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting4)
                End If
            Next
            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
            'If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
            'If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
            'If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
            'If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
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
