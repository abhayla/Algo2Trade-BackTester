Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Public Class ReverseVolumeHKRule
    Inherits StrategyRule
    Implements IDisposable
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
            Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing

            Dim ATRPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.ATR.CalculateATR(14, _inputPayload, ATRPayload)
            Dim potentialHighEntry As Decimal = 0
            Dim potentialLowEntry As Decimal = 0
            Dim signalCandleColor As Color = Color.White
            Dim supporting1 As String = Nothing
            Dim supporting2 As String = Nothing
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                    potentialHighEntry = 0
                    potentialLowEntry = 0
                End If

                If potentialHighEntry <> 0 AndAlso potentialLowEntry <> 0 Then
                    If _inputPayload(runningPayload).High >= potentialHighEntry AndAlso
                        _inputPayload(runningPayload).Low <= potentialLowEntry Then
                        If _inputPayload(runningPayload).CandleColor = Color.Green Then
                            If signalCandleColor = Color.Green Then
                                signal = -1
                                entryPrice = potentialLowEntry
                                slPrice = potentialHighEntry
                                targetPrice = entryPrice - (slPrice - entryPrice) * 2
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            Else
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            End If
                        ElseIf _inputPayload(runningPayload).CandleColor = Color.Red Then
                            If signalCandleColor = Color.Red Then
                                signal = 1
                                entryPrice = potentialHighEntry
                                slPrice = potentialLowEntry
                                targetPrice = entryPrice + (entryPrice - slPrice) * 2
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            Else
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            End If
                        End If
                    ElseIf _inputPayload(runningPayload).High >= potentialHighEntry Then
                        If signalCandleColor = Color.Red Then
                            signal = 1
                            entryPrice = potentialHighEntry
                            slPrice = potentialLowEntry
                            targetPrice = entryPrice + (entryPrice - slPrice) * 2
                            potentialHighEntry = 0
                            potentialLowEntry = 0
                        Else
                            potentialHighEntry = 0
                            potentialLowEntry = 0
                        End If
                    ElseIf _inputPayload(runningPayload).Low <= potentialLowEntry Then
                        If signalCandleColor = Color.Green Then
                            signal = -1
                            entryPrice = potentialLowEntry
                            slPrice = potentialHighEntry
                            targetPrice = entryPrice - (slPrice - entryPrice) * 2
                            potentialHighEntry = 0
                            potentialLowEntry = 0
                        Else
                            potentialHighEntry = 0
                            potentialLowEntry = 0
                        End If
                    End If
                End If

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    If _inputPayload(runningPayload).CandleColor <> _inputPayload(runningPayload).VolumeColor Then
                        If IsEngulfingCandle(_inputPayload(runningPayload)) Then
                            If _inputPayload(runningPayload).CandleRange <= ATRPayload(runningPayload) AndAlso
                            _inputPayload(runningPayload).CandleRange >= ATRPayload(runningPayload) / 2 Then
                                potentialHighEntry = _inputPayload(runningPayload).High
                                potentialHighEntry += CalculateBuffer(potentialHighEntry, RoundOfType.Floor)
                                potentialLowEntry = _inputPayload(runningPayload).Low
                                potentialLowEntry -= CalculateBuffer(potentialLowEntry, RoundOfType.Floor)
                                signalCandleColor = _inputPayload(runningPayload).CandleColor
                                supporting1 = Math.Round(ATRPayload(runningPayload), 2)
                                supporting2 = _inputPayload(runningPayload).CandleColor.ToString
                            End If
                        ElseIf IsInsideBar(_inputPayload(runningPayload)) Then
                            If _inputPayload(runningPayload).PreviousCandlePayload.CandleRange <= ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            _inputPayload(runningPayload).PreviousCandlePayload.CandleRange >= ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) / 2 Then
                                If _inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bearish OrElse
                                    _inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bullish Then
                                    potentialHighEntry = _inputPayload(runningPayload).High
                                    potentialHighEntry += CalculateBuffer(potentialHighEntry, RoundOfType.Floor)
                                    potentialLowEntry = _inputPayload(runningPayload).Low
                                    potentialLowEntry -= CalculateBuffer(potentialLowEntry, RoundOfType.Floor)
                                    signalCandleColor = _inputPayload(runningPayload).CandleColor
                                    supporting1 = Math.Round(ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2)
                                    supporting2 = _inputPayload(runningPayload).CandleColor.ToString
                                Else
                                    potentialHighEntry = _inputPayload(runningPayload).PreviousCandlePayload.High
                                    potentialHighEntry += CalculateBuffer(potentialHighEntry, RoundOfType.Floor)
                                    potentialLowEntry = _inputPayload(runningPayload).PreviousCandlePayload.Low
                                    potentialLowEntry -= CalculateBuffer(potentialLowEntry, RoundOfType.Floor)
                                    signalCandleColor = _inputPayload(runningPayload).PreviousCandlePayload.CandleColor
                                    supporting1 = Math.Round(ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2)
                                    supporting2 = _inputPayload(runningPayload).PreviousCandlePayload.CandleColor.ToString
                                End If
                            End If
                        End If
                    End If
                End If

                If outputSignalPayload IsNot Nothing AndAlso outputSignalPayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate) AndAlso
                    outputSignalPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate) = 1 Then
                    If _inputPayload(runningPayload).PreviousCandlePayload.Volume < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume Then
                        slPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low
                        slPrice -= CalculateBuffer(slPrice, RoundOfType.Floor)
                    End If
                ElseIf outputSignalPayload IsNot Nothing AndAlso outputSignalPayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate) AndAlso
                    outputSignalPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate) = -1 Then
                    If _inputPayload(runningPayload).PreviousCandlePayload.Volume < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume Then
                        slPrice = _inputPayload(runningPayload).PreviousCandlePayload.High
                        slPrice += CalculateBuffer(slPrice, RoundOfType.Floor)
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
                    If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                    outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
                End If
            Next
            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
            If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
            If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
        End If
    End Sub
    Private Function IsInsideBar(ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        ret = currentCandle.High <= currentCandle.PreviousCandlePayload.High AndAlso currentCandle.Low >= currentCandle.PreviousCandlePayload.Low
        Return ret
    End Function
    Private Function IsEngulfingCandle(ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        ret = currentCandle.High >= currentCandle.PreviousCandlePayload.High AndAlso currentCandle.Low <= currentCandle.PreviousCandlePayload.Low
        Return ret
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
