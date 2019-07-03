Imports System.Threading
Imports Algo2TradeBLL

Public Class _67CrudeOilLevelBasedStrategy
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _targetMultiplier As Decimal = 1
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

            Dim averagePrice As Decimal = 0
            Dim potentialHighEntryPrice As Decimal = 0
            Dim potentialLowEntryPrice As Decimal = 0
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing
                Dim supporting2 As String = Nothing

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    runningPayload.Date <> _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate Then
                    averagePrice = 0
                    potentialHighEntryPrice = 0
                    potentialLowEntryPrice = 0
                End If
                Dim signalCheckTime As Date = New Date(runningPayload.Year, runningPayload.Month, runningPayload.Day, 17, 0, 0)
                If runningPayload = signalCheckTime Then
                    Dim maxHigh As Decimal = _inputPayload.Max(Function(x)
                                                                   If x.Key.Date = signalCheckTime.Date AndAlso x.Key < signalCheckTime Then
                                                                       Return x.Value.High
                                                                   Else
                                                                       Return Decimal.MinValue
                                                                   End If
                                                               End Function)
                    Dim minLow As Decimal = _inputPayload.Min(Function(x)
                                                                  If x.Key.Date = signalCheckTime.Date AndAlso x.Key < signalCheckTime Then
                                                                      Return x.Value.Low
                                                                  Else
                                                                      Return Decimal.MaxValue
                                                                  End If
                                                              End Function)

                    If maxHigh <> Decimal.MinValue AndAlso maxHigh <> Decimal.MinValue Then
                        supporting1 = maxHigh
                        supporting2 = minLow
                        averagePrice = Utilities.Numbers.ConvertFloorCeling(((maxHigh + minLow) / 2), _tickSize, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                        Dim subset As Decimal = Utilities.Numbers.ConvertFloorCeling(((maxHigh - minLow) / 4), _tickSize, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                        potentialHighEntryPrice = averagePrice + subset
                        potentialLowEntryPrice = averagePrice - subset
                    End If
                End If

                    If _inputPayload(runningPayload).CandleColor = Color.Green Then
                    If potentialLowEntryPrice <> 0 Then
                        If _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                            signal = -1
                            entryPrice = potentialLowEntryPrice
                            slPrice = averagePrice
                            targetPrice = entryPrice - ((averagePrice - entryPrice) * _targetMultiplier)
                            potentialLowEntryPrice = 0
                            potentialHighEntryPrice = 0
                        End If
                    ElseIf potentialHighEntryPrice <> 0 Then
                        If _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                            signal = 1
                            entryPrice = potentialHighEntryPrice
                            slPrice = averagePrice
                            targetPrice = entryPrice + ((entryPrice - averagePrice) * _targetMultiplier)
                            potentialLowEntryPrice = 0
                            potentialHighEntryPrice = 0
                        End If
                    End If
                ElseIf _inputPayload(runningPayload).CandleColor = Color.Red Then
                    If potentialHighEntryPrice <> 0 Then
                        If _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                            signal = 1
                            entryPrice = potentialHighEntryPrice
                            slPrice = averagePrice
                            targetPrice = entryPrice + ((entryPrice - averagePrice) * _targetMultiplier)
                            potentialLowEntryPrice = 0
                            potentialHighEntryPrice = 0
                        End If
                    ElseIf potentialLowEntryPrice <> 0 Then
                        If _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                            signal = -1
                            entryPrice = potentialLowEntryPrice
                            slPrice = averagePrice
                            targetPrice = entryPrice - ((averagePrice - entryPrice) * _targetMultiplier)
                            potentialLowEntryPrice = 0
                            potentialHighEntryPrice = 0
                        End If
                    End If
                Else
                    If potentialHighEntryPrice <> 0 Then
                        If _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                            signal = 1
                            entryPrice = potentialHighEntryPrice
                            slPrice = averagePrice
                            targetPrice = entryPrice + ((entryPrice - averagePrice) * _targetMultiplier)
                            potentialLowEntryPrice = 0
                            potentialHighEntryPrice = 0
                        End If
                    ElseIf potentialLowEntryPrice <> 0 Then
                        If _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                            signal = -1
                            entryPrice = potentialLowEntryPrice
                            slPrice = averagePrice
                            targetPrice = entryPrice - ((averagePrice - entryPrice) * _targetMultiplier)
                            potentialLowEntryPrice = 0
                            potentialHighEntryPrice = 0
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