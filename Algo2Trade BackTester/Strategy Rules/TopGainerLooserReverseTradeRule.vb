Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Public Class TopGainerLooserReverseTrade
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _entryDirection As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal quantity As Integer, ByVal canceller As CancellationTokenSource, ByVal entryDirection As Integer)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _entryDirection = entryDirection
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

            Dim highATRBandPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim lowATRBandPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.ATRBands.CalculateATRBands(3, 14, Payload.PayloadFields.Close, _inputPayload, highATRBandPayload, lowATRBandPayload)
            'Dim tradingDate As Date = _inputPayload.Keys.LastOrDefault.Date

            Dim potentialEntry As Decimal = 0
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                'Dim supporting1 As String = Nothing
                'Dim supporting2 As String = Nothing
                Dim signalTime As Date = New Date(runningPayload.Year, runningPayload.Month, runningPayload.Day, 9, 20, 0)

                If runningPayload > signalTime AndAlso _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                    If _entryDirection = 1 AndAlso _inputPayload(runningPayload).High > _inputPayload(runningPayload).PreviousCandlePayload.High Then
                        Dim previousLow As Decimal = _inputPayload.Min(Function(x)
                                                                           If x.Key.Date = runningPayload.Date AndAlso
                                                                             x.Key < runningPayload Then
                                                                               Return x.Value.Low
                                                                           Else
                                                                               Return Decimal.MaxValue
                                                                           End If
                                                                       End Function)
                        entryPrice = _inputPayload(runningPayload).PreviousCandlePayload.High
                        slPrice = previousLow
                        If Math.Abs(((entryPrice / slPrice) - 1) * 100) <= 1 Then
                            targetPrice = Math.Round(highATRBandPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate), 2)
                            If (targetPrice - entryPrice) > (entryPrice - slPrice) AndAlso
                                Math.Abs(((targetPrice / entryPrice) - 1) * 100) <= 1.5 Then
                                signal = 1
                            Else
                                entryPrice = 0
                                slPrice = 0
                                targetPrice = 0
                            End If
                        Else
                            entryPrice = 0
                            slPrice = 0
                        End If
                    ElseIf _entryDirection = -1 AndAlso _inputPayload(runningPayload).Low < _inputPayload(runningPayload).PreviousCandlePayload.Low Then
                        Dim previousHigh As Decimal = _inputPayload.Max(Function(x)
                                                                            If x.Key.Date = runningPayload.Date AndAlso
                                                                             x.Key < runningPayload Then
                                                                                Return x.Value.High
                                                                            Else
                                                                                Return Decimal.MinValue
                                                                            End If
                                                                        End Function)
                        entryPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low
                        slPrice = previousHigh
                        If Math.Abs(((slPrice / entryPrice) - 1) * 100) <= 1 Then
                            targetPrice = Math.Round(lowATRBandPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate), 2)
                            If (entryPrice - targetPrice) > (slPrice - entryPrice) AndAlso
                                Math.Abs(((entryPrice / targetPrice) - 1) * 100) <= 1.5 Then
                                signal = -1
                            Else
                                entryPrice = 0
                                slPrice = 0
                                targetPrice = 0
                            End If
                        Else
                            entryPrice = 0
                            slPrice = 0
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
                    'If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                    'outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
                    'If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                    'outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
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
