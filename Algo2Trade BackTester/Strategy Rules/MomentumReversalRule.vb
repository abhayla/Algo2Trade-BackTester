Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Public Class MomentumReversalRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _targetMultiplier As Decimal = 3
    Private ReadOnly _candleWickSizePercentage As Decimal = 30
    Private ReadOnly _minimumCandleRangePercentage As Decimal = 0.1
    Private ReadOnly _maxStoplossPercentage As Decimal = 30

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
            'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing

            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim potentialEntryPrice As Decimal = 0
                Dim supporting1 As String = Nothing
                'Dim supporting2 As String = Nothing

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    Dim benchMarkWicksSize As Double = _inputPayload(runningPayload).PreviousCandlePayload.CandleRange * _candleWickSizePercentage / 100
                    If _minimumCandleRangePercentage < _inputPayload(runningPayload).PreviousCandlePayload.CandleRangePercentage Then
                        If _inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Top > benchMarkWicksSize Then
                            potentialEntryPrice = _inputPayload(runningPayload).PreviousCandlePayload.High
                            potentialEntryPrice += CalculateBuffer(potentialEntryPrice, RoundOfType.Celing)
                            If _inputPayload(runningPayload).High >= potentialEntryPrice Then
                                signal = 1
                                entryPrice = potentialEntryPrice
                                slPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low
                                slPrice -= CalculateBuffer(slPrice, RoundOfType.Celing)
                                targetPrice = entryPrice + ((entryPrice - slPrice)) * _targetMultiplier
                                Dim pl As Double = Strategy.CalculatePL(_inputPayload(runningPayload).TradingSymbol, entryPrice, slPrice, quantity, 1, Trade.TypeOfStock.Cash)
                                Dim capitalRequiredWithMargin As Decimal = entryPrice * quantity / Strategy.MarginMultiplier
                                If Math.Abs(pl) > capitalRequiredWithMargin * _maxStoplossPercentage / 100 Then
                                    slPrice = entryPrice - ((capitalRequiredWithMargin * _maxStoplossPercentage / 100) / quantity)
                                    supporting1 = "Stop Loss moved because of Capital preservation"
                                Else
                                    supporting1 = "Original"
                                End If
                            End If
                        ElseIf _inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Bottom > benchMarkWicksSize Then
                            potentialEntryPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low
                            potentialEntryPrice -= CalculateBuffer(potentialEntryPrice, RoundOfType.Celing)
                            If _inputPayload(runningPayload).Low <= potentialEntryPrice Then
                                signal = -1
                                entryPrice = potentialEntryPrice
                                slPrice = _inputPayload(runningPayload).PreviousCandlePayload.High
                                slPrice += CalculateBuffer(slPrice, RoundOfType.Celing)
                                targetPrice = entryPrice - ((slPrice - entryPrice)) * _targetMultiplier
                                Dim pl As Double = Strategy.CalculatePL(_inputPayload(runningPayload).TradingSymbol, slPrice, entryPrice, quantity, 1, Trade.TypeOfStock.Cash)
                                Dim capitalRequiredWithMargin As Decimal = entryPrice * quantity / Strategy.MarginMultiplier
                                If Math.Abs(pl) > capitalRequiredWithMargin * _maxStoplossPercentage / 100 Then
                                    slPrice = entryPrice + ((capitalRequiredWithMargin * _maxStoplossPercentage / 100) / quantity)
                                    supporting1 = "Stop Loss moved because of Capital preservation"
                                Else
                                    supporting1 = "Original"
                                End If
                            End If
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
            If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
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
