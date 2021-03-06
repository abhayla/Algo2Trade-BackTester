﻿Imports System.Threading
Imports Algo2TradeBLL

Public Class OutsideFractalsRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _target As Decimal = 0.05
    Private ReadOnly _maxStoploss As Decimal = 0.055
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal,
                   ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource)
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

            Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.CalculateFractalBands(_inputPayload, fractalHighPayload, fractalLowPayload)

            Dim potentialSignalCandle As Payload = Nothing
            Dim currentTradeCandle As Payload = Nothing
            Dim potentialSignal As Integer = 0
            Dim currentSignal As Integer = 0
            Dim checkingFractal As Decimal = 0
            Dim potentialEntryPrice As Decimal = 0
            Dim onlyOnce As Boolean = False
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing
                'Dim supporting2 As String = Nothing

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    runningPayload.Date <> _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date Then
                    potentialSignal = 0
                    currentSignal = 0
                End If

                'Changing Signal
                If currentSignal = 1 Then
                    If runningPayload > currentTradeCandle.PayloadDate Then
                        If onlyOnce AndAlso fractalHighPayload(runningPayload) > fractalHighPayload(currentTradeCandle.PayloadDate) Then
                            onlyOnce = False
                            checkingFractal = fractalHighPayload(runningPayload)
                        End If
                        If checkingFractal <> 0 AndAlso checkingFractal = fractalHighPayload(runningPayload) Then
                            If _inputPayload(runningPayload).Low < (fractalLowPayload(runningPayload) - CalculateBuffer(1)) Then
                                Dim stoploss As Decimal = fractalHighPayload(runningPayload) + CalculateBuffer(1)
                                If Math.Abs(stoploss - (fractalLowPayload(runningPayload) - CalculateBuffer(1))) <= _maxStoploss Then
                                    signal = -1
                                    entryPrice = fractalLowPayload(runningPayload) - CalculateBuffer(1)
                                    slPrice = stoploss
                                    targetPrice = entryPrice - _target
                                    supporting1 = "Wrong trade entry"
                                End If
                            End If
                        End If
                    End If
                ElseIf currentSignal = -1 Then
                    If runningPayload > currentTradeCandle.PayloadDate Then
                        If onlyOnce AndAlso fractalLowPayload(runningPayload) > fractalLowPayload(currentTradeCandle.PayloadDate) Then
                            onlyOnce = False
                            checkingFractal = fractalLowPayload(runningPayload)
                        End If
                        If checkingFractal <> 0 AndAlso checkingFractal = fractalLowPayload(runningPayload) Then
                            If _inputPayload(runningPayload).High < (fractalHighPayload(runningPayload) + CalculateBuffer(1)) Then
                                Dim stoploss As Decimal = fractalLowPayload(runningPayload) - CalculateBuffer(1)
                                If Math.Abs((fractalHighPayload(runningPayload) + CalculateBuffer(1)) - stoploss) <= _maxStoploss Then
                                    signal = 1
                                    entryPrice = fractalHighPayload(runningPayload) + CalculateBuffer(1)
                                    slPrice = stoploss
                                    targetPrice = entryPrice + _target
                                    supporting1 = "Wrong trade entry"
                                End If
                            End If
                        End If
                    End If
                End If

                'Taking trade
                If potentialSignal = -1 Then
                    If fractalLowPayload(runningPayload) > potentialEntryPrice Then
                        potentialEntryPrice = fractalLowPayload(runningPayload)
                    End If
                    Dim potentialEntry As Decimal = potentialEntryPrice - CalculateBuffer(1)
                    If _inputPayload(runningPayload).Low <= potentialEntry Then
                        Dim stoploss As Decimal = fractalHighPayload.Max(Function(x)
                                                                             If x.Key >= potentialSignalCandle.PayloadDate AndAlso x.Key < runningPayload Then
                                                                                 Return x.Value
                                                                             Else
                                                                                 Return Decimal.MinValue
                                                                             End If
                                                                         End Function) + CalculateBuffer(1)
                        If Math.Abs(stoploss - potentialEntry) <= _maxStoploss Then
                            signal = -1
                            entryPrice = potentialEntry
                            slPrice = stoploss
                            targetPrice = entryPrice - _target
                            supporting1 = potentialSignalCandle.PayloadDate
                            currentSignal = signal
                            currentTradeCandle = _inputPayload(runningPayload)
                            onlyOnce = True
                        End If
                        potentialSignal = 0
                    End If
                ElseIf potentialSignal = 1 Then
                    If fractalHighPayload(runningPayload) < potentialEntryPrice Then
                        potentialEntryPrice = fractalHighPayload(runningPayload)
                    End If
                    Dim potentialEntry As Decimal = potentialEntryPrice + CalculateBuffer(1)
                    If _inputPayload(runningPayload).High >= potentialEntry Then
                        Dim stoploss As Decimal = fractalLowPayload.Min(Function(x)
                                                                            If x.Key >= potentialSignalCandle.PayloadDate AndAlso x.Key < runningPayload Then
                                                                                Return x.Value
                                                                            Else
                                                                                Return Decimal.MaxValue
                                                                            End If
                                                                        End Function) - CalculateBuffer(1)
                        If Math.Abs(potentialEntry - stoploss) <= _maxStoploss Then
                            signal = 1
                            entryPrice = potentialEntry
                            slPrice = stoploss
                            targetPrice = entryPrice + _target
                            supporting1 = potentialSignalCandle.PayloadDate
                            currentSignal = signal
                            currentTradeCandle = _inputPayload(runningPayload)
                            onlyOnce = True
                        End If
                        potentialSignal = 0
                    End If
                End If

                'Cancelling Signal
                If potentialSignal = -1 Then
                    If _inputPayload(runningPayload).Close > potentialSignalCandle.High AndAlso _inputPayload(runningPayload).Low < fractalHighPayload(runningPayload) Then
                        potentialSignal = 0
                    End If
                ElseIf potentialSignal = 1 Then
                    If _inputPayload(runningPayload).Close < potentialSignalCandle.Low AndAlso _inputPayload(runningPayload).High > fractalLowPayload(runningPayload) Then
                        potentialSignal = 0
                    End If
                End If

                'Signal
                If potentialSignal = -1 Then
                    If _inputPayload(runningPayload).Low > fractalHighPayload(runningPayload) Then
                        If fractalHighPayload(runningPayload) = fractalHighPayload(potentialSignalCandle.PayloadDate) Then
                            If _inputPayload(runningPayload).High >= potentialSignalCandle.High Then
                                potentialSignalCandle = _inputPayload(runningPayload)
                                potentialSignal = -1
                                potentialEntryPrice = fractalLowPayload(runningPayload)
                            End If
                        Else
                            potentialSignalCandle = _inputPayload(runningPayload)
                            potentialSignal = -1
                            potentialEntryPrice = fractalLowPayload(runningPayload)
                        End If
                    End If
                ElseIf potentialSignal = 1 Then
                    If _inputPayload(runningPayload).High < fractalLowPayload(runningPayload) Then
                        If fractalLowPayload(runningPayload) = fractalLowPayload(potentialSignalCandle.PayloadDate) Then
                            If _inputPayload(runningPayload).Low <= potentialSignalCandle.Low Then
                                potentialSignalCandle = _inputPayload(runningPayload)
                                potentialSignal = 1
                                potentialEntryPrice = fractalHighPayload(runningPayload)
                            End If
                        Else
                            potentialSignalCandle = _inputPayload(runningPayload)
                            potentialSignal = 1
                            potentialEntryPrice = fractalHighPayload(runningPayload)
                        End If
                    End If
                Else
                    If _inputPayload(runningPayload).Low > fractalHighPayload(runningPayload) Then
                        potentialSignalCandle = _inputPayload(runningPayload)
                        potentialSignal = -1
                        potentialEntryPrice = fractalLowPayload(runningPayload)
                    ElseIf _inputPayload(runningPayload).High < fractalLowPayload(runningPayload) Then
                        potentialSignalCandle = _inputPayload(runningPayload)
                        potentialSignal = 1
                        potentialEntryPrice = fractalHighPayload(runningPayload)
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
