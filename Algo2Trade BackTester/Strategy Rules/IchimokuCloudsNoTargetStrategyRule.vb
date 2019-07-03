Imports System.Threading
Imports Algo2TradeBLL

Public Class IchimokuCloudsNoTargetStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _minStoplossPercentage As Decimal = 0.03
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
            Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            Dim conversionLinePayload As Dictionary(Of Date, Decimal) = Nothing
            Dim baseLinePayload As Dictionary(Of Date, Decimal) = Nothing
            Dim leadingSpanAPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim leadingSpanBPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim laggingSpanPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.IchimokuClouds.CalculateIchimokuClouds(9, 26, 52, 26, _inputPayload, conversionLinePayload, baseLinePayload, leadingSpanAPayload, leadingSpanBPayload, laggingSpanPayload)

            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim quantity As Integer = Strategy.CalculateQuantityFromInvestment(_quantity, 5000, _inputPayload.LastOrDefault.Value.Open, Trade.TypeOfStock.Cash)
                Dim supporting1 As String = Nothing
                Dim supporting2 As String = Nothing
                Dim supporting3 As String = Nothing
                Dim supporting4 As String = Nothing
                Dim supporting5 As String = Nothing

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    If leadingSpanAPayload IsNot Nothing AndAlso leadingSpanAPayload.Count > 0 AndAlso
                        _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                        leadingSpanAPayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate) Then
                        slPrice = leadingSpanAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate)
                    End If

                    If conversionLinePayload IsNot Nothing AndAlso conversionLinePayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                        baseLinePayload IsNot Nothing AndAlso baseLinePayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                        conversionLinePayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = baseLinePayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                        Dim leadingSpanA As KeyValuePair(Of Date, Decimal)? = Common.GetPayloadAt(leadingSpanAPayload, _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, 26)
                        Dim leadingSpanB As KeyValuePair(Of Date, Decimal)? = Common.GetPayloadAt(leadingSpanBPayload, _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, 26)
                        If leadingSpanA IsNot Nothing AndAlso leadingSpanB IsNot Nothing AndAlso
                            leadingSpanAPayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            leadingSpanBPayload.ContainsKey(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                            If leadingSpanA.Value.Value = leadingSpanB.Value.Value Then
                                Dim laggingSpan As KeyValuePair(Of Date, Payload)? = Common.GetPayloadAt(_inputPayload, _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, -26)
                                If laggingSpan IsNot Nothing Then
                                    Dim potentialEntryPrice As Decimal = _inputPayload(runningPayload).Open
                                    Dim potentialSLPrice As Decimal = leadingSpanAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                                    If _inputPayload(runningPayload).PreviousCandlePayload.Close > laggingSpan.Value.Value.Close Then
                                        If (potentialEntryPrice - potentialSLPrice) > potentialEntryPrice * _minStoplossPercentage / 100 AndAlso
                                            _inputPayload(runningPayload).Low > leadingSpanAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                            _inputPayload(runningPayload).Low > leadingSpanBPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                            signal = 1
                                            entryPrice = potentialEntryPrice
                                            targetPrice = entryPrice + 10000
                                            slPrice = potentialSLPrice
                                        End If
                                    ElseIf _inputPayload(runningPayload).PreviousCandlePayload.Close < laggingSpan.Value.Value.Close Then
                                        If (potentialSLPrice - potentialEntryPrice) > potentialEntryPrice * _minStoplossPercentage / 100 AndAlso
                                            _inputPayload(runningPayload).High < leadingSpanAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                            _inputPayload(runningPayload).High < leadingSpanBPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                            signal = -1
                                            entryPrice = potentialEntryPrice
                                            targetPrice = entryPrice - 10000
                                            slPrice = potentialSLPrice
                                        End If
                                    End If
                                    supporting1 = String.Format("Conversion: {0}", conversionLinePayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate))
                                    supporting2 = String.Format("Base: {0}", baseLinePayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate))
                                    supporting3 = String.Format("Span A: {0}, {1}", leadingSpanA.Value.Key, leadingSpanA.Value.Value)
                                    supporting4 = String.Format("Span B: {0}, {1}", leadingSpanB.Value.Key, leadingSpanB.Value.Value)
                                    supporting5 = String.Format("Pre Close: {0}, {1}", laggingSpan.Value.Key, laggingSpan.Value.Value.Close)
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
                    If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                    outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
                    If outputSupporting3Payload Is Nothing Then outputSupporting3Payload = New Dictionary(Of Date, String)
                    outputSupporting3Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting3)
                    If outputSupporting4Payload Is Nothing Then outputSupporting4Payload = New Dictionary(Of Date, String)
                    outputSupporting4Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting4)
                    If outputSupporting5Payload Is Nothing Then outputSupporting5Payload = New Dictionary(Of Date, String)
                    outputSupporting5Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting5)
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
            If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
            If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
            If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
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