Imports System.Threading
Imports Algo2TradeBLL

Public Class MarutiRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _stockATR As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal stockATR As Decimal, ByVal lotsize As Integer, ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, tickSize, lotsize, canceller)
        _stockATR = stockATR
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

            Dim HAPayload As Dictionary(Of Date, Payload) = Nothing
            Indicator.HeikenAshi.ConvertToHeikenAshi(_inputPayload, HAPayload)
            Dim potentialHighEntryPrice As Decimal = 0
            Dim potentialLowEntryPrice As Decimal = 0
            Dim tgtPoint As Decimal = _stockATR / 3
            Dim slPoint As Decimal = _stockATR / 6
            Dim tgtMultiplier As Decimal = 2.2 * slPoint / tgtPoint
            Dim slMultiplier As Decimal = tgtPoint / (slPoint * 2.2)
            tgtPoint = tgtPoint * tgtMultiplier
            'slPoint = slPoint * slMultiplier
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing
                Dim supporting2 As String = Nothing
                Dim supporting3 As String = Nothing
                Dim supporting4 As String = Nothing
                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                    signal = 0
                    entryPrice = 0
                    targetPrice = 0
                    slPrice = 0
                    supporting1 = Nothing
                    supporting2 = Nothing
                    supporting3 = Nothing
                    supporting4 = Nothing
                    potentialHighEntryPrice = 0
                    potentialLowEntryPrice = 0
                End If

                If _inputPayload(runningPayload).PayloadDate.Hour = 10 AndAlso
                    _inputPayload(runningPayload).PayloadDate.Minute = 15 Then
                    potentialHighEntryPrice = _inputPayload(runningPayload).PreviousCandlePayload.High * 1.0001
                    potentialLowEntryPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low * 0.9999
                End If
                If potentialHighEntryPrice <> 0 AndAlso potentialLowEntryPrice <> 0 Then
                    If _inputPayload(runningPayload).High >= potentialHighEntryPrice AndAlso
                        _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                        If _inputPayload(runningPayload).CandleColor = Color.Green Then
                            'If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red Then
                            signal = -1
                            entryPrice = potentialLowEntryPrice
                            targetPrice = entryPrice - tgtPoint
                            slPrice = entryPrice + slPoint
                            'End If
                            supporting1 = _stockATR
                            supporting4 = (runningPayload.Hour = 10 AndAlso runningPayload.Minute = 15).ToString
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                        ElseIf _inputPayload(runningPayload).CandleColor = Color.Red Then
                            'If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green Then
                            signal = 1
                            entryPrice = potentialHighEntryPrice
                            targetPrice = entryPrice + tgtPoint
                            slPrice = entryPrice - slPoint
                            'End If
                            supporting1 = _stockATR
                            supporting4 = (runningPayload.Hour = 10 AndAlso runningPayload.Minute = 15).ToString
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                        End If
                    ElseIf _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                        'If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green Then
                        signal = 1
                        entryPrice = potentialHighEntryPrice
                        targetPrice = entryPrice + tgtPoint
                        slPrice = entryPrice - slPoint
                        'End If
                        supporting1 = _stockATR
                        supporting4 = (runningPayload.Hour = 10 AndAlso runningPayload.Minute = 15).ToString
                        potentialHighEntryPrice = 0
                        potentialLowEntryPrice = 0
                    ElseIf _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                        'If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red Then
                        signal = -1
                        entryPrice = potentialLowEntryPrice
                        targetPrice = entryPrice - tgtPoint
                        slPrice = entryPrice + slPoint
                        'End If
                        supporting1 = _stockATR
                        supporting4 = (runningPayload.Hour = 10 AndAlso runningPayload.Minute = 15).ToString
                        potentialHighEntryPrice = 0
                        potentialLowEntryPrice = 0
                    End If
                End If

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    If _inputPayload(runningPayload).PreviousCandlePayload.CandleStrengthNormal = Payload.StrongCandle.Bullish Then
                        supporting2 = "LL"
                    ElseIf _inputPayload(runningPayload).PreviousCandlePayload.CandleStrengthNormal = Payload.StrongCandle.Bearish Then
                        supporting2 = "SS"
                    ElseIf _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green Then
                        supporting2 = "L"
                    ElseIf _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red Then
                        supporting2 = "S"
                    End If
                    If HAPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bullish Then
                        supporting3 = "LL"
                    ElseIf HAPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bearish Then
                        supporting3 = "SS"
                    ElseIf HAPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green Then
                        supporting3 = "L"
                    ElseIf HAPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red Then
                        supporting3 = "S"
                    End If

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
