Imports Algo2TradeBLL
Imports System.Threading

Public Class PSAR_EMAStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Private ReadOnly _allowableDistancePercentageFromPSAR = 0.32
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal, ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common,
                   ByVal tradingDate As Date,
                   ByVal timeframe As Integer)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
        _tradingDate = tradingDate
        _timeframe = timeframe
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
            'Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            Dim EMAPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.EMA.CalculateEMA(200, Payload.PayloadFields.Close, _inputPayload, EMAPayload)
            Dim PSARPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim TrendPayload As Dictionary(Of Date, Color) = Nothing
            Indicator.ParabolicSAR.CalculatePSAR(0.02, 0.2, _inputPayload, PSARPayload, TrendPayload)

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim firstCandleOfTradingDate As Boolean = True
                Dim changeTime As Boolean = False
                For Each runningPayload In _inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim slPrice As Decimal = 0
                    Dim targetPrice As Decimal = 0
                    Dim quantity As Integer = _quantity
                    Dim originalTime As Date = Date.MinValue
                    Dim supporting1 As String = Nothing
                    Dim supporting2 As String = Nothing
                    Dim supporting3 As String = Nothing
                    Dim supporting4 As String = Nothing
                    'Dim supporting5 As String = Nothing
                    If runningPayload.Date = _tradingDate.Date Then
                        supporting1 = PSARPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                        supporting2 = TrendPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).ToString

                        If _inputPayload(runningPayload).PreviousCandlePayload.Low > EMAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            TrendPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = Color.Green Then
                            Dim diff As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.Low - PSARPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                            supporting3 = diff
                            supporting4 = _inputPayload(runningPayload).PreviousCandlePayload.Low * _allowableDistancePercentageFromPSAR / 100
                            If diff <= _inputPayload(runningPayload).PreviousCandlePayload.Low * _allowableDistancePercentageFromPSAR / 100 Then
                                signal = 1
                                entryPrice = _inputPayload(runningPayload).Open
                                slPrice = PSARPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                                targetPrice = entryPrice + 1000000
                            End If
                        ElseIf _inputPayload(runningPayload).PreviousCandlePayload.High < EMAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            TrendPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = Color.Red Then
                            Dim diff As Decimal = PSARPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) - _inputPayload(runningPayload).PreviousCandlePayload.High
                            supporting3 = diff
                            supporting4 = _inputPayload(runningPayload).PreviousCandlePayload.High * _allowableDistancePercentageFromPSAR / 100
                            If diff <= _inputPayload(runningPayload).PreviousCandlePayload.High * _allowableDistancePercentageFromPSAR / 100 Then
                                signal = -1
                                entryPrice = _inputPayload(runningPayload).Open
                                slPrice = PSARPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                                targetPrice = entryPrice - 1000000
                            End If
                        End If

                        slPrice = PSARPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate)

                        If firstCandleOfTradingDate Then
                            originalTime = _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate = runningPayload.AddMinutes(_timeframe * -1)
                            firstCandleOfTradingDate = False
                            changeTime = True
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
                        'If outputSupporting5Payload Is Nothing Then outputSupporting5Payload = New Dictionary(Of Date, String)
                        'outputSupporting5Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting5)
                        If changeTime Then
                            _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate = originalTime
                            changeTime = False
                        End If
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
                'If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
            End If
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
