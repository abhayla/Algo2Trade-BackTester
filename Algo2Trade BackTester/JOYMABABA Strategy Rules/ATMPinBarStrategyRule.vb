Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers

Public Class ATMPinBarStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Public TargetMultiplier As Decimal = 3
    Public CapitalToBeUsed As Decimal = 20000
    Public ATRMultiplier As Decimal = 1.5

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Private ReadOnly _stockType As Trade.TypeOfStock
    Private ReadOnly _signalCandleTime As Date
    Private _firstEntryQuantity As Integer = 0
    Private _usableATR As Decimal = Decimal.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal, ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common,
                   ByVal tradingDate As Date,
                   ByVal timeframe As Integer,
                   ByVal stockType As Trade.TypeOfStock)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
        _tradingDate = tradingDate
        _timeframe = timeframe
        _stockType = stockType
        _signalCandleTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 16, 0)
    End Sub

    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, EntryDetails) = Nothing
            'Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            'Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            'Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            'Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            ''Dim outputModifyStoplossPayload As Dictionary(Of Date, ModifyStoploss) = Nothing
            'Dim outputModifyTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim ATRPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATR.CalculateATR(14, _inputPayload, ATRPayload)

                Dim potentialHighEntryPrice As Decimal = 0
                Dim potentialLowEntryPrice As Decimal = 0
                Dim signalCandle As Payload = Nothing
                Dim signal As Integer = 0
                Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
                For Each runningPayload In _inputPayload.Keys
                    Dim entryData As New EntryDetails
                    With entryData
                        .BuySignal = 0
                        .BuyEntry = 0
                        .BuyStoploss = 0
                        .BuyTarget = 0
                        .BuyQuantity = _quantity
                        .SellSignal = 0
                        .SellEntry = 0
                        .SellStoploss = 0
                        .SellTarget = 0
                        .SellQuantity = _quantity
                        .FirstSignal = 0
                    End With
                    'Dim modifySLPrice As ModifyStoploss = Nothing
                    'Dim modifyTargetPrice As Decimal = 0
                    Dim supporting1 As String = Nothing
                    Dim supporting2 As String = Nothing
                    'Dim supporting3 As String = Nothing
                    'Dim supporting4 As String = Nothing
                    'Dim supporting5 As String = Nothing

                    If runningPayload.Date = _tradingDate.Date Then
                        If _usableATR = Decimal.MinValue Then _usableATR = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                        If signalCandle IsNot Nothing Then
                            If signal = 1 Then
                                If _inputPayload(runningPayload).PreviousCandlePayload.Low < signalCandle.Low Then
                                    signalCandle = Nothing
                                    signal = 0
                                    potentialHighEntryPrice = 0
                                    potentialLowEntryPrice = 0
                                End If
                            ElseIf signal = -1 Then
                                If _inputPayload(runningPayload).PreviousCandlePayload.High > signalCandle.High Then
                                    signalCandle = Nothing
                                    signal = 0
                                    potentialHighEntryPrice = 0
                                    potentialLowEntryPrice = 0
                                End If
                            End If
                        End If
                        If runningPayload > _signalCandleTime Then
                            If IsSignalCandle(_inputPayload(runningPayload).PreviousCandlePayload).Item1 Then
                                signalCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                If IsSignalCandle(_inputPayload(runningPayload).PreviousCandlePayload).Item2 = 1 Then
                                    potentialHighEntryPrice = signalCandle.High + CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                                    signal = 1
                                ElseIf IsSignalCandle(_inputPayload(runningPayload).PreviousCandlePayload).Item2 = -1 Then
                                    potentialLowEntryPrice = signalCandle.Low - CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                                    signal = -1
                                End If
                            End If
                        End If
                        If potentialHighEntryPrice <> 0 AndAlso _inputPayload(runningPayload).High >= potentialHighEntryPrice AndAlso signal = 1 Then
                            entryData.BuySignal = 1
                            entryData.BuyEntry = potentialHighEntryPrice
                            If _inputPayload(runningPayload).Open > potentialHighEntryPrice Then
                                entryData.BuyEntry = _inputPayload(runningPayload).Open
                            End If
                            entryData.BuyStoploss = entryData.BuyEntry - ConvertFloorCeling(_usableATR * ATRMultiplier, _tickSize, RoundOfType.Celing)
                            entryData.BuyTarget = entryData.BuyEntry + ConvertFloorCeling((entryData.BuyEntry - entryData.BuyStoploss) * TargetMultiplier, _tickSize, RoundOfType.Celing)
                            If _firstEntryQuantity = 0 Then
                                Dim capitalRequired As Decimal = entryData.BuyEntry * entryData.BuyQuantity / Strategy.MarginMultiplier
                                If capitalRequired < CapitalToBeUsed Then
                                    entryData.BuyQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                    _firstEntryQuantity = entryData.BuyQuantity
                                End If
                            Else
                                entryData.BuyQuantity = _firstEntryQuantity
                            End If
                            supporting1 = signalCandle.PayloadDate.ToShortTimeString
                            supporting2 = _usableATR

                            signalCandle = Nothing
                            signal = 0
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                        End If
                        If potentialLowEntryPrice <> 0 AndAlso _inputPayload(runningPayload).Low <= potentialLowEntryPrice AndAlso signal = -1 Then
                            entryData.SellSignal = -1
                            entryData.SellEntry = potentialLowEntryPrice
                            If _inputPayload(runningPayload).Open < potentialLowEntryPrice Then
                                entryData.SellEntry = _inputPayload(runningPayload).Open
                            End If
                            entryData.SellStoploss = entryData.SellEntry + ConvertFloorCeling(_usableATR * ATRMultiplier, _tickSize, RoundOfType.Celing)
                            entryData.SellTarget = entryData.SellEntry - ConvertFloorCeling((entryData.SellStoploss - entryData.SellEntry) * TargetMultiplier, _tickSize, RoundOfType.Celing)
                            If _firstEntryQuantity = 0 Then
                                Dim capitalRequired As Decimal = entryData.SellEntry * entryData.SellQuantity / Strategy.MarginMultiplier
                                If capitalRequired < CapitalToBeUsed Then
                                    entryData.SellQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                End If
                                _firstEntryQuantity = entryData.SellQuantity
                            Else
                                entryData.SellQuantity = _firstEntryQuantity
                            End If
                            supporting1 = signalCandle.PayloadDate.ToShortTimeString
                            supporting2 = _usableATR

                            signalCandle = Nothing
                            signal = 0
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                        End If
                    End If

                    If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                        If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, EntryDetails)
                        outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryData)

                        If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                        outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
                        If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                        outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
                        'If outputSupporting3Payload Is Nothing Then outputSupporting3Payload = New Dictionary(Of Date, String)
                        'outputSupporting3Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting3)
                        'If outputSupporting4Payload Is Nothing Then outputSupporting4Payload = New Dictionary(Of Date, String)
                        'outputSupporting4Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting4)
                        'If outputSupporting5Payload Is Nothing Then outputSupporting5Payload = New Dictionary(Of Date, String)
                        'outputSupporting5Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting5)
                    End If
                Next

                If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
                If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
                If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
                If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
                If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
                If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
                If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
            End If
        End If
    End Sub

    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, Integer)
        Dim ret As Tuple(Of Boolean, Integer) = New Tuple(Of Boolean, Integer)(False, 0)
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If (candle.CandleWicks.Top / candle.CandleRange) * 100 >= 66 Then
                If candle.High > candle.PreviousCandlePayload.High AndAlso
                   candle.Low > candle.PreviousCandlePayload.Low Then
                    ret = New Tuple(Of Boolean, Integer)(True, -1)
                End If
            ElseIf (candle.CandleWicks.Bottom / candle.CandleRange) * 100 >= 66 Then
                If candle.Low < candle.PreviousCandlePayload.Low AndAlso
                   candle.High < candle.PreviousCandlePayload.High Then
                    ret = New Tuple(Of Boolean, Integer)(True, 1)
                End If
            End If
        End If
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