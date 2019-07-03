Imports Algo2TradeBLL
Imports System.Threading

Public Class TrendReversalStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Private ReadOnly _mtmProfitPerday As Decimal
    Private ReadOnly _stockPrice As Decimal
    Private ReadOnly _maxTargetPerTrade As Decimal = 1500
    Private ReadOnly _maxStoplossPerTrade As Decimal = -500

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal, ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common,
                   ByVal tradingDate As Date,
                   ByVal timeframe As Integer,
                   ByVal mtmProfitPerDay As Decimal,
                   ByVal stockPrice As Decimal)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
        _tradingDate = tradingDate
        _timeframe = timeframe
        _mtmProfitPerday = mtmProfitPerDay
        _stockPrice = stockPrice
    End Sub

    'Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
    '    If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
    '        Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
    '        Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
    '        'Dim outputModifyStoplossPayload As Dictionary(Of Date, ModifyStoploss) = Nothing
    '        'Dim outputModifyTargetPayload As Dictionary(Of Date, Decimal) = Nothing
    '        'Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

    '        Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Indicator.VWAP.CalculateVWAP(_inputPayload, VWAPPayload)

    '        If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
    '            Dim startCandleForTheTradingDay As Boolean = True
    '            Dim lastTrend As Integer = 0
    '            Dim candleCount As Integer = 1
    '            For Each runningPayload In _inputPayload.Keys
    '                Dim signal As Integer = 0
    '                Dim entryPrice As Decimal = 0
    '                Dim slPrice As Decimal = 0
    '                Dim targetPrice As Decimal = 0
    '                Dim quantity As Integer = Strategy.CalculateQuantityFromInvestment(_quantity, 20000, _stockPrice, Trade.TypeOfStock.Cash)

    '                'Dim modifySLPrice As ModifyStoploss = Nothing
    '                'Dim modifyTargetPrice As Decimal = 0
    '                'Dim supporting1 As String = Nothing
    '                'Dim supporting2 As String = Nothing
    '                'Dim supporting3 As String = Nothing
    '                'Dim supporting4 As String = Nothing
    '                'Dim supporting5 As String = Nothing
    '                If runningPayload.Date = _tradingDate.Date Then
    '                    If Not startCandleForTheTradingDay Then
    '                        If _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High > _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
    '                            (_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green OrElse
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.White) AndAlso
    '                            (_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green OrElse
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.White) Then
    '                            lastTrend = 1
    '                        ElseIf _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High <= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
    '                            (_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Red OrElse
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.White) AndAlso
    '                            (_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Red OrElse
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.White) Then
    '                            lastTrend = -1
    '                        End If
    '                        If lastTrend = -1 Then
    '                            If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green Then
    '                                Dim potentialHighEntryPrice As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.High
    '                                potentialHighEntryPrice += CalculateBuffer(potentialHighEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                                If _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
    '                                    signal = 1
    '                                    entryPrice = potentialHighEntryPrice
    '                                    slPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low
    '                                    slPrice -= CalculateBuffer(slPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                                    targetPrice = entryPrice + 100000
    '                                    lastTrend = 0
    '                                End If
    '                            End If
    '                        ElseIf lastTrend = 1 Then
    '                            If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red Then
    '                                Dim potentialLowEntryPrice As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.Low
    '                                potentialLowEntryPrice -= CalculateBuffer(potentialLowEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                                If _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
    '                                    signal = -1
    '                                    entryPrice = potentialLowEntryPrice
    '                                    slPrice = _inputPayload(runningPayload).PreviousCandlePayload.High
    '                                    slPrice += CalculateBuffer(slPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                                    targetPrice = entryPrice - 100000
    '                                    lastTrend = 0
    '                                End If
    '                            End If
    '                        End If
    '                    Else
    '                        candleCount += 1
    '                        If candleCount = 4 Then startCandleForTheTradingDay = False
    '                    End If
    '                End If

    '                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
    '                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
    '                    outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, signal)
    '                    If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
    '                    outputEntryPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryPrice)
    '                    If outputTargetPayload Is Nothing Then outputTargetPayload = New Dictionary(Of Date, Decimal)
    '                    outputTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, targetPrice)
    '                    If outputStoplossPayload Is Nothing Then outputStoplossPayload = New Dictionary(Of Date, Decimal)
    '                    outputStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, slPrice)
    '                    If outputQuantityPayload Is Nothing Then outputQuantityPayload = New Dictionary(Of Date, Integer)
    '                    outputQuantityPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, quantity)

    '                    'If outputModifyStoplossPayload Is Nothing Then outputModifyStoplossPayload = New Dictionary(Of Date, ModifyStoploss)
    '                    'outputModifyStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifySLPrice)
    '                    'If outputModifyTargetPayload Is Nothing Then outputModifyTargetPayload = New Dictionary(Of Date, Decimal)
    '                    'outputModifyTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifyTargetPrice)
    '                    'If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
    '                    'If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
    '                    'If outputSupporting3Payload Is Nothing Then outputSupporting3Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting3Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting3)
    '                    'If outputSupporting4Payload Is Nothing Then outputSupporting4Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting4Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting4)
    '                    'If outputSupporting5Payload Is Nothing Then outputSupporting5Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting5Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting5)
    '                End If
    '            Next

    '            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
    '            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
    '            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
    '            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
    '            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
    '            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
    '            'If outputModifyStoplossPayload IsNot Nothing Then outputPayload.Add("ModifyStoploss", outputModifyStoplossPayload)
    '            'If outputModifyTargetPayload IsNot Nothing Then outputPayload.Add("ModifyTarget", outputModifyTargetPayload)
    '            'If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
    '            'If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
    '            'If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
    '            'If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
    '            'If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
    '        End If
    '    End If
    'End Sub
    'Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
    '    If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
    '        Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
    '        Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
    '        'Dim outputModifyStoplossPayload As Dictionary(Of Date, ModifyStoploss) = Nothing
    '        'Dim outputModifyTargetPayload As Dictionary(Of Date, Decimal) = Nothing
    '        'Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
    '        'Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

    '        Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
    '        Indicator.VWAP.CalculateVWAP(_inputPayload, VWAPPayload)

    '        If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
    '            Dim startCandleForTheTradingDay As Boolean = True
    '            Dim lastTrend As Integer = 0
    '            Dim candleCount As Integer = 1
    '            For Each runningPayload In _inputPayload.Keys
    '                Dim signal As Integer = 0
    '                Dim entryPrice As Decimal = 0
    '                Dim slPrice As Decimal = 0
    '                Dim targetPrice As Decimal = 0
    '                Dim quantity As Integer = Strategy.CalculateQuantityFromInvestment(_quantity, 20000, _stockPrice, Trade.TypeOfStock.Cash)

    '                'Dim modifySLPrice As ModifyStoploss = Nothing
    '                'Dim modifyTargetPrice As Decimal = 0
    '                'Dim supporting1 As String = Nothing
    '                'Dim supporting2 As String = Nothing
    '                'Dim supporting3 As String = Nothing
    '                'Dim supporting4 As String = Nothing
    '                'Dim supporting5 As String = Nothing
    '                If runningPayload.Date = _tradingDate.Date Then
    '                    If Not startCandleForTheTradingDay Then
    '                        If _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High > _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low Then
    '                            lastTrend = 1
    '                        ElseIf _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High <= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
    '                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low Then
    '                            lastTrend = -1
    '                        End If
    '                        If lastTrend = -1 Then
    '                            'If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green Then
    '                            Dim potentialHighEntryPrice As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.High
    '                            potentialHighEntryPrice += CalculateBuffer(potentialHighEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                            If _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
    '                                signal = 1
    '                                entryPrice = potentialHighEntryPrice
    '                                slPrice = _inputPayload(runningPayload).PreviousCandlePayload.Low
    '                                slPrice -= CalculateBuffer(slPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                                targetPrice = entryPrice + 100000
    '                            End If
    '                            'End If
    '                            lastTrend = 0
    '                        ElseIf lastTrend = 1 Then
    '                            'If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red Then
    '                            Dim potentialLowEntryPrice As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.Low
    '                            potentialLowEntryPrice -= CalculateBuffer(potentialLowEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                            If _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
    '                                signal = -1
    '                                entryPrice = potentialLowEntryPrice
    '                                slPrice = _inputPayload(runningPayload).PreviousCandlePayload.High
    '                                slPrice += CalculateBuffer(slPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
    '                                targetPrice = entryPrice - 100000
    '                            End If
    '                            'End If
    '                            lastTrend = 0
    '                        End If
    '                    Else
    '                        candleCount += 1
    '                        If candleCount = 4 Then startCandleForTheTradingDay = False
    '                    End If
    '                End If

    '                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
    '                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
    '                    outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, signal)
    '                    If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
    '                    outputEntryPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryPrice)
    '                    If outputTargetPayload Is Nothing Then outputTargetPayload = New Dictionary(Of Date, Decimal)
    '                    outputTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, targetPrice)
    '                    If outputStoplossPayload Is Nothing Then outputStoplossPayload = New Dictionary(Of Date, Decimal)
    '                    outputStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, slPrice)
    '                    If outputQuantityPayload Is Nothing Then outputQuantityPayload = New Dictionary(Of Date, Integer)
    '                    outputQuantityPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, quantity)

    '                    'If outputModifyStoplossPayload Is Nothing Then outputModifyStoplossPayload = New Dictionary(Of Date, ModifyStoploss)
    '                    'outputModifyStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifySLPrice)
    '                    'If outputModifyTargetPayload Is Nothing Then outputModifyTargetPayload = New Dictionary(Of Date, Decimal)
    '                    'outputModifyTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifyTargetPrice)
    '                    'If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
    '                    'If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
    '                    'If outputSupporting3Payload Is Nothing Then outputSupporting3Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting3Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting3)
    '                    'If outputSupporting4Payload Is Nothing Then outputSupporting4Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting4Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting4)
    '                    'If outputSupporting5Payload Is Nothing Then outputSupporting5Payload = New Dictionary(Of Date, String)
    '                    'outputSupporting5Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting5)
    '                End If
    '            Next

    '            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
    '            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
    '            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
    '            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
    '            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
    '            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
    '            'If outputModifyStoplossPayload IsNot Nothing Then outputPayload.Add("ModifyStoploss", outputModifyStoplossPayload)
    '            'If outputModifyTargetPayload IsNot Nothing Then outputPayload.Add("ModifyTarget", outputModifyTargetPayload)
    '            'If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
    '            'If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
    '            'If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
    '            'If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
    '            'If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
    '        End If
    '    End If
    'End Sub

    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            'Dim outputModifyStoplossPayload As Dictionary(Of Date, ModifyStoploss) = Nothing
            'Dim outputModifyTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.VWAP.CalculateVWAP(_inputPayload, VWAPPayload)

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim startCandleForTheTradingDay As Boolean = True
                Dim lastBuyTrend As Integer = 0
                Dim lastBuyTrendTime As Date = Nothing
                Dim lastSellTrend As Integer = 0
                Dim lastSellTrendTime As Date = Nothing
                Dim highLevelBuyTrend As Decimal = 0
                Dim lowLevelBuyTrend As Decimal = 0
                Dim highLevelSellTrend As Decimal = 0
                Dim lowLevelSellTrend As Decimal = 0
                Dim candleCount As Integer = 1
                Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
                For Each runningPayload In _inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim slPrice As Decimal = 0
                    Dim targetPrice As Decimal = 0
                    Dim quantity As Integer = Strategy.CalculateQuantityFromInvestment(_quantity, 20000, _stockPrice, Trade.TypeOfStock.Cash) - 1

                    'Dim modifySLPrice As ModifyStoploss = Nothing
                    'Dim modifyTargetPrice As Decimal = 0
                    Dim supporting1 As String = Nothing
                    'Dim supporting2 As String = Nothing
                    'Dim supporting3 As String = Nothing
                    'Dim supporting4 As String = Nothing
                    'Dim supporting5 As String = Nothing
                    If runningPayload.Date = _tradingDate.Date Then
                        If Not startCandleForTheTradingDay Then
                            If _inputPayload(runningPayload).PreviousCandlePayload.High > _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Low > _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green Then
                                lastBuyTrend = 1
                                highLevelBuyTrend = _inputPayload(runningPayload).PreviousCandlePayload.High
                                lowLevelBuyTrend = _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low
                                lowLevelBuyTrend -= CalculateBuffer(lowLevelBuyTrend, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                lastBuyTrendTime = _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            ElseIf _inputPayload(runningPayload).PreviousCandlePayload.High < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Low < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Red Then
                                lastSellTrend = -1
                                highLevelSellTrend = _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High
                                highLevelSellTrend += CalculateBuffer(highLevelSellTrend, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                lowLevelSellTrend = _inputPayload(runningPayload).PreviousCandlePayload.Low
                                lastSellTrendTime = _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            End If
                            If lastSellTrend = -1 Then
                                If _inputPayload(runningPayload).High >= highLevelSellTrend Then
                                    signal = 1
                                    entryPrice = highLevelSellTrend
                                    'slPrice = lowLevelSellTrend
                                    slPrice = Strategy.CalculatorTargetOrStoploss(tradingSymbol, entryPrice, quantity, _maxStoplossPerTrade, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)
                                    targetPrice = Strategy.CalculatorTargetOrStoploss(tradingSymbol, entryPrice, quantity, _maxTargetPerTrade, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)
                                    supporting1 = lastSellTrendTime
                                    lastSellTrend = 0
                                ElseIf _inputPayload(runningPayload).Close < lowLevelSellTrend Then
                                    lastSellTrend = 0
                                End If
                            End If
                            If lastBuyTrend = 1 Then
                                If _inputPayload(runningPayload).Low <= lowLevelBuyTrend Then
                                    signal = -1
                                    entryPrice = lowLevelBuyTrend
                                    'slPrice = highLevelBuyTrend
                                    slPrice = Strategy.CalculatorTargetOrStoploss(tradingSymbol, entryPrice, quantity, _maxStoplossPerTrade, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)
                                    targetPrice = Strategy.CalculatorTargetOrStoploss(tradingSymbol, entryPrice, quantity, _maxTargetPerTrade, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)
                                    supporting1 = lastBuyTrendTime
                                    lastBuyTrend = 0
                                ElseIf _inputPayload(runningPayload).Close > highLevelBuyTrend Then
                                    lastBuyTrend = 0
                                End If
                            End If
                        Else
                            candleCount += 1
                            If candleCount = 3 Then startCandleForTheTradingDay = False
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

                        'If outputModifyStoplossPayload Is Nothing Then outputModifyStoplossPayload = New Dictionary(Of Date, ModifyStoploss)
                        'outputModifyStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifySLPrice)
                        'If outputModifyTargetPayload Is Nothing Then outputModifyTargetPayload = New Dictionary(Of Date, Decimal)
                        'outputModifyTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifyTargetPrice)
                        If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                        outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
                        'If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                        'outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
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
                If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
                If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
                If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
                If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
                'If outputModifyStoplossPayload IsNot Nothing Then outputPayload.Add("ModifyStoploss", outputModifyStoplossPayload)
                'If outputModifyTargetPayload IsNot Nothing Then outputPayload.Add("ModifyTarget", outputModifyTargetPayload)
                If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
                'If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
                'If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
                'If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
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
