Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers

Public Class VolumeReversalStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    'Public QuantityFlag As Integer = 1
    'Public MaxStoplossAmount As Decimal = 100
    'Public FirstTradeTargetMultiplier As Decimal = 2
    'Public EarlyStoploss As Boolean = False
    'Public ForwardTradeTargetMultiplier As Decimal = 3
    Public CapitalToBeUsed As Decimal = 20000
    'Public CandleBasedEntry As Boolean = False

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Private ReadOnly _stockType As Trade.TypeOfStock
    Private _firstEntryQuantity As Integer = 0

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

                Dim firstCandleOfTheTradingDay As Boolean = True
                Dim potentialHighEntryPrice As Decimal = 0
                Dim potentialLowEntryPrice As Decimal = 0
                Dim signalCandle As Payload = Nothing
                Dim trendType As TypeOfTrend = TypeOfTrend.None
                Dim firstTradeEnterd As Boolean = False
                Dim lastSignal As Integer = 0
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
                    End With
                    'Dim modifySLPrice As ModifyStoploss = Nothing
                    'Dim modifyTargetPrice As Decimal = 0
                    Dim supporting1 As String = Nothing
                    Dim supporting2 As String = Nothing
                    'Dim supporting3 As String = Nothing
                    'Dim supporting4 As String = Nothing
                    'Dim supporting5 As String = Nothing
                    If runningPayload.Date = _tradingDate.Date Then
                        If _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date AndAlso
                            _inputPayload(runningPayload).PreviousCandlePayload.Volume < _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume AndAlso
                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Volume AndAlso
                            _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Volume >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Volume Then
                            Dim highTrend As Boolean = False
                            Dim lowTrend As Boolean = False
                            If _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High Then
                                highTrend = True
                            End If
                            If _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low <= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low <= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low Then
                                lowTrend = True
                            End If
                            If highTrend AndAlso Not lowTrend AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Open <= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Close <= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High Then
                                signalCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                trendType = TypeOfTrend.High
                                potentialHighEntryPrice = 0
                                potentialLowEntryPrice = signalCandle.Low
                                potentialLowEntryPrice -= CalculateBuffer(potentialLowEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                Dim lowestLow As Decimal = Math.Min(Math.Min(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low, _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low), _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low)
                                If potentialLowEntryPrice <= lowestLow Then
                                    signalCandle = Nothing
                                    trendType = TypeOfTrend.None
                                    potentialHighEntryPrice = 0
                                    potentialLowEntryPrice = 0
                                End If
                            ElseIf lowTrend AndAlso Not highTrend AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Open >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Close >= _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low Then
                                signalCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                trendType = TypeOfTrend.Low
                                potentialHighEntryPrice = signalCandle.High
                                potentialHighEntryPrice += CalculateBuffer(potentialHighEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                potentialLowEntryPrice = 0
                                Dim highestHigh As Decimal = Math.Max(Math.Max(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High, _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High), _inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High)
                                If potentialHighEntryPrice >= highestHigh Then
                                    signalCandle = Nothing
                                    trendType = TypeOfTrend.None
                                    potentialHighEntryPrice = 0
                                    potentialLowEntryPrice = 0
                                End If
                            End If
                        End If
                        'Signal cancellation
                        If signalCandle IsNot Nothing AndAlso trendType <> TypeOfTrend.None Then
                            If trendType = TypeOfTrend.High Then
                                If _inputPayload(runningPayload).Open > signalCandle.PreviousCandlePayload.High OrElse
                                    _inputPayload(runningPayload).Close > signalCandle.PreviousCandlePayload.High Then
                                    If Not (_inputPayload(runningPayload).Low <= potentialLowEntryPrice AndAlso
                                        _inputPayload(runningPayload).CandleColor = Color.Green) Then
                                        signalCandle = Nothing
                                        trendType = TypeOfTrend.None
                                        potentialHighEntryPrice = 0
                                        potentialLowEntryPrice = 0
                                    End If
                                End If
                            ElseIf trendType = TypeOfTrend.Low Then
                                If _inputPayload(runningPayload).Open < signalCandle.PreviousCandlePayload.Low OrElse
                                    _inputPayload(runningPayload).Close < signalCandle.PreviousCandlePayload.Low Then
                                    If Not (_inputPayload(runningPayload).High >= potentialHighEntryPrice AndAlso
                                        _inputPayload(runningPayload).CandleColor = Color.Red) Then
                                        signalCandle = Nothing
                                        trendType = TypeOfTrend.None
                                        potentialHighEntryPrice = 0
                                        potentialLowEntryPrice = 0
                                    End If
                                End If
                            End If
                        End If

                        If potentialHighEntryPrice <> 0 AndAlso _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                            entryData.BuySignal = 1
                            entryData.BuyEntry = potentialHighEntryPrice
                            entryData.BuyStoploss = Math.Min(entryData.BuyEntry - ConvertFloorCeling(entryData.BuyEntry * 0.3 / 100, _tickSize, RoundOfType.Celing), signalCandle.Low - CalculateBuffer(signalCandle.Low, RoundOfType.Floor))
                            entryData.BuyTarget = entryData.BuyEntry + ConvertFloorCeling((entryData.BuyEntry - entryData.BuyStoploss) * 1.1, _tickSize, RoundOfType.Celing)

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
                            supporting2 = If(entryData.BuyStoploss = signalCandle.Low - CalculateBuffer(signalCandle.Low, RoundOfType.Floor), "Low", "0.3%")

                            signalCandle = Nothing
                            trendType = TypeOfTrend.None
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                        End If
                        If potentialLowEntryPrice <> 0 AndAlso _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                            entryData.SellSignal = -1
                            entryData.SellEntry = potentialLowEntryPrice
                            entryData.SellStoploss = Math.Max(entryData.SellEntry + ConvertFloorCeling(entryData.SellEntry * 0.3 / 100, _tickSize, RoundOfType.Celing), signalCandle.High + CalculateBuffer(signalCandle.High, RoundOfType.Floor))
                            entryData.SellTarget = entryData.SellEntry - ConvertFloorCeling((entryData.SellStoploss - entryData.SellEntry) * 1.1, _tickSize, RoundOfType.Celing)

                            If _firstEntryQuantity = 0 Then
                                Dim capitalRequired As Decimal = entryData.SellEntry * entryData.SellQuantity / Strategy.MarginMultiplier
                                If capitalRequired < CapitalToBeUsed Then
                                    entryData.SellQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                    _firstEntryQuantity = entryData.SellQuantity
                                End If
                            Else
                                entryData.SellQuantity = _firstEntryQuantity
                            End If
                            supporting1 = signalCandle.PayloadDate.ToShortTimeString
                            supporting2 = If(entryData.SellStoploss = signalCandle.High + CalculateBuffer(signalCandle.High, RoundOfType.Floor), "High", "0.3%")

                            signalCandle = Nothing
                            trendType = TypeOfTrend.None
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                        End If
                    End If

                    If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                        If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, EntryDetails)
                        outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryData)
                        'If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
                        'outputEntryPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryPrice)
                        'If outputTargetPayload Is Nothing Then outputTargetPayload = New Dictionary(Of Date, Decimal)
                        'outputTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, targetPrice)
                        'If outputStoplossPayload Is Nothing Then outputStoplossPayload = New Dictionary(Of Date, Decimal)
                        'outputStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, slPrice)
                        'If outputQuantityPayload Is Nothing Then outputQuantityPayload = New Dictionary(Of Date, Integer)
                        'outputQuantityPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, quantity)

                        'If outputModifyStoplossPayload Is Nothing Then outputModifyStoplossPayload = New Dictionary(Of Date, ModifyStoploss)
                        'outputModifyStoplossPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifySLPrice)
                        'If outputModifyTargetPayload Is Nothing Then outputModifyTargetPayload = New Dictionary(Of Date, Decimal)
                        'outputModifyTargetPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, modifyTargetPrice)
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
                'If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
                'If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
                'If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
                'If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
                'If outputModifyStoplossPayload IsNot Nothing Then outputPayload.Add("ModifyStoploss", outputModifyStoplossPayload)
                'If outputModifyTargetPayload IsNot Nothing Then outputPayload.Add("ModifyTarget", outputModifyTargetPayload)
                If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
                If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
                If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
                If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
                If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
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

    Enum TypeOfTrend
        High = 1
        Low
        None
    End Enum
End Class
