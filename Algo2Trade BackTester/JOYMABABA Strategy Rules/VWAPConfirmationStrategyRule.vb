Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers

Public Class VWAPConfirmationStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    'Public QuantityFlag As Integer = 1
    'Public MaxStoplossAmount As Decimal = 100
    'Public FirstTradeTargetMultiplier As Decimal = 3
    'Public EarlyStoploss As Boolean = False
    Public ForwardTradeTargetMultiplier As Decimal = 3
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
                Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.VWAP.CalculateVWAP(_inputPayload, VWAPPayload)
                Dim ATRPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATR.CalculateATR(14, _inputPayload, ATRPayload)

                Dim firstCandleOfTheTradingDay As Boolean = True
                Dim potentialHighEntryPrice As Decimal = 0
                Dim potentialLowEntryPrice As Decimal = 0
                Dim signalCandle As Payload = Nothing
                Dim confirmationCandle As Payload = Nothing
                Dim vwapSide As SideOfVWAP = SideOfVWAP.None
                Dim eligibleForSignalCheck As Boolean = False
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
                    Dim supporting3 As String = Nothing
                    'Dim supporting4 As String = Nothing
                    'Dim supporting5 As String = Nothing
                    If runningPayload.Date = _tradingDate.Date Then
                        entryData.BuyStoploss = ConvertFloorCeling(VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), _tickSize, RoundOfType.Celing)
                        entryData.BuyStoploss -= CalculateBuffer(entryData.BuyStoploss, RoundOfType.Floor)
                        entryData.SellStoploss = ConvertFloorCeling(VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), _tickSize, RoundOfType.Celing)
                        entryData.SellStoploss += CalculateBuffer(entryData.BuyStoploss, RoundOfType.Floor)

                        If Not eligibleForSignalCheck Then
                            If _inputPayload(runningPayload).PreviousCandlePayload.High >= VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Low <= VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                eligibleForSignalCheck = True
                            End If
                        End If
                        'Signal cancellation
                        If signalCandle IsNot Nothing AndAlso vwapSide <> SideOfVWAP.None Then
                            If vwapSide = SideOfVWAP.Above Then
                                If _inputPayload(runningPayload).PreviousCandlePayload.Close < VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                    signalCandle = Nothing
                                    confirmationCandle = Nothing
                                    vwapSide = SideOfVWAP.None
                                    potentialHighEntryPrice = 0
                                    potentialLowEntryPrice = 0
                                End If
                            ElseIf vwapSide = SideOfVWAP.Below Then
                                If _inputPayload(runningPayload).PreviousCandlePayload.Close > VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                    signalCandle = Nothing
                                    confirmationCandle = Nothing
                                    vwapSide = SideOfVWAP.None
                                    potentialHighEntryPrice = 0
                                    potentialLowEntryPrice = 0
                                End If
                            End If
                        End If
                        'Check Confirmation
                        If signalCandle IsNot Nothing AndAlso vwapSide <> SideOfVWAP.None AndAlso
                            Not _inputPayload(runningPayload).PreviousCandlePayload.BlankCandle Then
                            If vwapSide = SideOfVWAP.Above Then
                                If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Red OrElse
                                    _inputPayload(runningPayload).PreviousCandlePayload.Low < signalCandle.Low Then
                                    confirmationCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                    potentialHighEntryPrice = confirmationCandle.High
                                    potentialHighEntryPrice += CalculateBuffer(potentialHighEntryPrice, RoundOfType.Floor)
                                End If
                            ElseIf vwapSide = SideOfVWAP.Below Then
                                If _inputPayload(runningPayload).PreviousCandlePayload.CandleColor = Color.Green OrElse
                                    _inputPayload(runningPayload).PreviousCandlePayload.High > signalCandle.High Then
                                    confirmationCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                    potentialLowEntryPrice = confirmationCandle.Low
                                    potentialLowEntryPrice -= CalculateBuffer(potentialLowEntryPrice, RoundOfType.Floor)
                                End If
                            End If
                        End If
                        'Check Signal
                        If eligibleForSignalCheck AndAlso signalCandle Is Nothing AndAlso
                            Not _inputPayload(runningPayload).PreviousCandlePayload.BlankCandle Then
                            If _inputPayload(runningPayload).PreviousCandlePayload.Low > VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                signalCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                vwapSide = SideOfVWAP.Above
                            ElseIf _inputPayload(runningPayload).PreviousCandlePayload.High < VWAPPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                signalCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                vwapSide = SideOfVWAP.Below
                            End If
                            If signalCandle IsNot Nothing Then
                                If vwapSide = SideOfVWAP.Above Then
                                    If signalCandle.CandleColor = Color.Red Then
                                        confirmationCandle = signalCandle
                                        potentialHighEntryPrice = confirmationCandle.High
                                        potentialHighEntryPrice += CalculateBuffer(potentialHighEntryPrice, RoundOfType.Floor)
                                    End If
                                ElseIf vwapSide = SideOfVWAP.Below Then
                                    If signalCandle.CandleColor = Color.Green Then
                                        confirmationCandle = signalCandle
                                        potentialLowEntryPrice = confirmationCandle.Low
                                        potentialLowEntryPrice -= CalculateBuffer(potentialLowEntryPrice, RoundOfType.Floor)
                                    End If
                                End If
                            End If
                        End If

                        'Check Breakout
                        If potentialHighEntryPrice <> 0 AndAlso _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                            entryData.BuySignal = 1
                            entryData.BuyEntry = potentialHighEntryPrice
                            'entryData.BuyStoploss = Math.Min(entryData.BuyEntry - ConvertFloorCeling(entryData.BuyEntry * 0.3 / 100, _tickSize, RoundOfType.Celing), signalCandle.Low - CalculateBuffer(signalCandle.Low, RoundOfType.Floor))
                            entryData.BuyTarget = entryData.BuyEntry + ConvertFloorCeling(ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * ForwardTradeTargetMultiplier, _tickSize, RoundOfType.Celing)

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
                            supporting2 = confirmationCandle.PayloadDate.ToShortTimeString
                            supporting3 = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)

                            signalCandle = Nothing
                            confirmationCandle = Nothing
                            vwapSide = SideOfVWAP.None
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                            eligibleForSignalCheck = False
                        End If
                        If potentialLowEntryPrice <> 0 AndAlso _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                            entryData.SellSignal = -1
                            entryData.SellEntry = potentialLowEntryPrice
                            'entryData.SellStoploss = Math.Max(entryData.SellEntry + ConvertFloorCeling(entryData.SellEntry * 0.3 / 100, _tickSize, RoundOfType.Celing), signalCandle.High + CalculateBuffer(signalCandle.High, RoundOfType.Floor))
                            entryData.SellTarget = entryData.SellEntry - ConvertFloorCeling(ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) * ForwardTradeTargetMultiplier, _tickSize, RoundOfType.Celing)

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
                            supporting2 = confirmationCandle.PayloadDate.ToShortTimeString
                            supporting3 = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)

                            signalCandle = Nothing
                            confirmationCandle = Nothing
                            vwapSide = SideOfVWAP.None
                            potentialHighEntryPrice = 0
                            potentialLowEntryPrice = 0
                            eligibleForSignalCheck = False
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
                        If outputSupporting3Payload Is Nothing Then outputSupporting3Payload = New Dictionary(Of Date, String)
                        outputSupporting3Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting3)
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

    Enum SideOfVWAP
        Above = 1
        Below
        None
    End Enum
End Class
