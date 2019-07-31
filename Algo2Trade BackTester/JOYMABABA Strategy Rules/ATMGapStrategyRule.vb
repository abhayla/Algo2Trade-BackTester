Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers

Public Class ATMGapStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Public TargetMultiplier As Decimal = 3
    Public CapitalToBeUsed As Decimal = 20000
    Public ATRToBeUsed As ATRCandle = ATRCandle.SignalCandle
    Public ATRMultiplier As Decimal = 1

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Private ReadOnly _stockType As Trade.TypeOfStock
    Private ReadOnly _signalCandleTime As Date
    Private ReadOnly _gapPercentage As Decimal
    Private ReadOnly _lastDayClose As Decimal
    Private _firstEntryQuantity As Integer = 0
    Private _usableATR As Decimal = Decimal.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal, ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common,
                   ByVal tradingDate As Date,
                   ByVal timeframe As Integer,
                   ByVal stockType As Trade.TypeOfStock,
                   ByVal gapPercentage As Decimal,
                   ByVal lastDayClose As Decimal)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
        _tradingDate = tradingDate
        _timeframe = timeframe
        _stockType = stockType
        _gapPercentage = gapPercentage
        _lastDayClose = lastDayClose
        _signalCandleTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 16, 0)
    End Sub

    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _timeframe > 1 Then Throw New ApplicationException(String.Format("Can not run this rule on {0} minute timeframe. Only 1 minute timeframe is allowed.", _timeframe))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, EntryDetails) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
                Dim rawInstrumentName As String = tradingSymbol.Remove(tradingSymbol.Count - 8)
                Dim cashInstrumentPayload As Dictionary(Of Date, Payload) = _cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Cash, rawInstrumentName, _tradingDate, _tradingDate)
                Dim gapExists As Boolean = False
                If cashInstrumentPayload IsNot Nothing AndAlso cashInstrumentPayload.Count > 0 AndAlso
                    cashInstrumentPayload.ContainsKey(_signalCandleTime) AndAlso cashInstrumentPayload(_signalCandleTime).PreviousCandlePayload IsNot Nothing Then
                    If Math.Abs(_gapPercentage) >= 0.5 Then
                        If _gapPercentage >= 0 Then
                            If cashInstrumentPayload(_signalCandleTime).PreviousCandlePayload.Low - _lastDayClose > 0 AndAlso
                                cashInstrumentPayload(_signalCandleTime).Open - _lastDayClose > 0 Then
                                gapExists = True
                            End If
                        Else
                            If cashInstrumentPayload(_signalCandleTime).PreviousCandlePayload.High - _lastDayClose < 0 AndAlso
                                cashInstrumentPayload(_signalCandleTime).Open - _lastDayClose < 0 Then
                                gapExists = True
                            End If
                        End If
                    End If
                End If
                If gapExists Then
                    Dim ATRPayload As Dictionary(Of Date, Decimal) = Nothing
                    Indicator.ATR.CalculateATR(14, _inputPayload, ATRPayload)

                    Dim potentialHighEntryPrice As Decimal = 0
                    Dim potentialLowEntryPrice As Decimal = 0
                    Dim levelPrice As Decimal = 0
                    Dim signalCandle As Payload = Nothing
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
                        Dim supporting4 As String = Nothing
                        Dim supporting5 As String = Nothing

                        If runningPayload.Date = _tradingDate.Date Then
                            If ATRToBeUsed = ATRCandle.PreviousDayLastCandle AndAlso _usableATR = Decimal.MinValue Then
                                _usableATR = ATRPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                            End If
                            If potentialHighEntryPrice = 0 AndAlso potentialLowEntryPrice = 0 AndAlso levelPrice = 0 Then
                                If runningPayload = _signalCandleTime Then
                                    signalCandle = _inputPayload(runningPayload).PreviousCandlePayload
                                    If ATRToBeUsed = ATRCandle.SignalCandle Then
                                        _usableATR = ATRPayload(signalCandle.PreviousCandlePayload.PayloadDate)
                                    End If
                                    If _gapPercentage >= 0 Then
                                        levelPrice = _inputPayload(runningPayload).Open + ConvertFloorCeling(_usableATR * ATRMultiplier, _tickSize, RoundOfType.Celing)
                                    Else
                                        levelPrice = _inputPayload(runningPayload).Open - ConvertFloorCeling(_usableATR * ATRMultiplier, _tickSize, RoundOfType.Celing)
                                    End If
                                    potentialHighEntryPrice = levelPrice + ConvertFloorCeling(_usableATR * ATRMultiplier, _tickSize, RoundOfType.Celing)
                                    potentialLowEntryPrice = levelPrice - ConvertFloorCeling(_usableATR * ATRMultiplier, _tickSize, RoundOfType.Celing)
                                End If
                            End If
                            If potentialHighEntryPrice <> 0 AndAlso potentialLowEntryPrice <> 0 AndAlso levelPrice <> 0 Then
                                If _inputPayload(runningPayload).High >= potentialHighEntryPrice AndAlso
                                _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                                    entryData.BuySignal = 1
                                    entryData.BuyEntry = potentialHighEntryPrice
                                    entryData.BuyStoploss = levelPrice
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
                                    supporting2 = levelPrice
                                    supporting3 = _usableATR
                                    supporting4 = _gapPercentage
                                    supporting5 = _lastDayClose

                                    entryData.SellSignal = -1
                                    entryData.SellEntry = potentialLowEntryPrice
                                    entryData.SellStoploss = levelPrice
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
                                    supporting2 = levelPrice
                                    supporting3 = _usableATR
                                    supporting4 = _gapPercentage
                                    supporting5 = _lastDayClose
                                Else
                                    If _inputPayload(runningPayload).High >= potentialHighEntryPrice Then
                                        entryData.BuySignal = 1
                                        entryData.BuyEntry = potentialHighEntryPrice
                                        entryData.BuyStoploss = levelPrice
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
                                        supporting2 = levelPrice
                                        supporting3 = _usableATR
                                        supporting4 = _gapPercentage
                                        supporting5 = _lastDayClose
                                    End If
                                    If _inputPayload(runningPayload).Low <= potentialLowEntryPrice Then
                                        entryData.SellSignal = -1
                                        entryData.SellEntry = potentialLowEntryPrice
                                        entryData.SellStoploss = levelPrice
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
                                        supporting2 = levelPrice
                                        supporting3 = _usableATR
                                        supporting4 = _gapPercentage
                                        supporting5 = _lastDayClose
                                    End If
                                End If
                            End If
                        End If

                        If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                            If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, EntryDetails)
                            outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryData)

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
                    If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
                    If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
                    If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
                    If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
                    If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
                End If
            End If
        End If
    End Sub

    Public Enum ATRCandle
        PreviousDayLastCandle
        SignalCandle
    End Enum

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