Imports Algo2TradeBLL
Imports System.Threading

Public Class JOYMA_ATMStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Public MinimumWickPercentageOfCandleRange = 50
    Public MinimumWickPercentageOutsideATRBand = 50
    Public QuantityFlag As Integer = 1
    Public MaxStoplossAmount As Decimal = 100
    Public FirstTradeTargetMultiplier As Decimal = 2
    Public TargetMultiplierWithBuffer As Boolean = False
    Public ForwardTradeTargetMultiplier As Decimal = 3
    Public CapitalToBeUsed As Decimal = 20000

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Private ReadOnly _stockType As Trade.TypeOfStock

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
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim ATRHighPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim ATRLowPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATRBands.CalculateATRBands(0.5, 14, Payload.PayloadFields.Close, _inputPayload, ATRHighPayload, ATRLowPayload)

                Dim firstCandleOfTheTradingDay As Boolean = True
                Dim potentialHighEntryPrice As Decimal = 0
                Dim potentialLowEntryPrice As Decimal = 0
                Dim signalCandle As Payload = Nothing
                Dim firstTradeEnterd As Boolean = False
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
                    Dim supporting4 As String = Nothing
                    Dim supporting5 As String = Nothing
                    If runningPayload.Date = _tradingDate.Date Then
                        'Dim highBuffer As Decimal = CalculateBuffer(potentialHighEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                        Dim highBuffer As Decimal = 2
                        If potentialHighEntryPrice <> 0 Then
                            If _inputPayload(runningPayload).High >= potentialHighEntryPrice + highBuffer Then
                                If _inputPayload(runningPayload).Low <= signalCandle.Low - highBuffer AndAlso _inputPayload(runningPayload).CandleColor = Color.Green Then
                                    potentialHighEntryPrice = 0
                                Else
                                    entryData.BuySignal = 1
                                    entryData.BuyEntry = potentialHighEntryPrice + highBuffer
                                    entryData.BuyStoploss = signalCandle.Low - highBuffer
                                    'If Not firstTradeEnterd Then
                                    '    If FirstTradeTargetMultiplier < 3 OrElse TargetMultiplierWithBuffer Then
                                    '        entryData.BuyTarget = entryData.BuyEntry + (entryData.BuyEntry - entryData.BuyStoploss) * FirstTradeTargetMultiplier
                                    '    Else
                                    '        entryData.BuyTarget = entryData.BuyEntry + (signalCandle.High - signalCandle.Low) * FirstTradeTargetMultiplier
                                    '    End If
                                    'Else
                                    If TargetMultiplierWithBuffer Then
                                        entryData.BuyTarget = entryData.BuyEntry + (entryData.BuyEntry - entryData.BuyStoploss) * ForwardTradeTargetMultiplier
                                    Else
                                        entryData.BuyTarget = entryData.BuyEntry + (signalCandle.High - signalCandle.Low) * ForwardTradeTargetMultiplier
                                    End If
                                    'End If
                                    If _stockType = Trade.TypeOfStock.Cash Then
                                        Select Case QuantityFlag
                                            Case 1
                                                entryData.BuyQuantity = Strategy.CalculateQuantityFromInvestment(_quantity, CapitalToBeUsed, entryData.BuyEntry, _stockType)
                                            Case 2
                                                entryData.BuyQuantity = Strategy.CalculateQuantityFromSL(tradingSymbol, entryData.BuyEntry, entryData.BuyStoploss, Math.Abs(MaxStoplossAmount) * -1, Trade.TypeOfStock.Cash)
                                            Case 3
                                                Dim capitalRequired As Decimal = entryData.BuyEntry * entryData.BuyQuantity / Strategy.MarginMultiplier
                                                If capitalRequired < CapitalToBeUsed Then
                                                    entryData.BuyQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                                End If
                                        End Select
                                    ElseIf _stockType = Trade.TypeOfStock.Futures Then
                                        Dim capitalRequired As Decimal = entryData.BuyEntry * entryData.BuyQuantity / Strategy.MarginMultiplier
                                        If capitalRequired < CapitalToBeUsed Then
                                            entryData.BuyQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                        End If
                                    End If
                                    supporting1 = signalCandle.PayloadDate.ToShortTimeString
                                    supporting2 = (signalCandle.CandleWicks.Bottom / signalCandle.CandleRange) * 100
                                    supporting3 = signalCandle.Low + (signalCandle.CandleWicks.Bottom * MinimumWickPercentageOutsideATRBand / 100)
                                    supporting4 = Math.Round(ATRHighPayload(signalCandle.PayloadDate), 2)
                                    supporting5 = Math.Round(ATRLowPayload(signalCandle.PayloadDate), 2)
                                    firstTradeEnterd = True
                                    potentialHighEntryPrice = 0
                                End If
                            ElseIf _inputPayload(runningPayload).Low <= signalCandle.Low - highBuffer Then
                                potentialHighEntryPrice = 0
                            End If
                        End If

                        'Dim lowBuffer As Decimal = CalculateBuffer(potentialLowEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                        Dim lowBuffer As Decimal = 2
                        If potentialLowEntryPrice <> 0 Then
                            If _inputPayload(runningPayload).Low <= potentialLowEntryPrice - lowBuffer Then
                                If _inputPayload(runningPayload).High >= signalCandle.High + lowBuffer AndAlso _inputPayload(runningPayload).CandleColor = Color.Red Then
                                    potentialLowEntryPrice = 0
                                Else
                                    entryData.SellSignal = -1
                                    entryData.SellEntry = potentialLowEntryPrice - lowBuffer
                                    entryData.SellStoploss = signalCandle.High + lowBuffer
                                    'If Not firstTradeEnterd Then
                                    '    If FirstTradeTargetMultiplier < 3 OrElse TargetMultiplierWithBuffer Then
                                    '        entryData.SellTarget = entryData.SellEntry - (entryData.SellStoploss - entryData.SellEntry) * FirstTradeTargetMultiplier
                                    '    Else
                                    '        entryData.SellTarget = entryData.SellEntry - (signalCandle.High - signalCandle.Low) * FirstTradeTargetMultiplier
                                    '    End If
                                    'Else
                                    If TargetMultiplierWithBuffer Then
                                        entryData.SellTarget = entryData.SellEntry - (entryData.SellStoploss - entryData.SellEntry) * ForwardTradeTargetMultiplier
                                    Else
                                        entryData.SellTarget = entryData.SellEntry - (signalCandle.High - signalCandle.Low) * ForwardTradeTargetMultiplier
                                    End If
                                    'End If
                                    If _stockType = Trade.TypeOfStock.Cash Then
                                        Select Case QuantityFlag
                                            Case 1
                                                entryData.SellQuantity = Strategy.CalculateQuantityFromInvestment(_quantity, CapitalToBeUsed, entryData.SellEntry, _stockType)
                                            Case 2
                                                entryData.SellQuantity = Strategy.CalculateQuantityFromSL(tradingSymbol, entryData.SellStoploss, entryData.SellEntry, Math.Abs(MaxStoplossAmount) * -1, Trade.TypeOfStock.Cash)
                                            Case 3
                                                Dim capitalRequired As Decimal = entryData.SellEntry * entryData.SellQuantity / Strategy.MarginMultiplier
                                                If capitalRequired < CapitalToBeUsed Then
                                                    entryData.SellQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                                End If
                                        End Select
                                    ElseIf _stockType = Trade.TypeOfStock.Futures Then
                                        Dim capitalRequired As Decimal = entryData.SellEntry * entryData.SellQuantity / Strategy.MarginMultiplier
                                        If capitalRequired < CapitalToBeUsed Then
                                            entryData.SellQuantity = Math.Ceiling(CapitalToBeUsed / capitalRequired) * _quantity
                                        End If
                                    End If
                                    supporting1 = signalCandle.PayloadDate.ToShortTimeString
                                    supporting2 = (signalCandle.CandleWicks.Top / signalCandle.CandleRange) * 100
                                    supporting3 = signalCandle.High - (signalCandle.CandleWicks.Top * MinimumWickPercentageOutsideATRBand / 100)
                                    supporting4 = Math.Round(ATRLowPayload(signalCandle.PayloadDate), 2)
                                    supporting5 = Math.Round(ATRLowPayload(signalCandle.PayloadDate), 2)
                                    firstTradeEnterd = True
                                    potentialLowEntryPrice = 0
                                End If
                            ElseIf _inputPayload(runningPayload).High >= signalCandle.High + lowBuffer Then
                                potentialLowEntryPrice = 0
                            End If
                        End If

                        'Signal
                        If _inputPayload(runningPayload).CandleWicks.Top >= _inputPayload(runningPayload).CandleRange * MinimumWickPercentageOfCandleRange / 100 Then
                            If _inputPayload(runningPayload).High - (_inputPayload(runningPayload).CandleWicks.Top * MinimumWickPercentageOutsideATRBand / 100) >= Math.Round(ATRHighPayload(_inputPayload(runningPayload).PayloadDate), 2) Then
                                signalCandle = _inputPayload(runningPayload)
                                potentialHighEntryPrice = 0
                                potentialLowEntryPrice = _inputPayload(runningPayload).Low
                            End If
                        ElseIf _inputPayload(runningPayload).CandleWicks.Bottom >= _inputPayload(runningPayload).CandleRange * MinimumWickPercentageOfCandleRange / 100 Then
                            If _inputPayload(runningPayload).Low + (_inputPayload(runningPayload).CandleWicks.Bottom * MinimumWickPercentageOutsideATRBand / 100) <= Math.Round(ATRLowPayload(_inputPayload(runningPayload).PayloadDate), 2) Then
                                signalCandle = _inputPayload(runningPayload)
                                potentialHighEntryPrice = _inputPayload(runningPayload).High
                                potentialLowEntryPrice = 0
                            End If
                        End If
                    End If

                    If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                        If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, EntryDetails)
                        outputSignalPayload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryData)

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
