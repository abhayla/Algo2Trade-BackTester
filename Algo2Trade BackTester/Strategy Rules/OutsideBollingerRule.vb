Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Public Class OutsideBollingerRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _candleWickPercentage As Decimal = 40 / 100
    Private ReadOnly _winRatio As Decimal = 1
    Private ReadOnly _marginMultiplier As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal quantity As Integer, ByVal canceller As CancellationTokenSource, ByVal marginMultiplier As Decimal)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _marginMultiplier = marginMultiplier
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

            Dim highEMAPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim lowEMAPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.EMA.CalculateEMA(5, Payload.PayloadFields.High, _inputPayload, highEMAPayload)
            Indicator.EMA.CalculateEMA(5, Payload.PayloadFields.Low, _inputPayload, lowEMAPayload)
            Dim highBollingerPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim lowBollingerPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim SMAPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _inputPayload, highBollingerPayload, lowBollingerPayload, SMAPayload)
            Dim tradingDate As Date = _inputPayload.Keys.LastOrDefault.Date
            Dim previousTradingDate As Date = Nothing
            Using cmn As New Common(Nothing)
                previousTradingDate = cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Cash, _inputPayload.LastOrDefault.Value.TradingSymbol, tradingDate)
            End Using
            Dim previousDayHigh As Decimal = _inputPayload.Max(Function(x)
                                                                   If x.Key.Date = previousTradingDate.Date Then
                                                                       Return x.Value.High
                                                                   Else
                                                                       Return Decimal.MinValue
                                                                   End If
                                                               End Function)
            Dim previousDayLow As Decimal = _inputPayload.Min(Function(x)
                                                                  If x.Key.Date = previousTradingDate.Date Then
                                                                      Return x.Value.Low
                                                                  Else
                                                                      Return Decimal.MaxValue
                                                                  End If
                                                              End Function)

            Dim potentialSignal As Integer = 0
            For Each runningPayload In _inputPayload.Keys
                Dim potentialEntry As Decimal = 0
                Dim potentialSL As Decimal = 0
                Dim targetPoint As Decimal = 0
                Dim pl As Decimal = 0
                Dim capital As Decimal = 0

                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = Nothing
                Dim supporting2 As String = Nothing

                If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                    If _inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                        signal = 0
                    End If
                    If _inputPayload(runningPayload).PreviousCandlePayload.Low <= lowBollingerPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.High < _inputPayload(runningPayload).High AndAlso
                    _inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Bottom >= _inputPayload(runningPayload).PreviousCandlePayload.CandleRange * _candleWickPercentage Then
                        potentialEntry = _inputPayload(runningPayload).PreviousCandlePayload.High
                        potentialEntry += CalculateBuffer(potentialEntry, RoundOfType.Floor)
                        potentialSL = _inputPayload(runningPayload).PreviousCandlePayload.Low
                        potentialSL -= CalculateBuffer(potentialSL, RoundOfType.Floor)
                        targetPoint = highEMAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) - potentialEntry
                        pl = Strategy.CalculatePL(_inputPayload(runningPayload).TradingSymbol, potentialEntry, potentialSL, quantity, quantity, Trade.TypeOfStock.Cash)
                        'capital = potentialEntry * quantity / _marginMultiplier
                        'If Math.Abs(pl / capital) * 100 <= 5 AndAlso
                        If targetPoint / _winRatio >= _inputPayload(runningPayload).PreviousCandlePayload.CandleRange Then
                            signal = 1
                            entryPrice = potentialEntry
                            slPrice = potentialSL
                            potentialSignal = signal
                            supporting1 = pl
                            If _inputPayload(runningPayload).PreviousCandlePayload.High >= previousDayLow AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Low <= previousDayLow Then
                                supporting2 = "Touch"
                            Else
                                supporting2 = String.Format("Gap {0}", entryPrice - previousDayLow)
                            End If
                        End If
                    ElseIf _inputPayload(runningPayload).PreviousCandlePayload.High >= highBollingerPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                        _inputPayload(runningPayload).PreviousCandlePayload.Low > _inputPayload(runningPayload).Low AndAlso
                        _inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Top >= _inputPayload(runningPayload).PreviousCandlePayload.CandleRange * _candleWickPercentage Then
                        potentialEntry = _inputPayload(runningPayload).PreviousCandlePayload.Low
                        potentialEntry -= CalculateBuffer(potentialEntry, RoundOfType.Floor)
                        potentialSL = _inputPayload(runningPayload).PreviousCandlePayload.High
                        potentialSL += CalculateBuffer(potentialSL, RoundOfType.Floor)
                        targetPoint = potentialEntry - lowEMAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                        pl = Strategy.CalculatePL(_inputPayload(runningPayload).TradingSymbol, potentialSL, potentialEntry, quantity, quantity, Trade.TypeOfStock.Cash)
                        'capital = potentialEntry * quantity / _marginMultiplier
                        'If Math.Abs(pl / capital) * 100 <= 5 AndAlso
                        If targetPoint / _winRatio >= _inputPayload(runningPayload).PreviousCandlePayload.CandleRange Then
                            signal = -1
                            entryPrice = potentialEntry
                            slPrice = potentialSL
                            potentialSignal = signal
                            supporting1 = pl
                            If _inputPayload(runningPayload).PreviousCandlePayload.High >= previousDayHigh AndAlso
                                _inputPayload(runningPayload).PreviousCandlePayload.Low <= previousDayHigh Then
                                supporting2 = "Touch"
                            Else
                                supporting2 = String.Format("Gap {0}", entryPrice - previousDayHigh)
                            End If
                        End If
                    End If

                    If potentialSignal = 1 Then
                        targetPrice = highEMAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate)
                    ElseIf potentialSignal = -1 Then
                        targetPrice = lowEMAPayload(_inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate)
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
