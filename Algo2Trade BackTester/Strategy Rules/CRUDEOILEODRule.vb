Imports System.Threading
Imports Algo2TradeBLL
Public Class CRUDEOILEODRule
    Inherits StrategyRule
    Implements IDisposable

    Private _cmn As Common = Nothing
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tickSize As Decimal, ByVal quantity As Integer, ByVal canceller As CancellationTokenSource, ByVal cmn As Common)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
    End Sub
    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            'Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing

            Dim tradingDate As Date = _inputPayload.LastOrDefault.Key.Date
            Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
            Dim previousTradingDay As Date = _cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Commodity, tradingSymbol.Remove(tradingSymbol.Count - 8), tradingDate)
            Dim eodPayload As Dictionary(Of Date, Payload) = _cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Commodity, tradingSymbol, previousTradingDay.AddDays(-10), previousTradingDay)
            Dim rulePayload As IEnumerable(Of KeyValuePair(Of Date, Payload)) = eodPayload.Skip(Math.Max(0, eodPayload.Count() - 5))
            Dim minAvg As Decimal = Nothing
            Dim maxAvg As Decimal = Nothing
            If rulePayload IsNot Nothing AndAlso rulePayload.Count > 0 Then
                If rulePayload.Count <> 5 Then Throw New ApplicationException(String.Format("Invalid Dataset. Count:{0}", rulePayload.Count))
                Dim highOpen As List(Of Decimal) = Nothing
                For Each runningRulePayload In rulePayload
                    If highOpen Is Nothing Then highOpen = New List(Of Decimal)
                    highOpen.Add(runningRulePayload.Value.High - runningRulePayload.Value.Open)
                Next

                Dim openLow As List(Of Decimal) = Nothing
                For Each runningRulePayload In rulePayload
                    If openLow Is Nothing Then openLow = New List(Of Decimal)
                    openLow.Add(runningRulePayload.Value.Open - runningRulePayload.Value.Low)
                Next

                Dim min As List(Of Decimal) = Nothing
                For i = 0 To openLow.Count - 1
                    If min Is Nothing Then min = New List(Of Decimal)
                    Dim result As Decimal = If(Math.Min(highOpen(i), openLow(i)) < 10, 10, Math.Min(highOpen(i), openLow(i)))
                    min.Add(result)
                Next

                Dim max As List(Of Decimal) = Nothing
                For i = 0 To openLow.Count - 1
                    If max Is Nothing Then max = New List(Of Decimal)
                    max.Add(Math.Max(highOpen(i), openLow(i)))
                Next

                minAvg = min.Average()
                maxAvg = max.Average()
            End If
            Dim potentialLongEntry As Decimal = 0
            Dim potentialShortEntry As Decimal = 0
            Dim oncePerDay As Boolean = True
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                'Dim supporting1 As String = Nothing
                'Dim supporting2 As String = Nothing
                If runningPayload.Date = tradingDate.Date Then
                    If oncePerDay Then
                        potentialLongEntry = _inputPayload(runningPayload).Open + minAvg
                        potentialShortEntry = _inputPayload(runningPayload).Open - minAvg
                        oncePerDay = False
                    End If
                    If Not potentialLongEntry = 0 AndAlso Not potentialShortEntry = 0 Then
                        If _inputPayload(runningPayload).High > potentialLongEntry AndAlso
                        _inputPayload(runningPayload).Low < potentialShortEntry Then
                            If _inputPayload(runningPayload).CandleColor = Color.Green Then
                                signal = -1
                                entryPrice = potentialShortEntry
                                slPrice = entryPrice + 15
                                targetPrice = entryPrice - maxAvg
                                potentialLongEntry = 0
                                potentialShortEntry = 0
                            Else
                                signal = 1
                                entryPrice = potentialLongEntry
                                slPrice = entryPrice - 15
                                targetPrice = entryPrice + maxAvg
                                potentialLongEntry = 0
                                potentialShortEntry = 0
                            End If
                        ElseIf _inputPayload(runningPayload).High > potentialLongEntry Then
                            signal = 1
                            entryPrice = potentialLongEntry
                            slPrice = entryPrice - 15
                            targetPrice = entryPrice + maxAvg
                            potentialLongEntry = 0
                            potentialShortEntry = 0
                        ElseIf _inputPayload(runningPayload).Low < potentialShortEntry Then
                            signal = -1
                            entryPrice = potentialShortEntry
                            slPrice = entryPrice + 15
                            targetPrice = entryPrice - maxAvg
                            potentialLongEntry = 0
                            potentialShortEntry = 0
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
                    'If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                    'outputSupporting1Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting1)
                    'If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                    'outputSupporting2Payload.Add(_inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, supporting2)
                End If
            Next
            If outputPayload Is Nothing Then outputPayload = New Dictionary(Of String, Object)
            If outputSignalPayload IsNot Nothing Then outputPayload.Add("Signal", outputSignalPayload)
            If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
            If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
            If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
            If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
            'If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
            'If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
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
