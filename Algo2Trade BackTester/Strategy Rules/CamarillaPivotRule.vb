Imports System.Threading
Imports Algo2TradeBLL
Public Class CamarillaPivotRule
    Inherits StrategyRule
    Implements IDisposable

    Private _cmn As Common = Nothing
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal,
                   ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common)
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
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing

            Dim tradingDate As Date = _inputPayload.LastOrDefault.Key.Date
            Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
            Dim previousTradingDay As Date = _cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Cash, tradingSymbol, tradingDate)
            Dim eodPayload As Dictionary(Of Date, Payload) = _cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Cash, tradingSymbol, previousTradingDay, previousTradingDay)
            Dim camarillaPivot As Common.CamarillaPivotPoints = Nothing
            If eodPayload IsNot Nothing AndAlso eodPayload.Count > 0 Then
                camarillaPivot = _cmn.CalculateCamarillaPivotPoints(eodPayload.LastOrDefault.Value.High, eodPayload.LastOrDefault.Value.Low, eodPayload.LastOrDefault.Value.Close)
            End If

            Dim potentialHighEntry As Decimal = 0
            Dim potentialLowEntry As Decimal = 0
            Dim oncePerDay As Boolean = True
            For Each runningPayload In _inputPayload.Keys
                Dim signal As Integer = 0
                Dim entryPrice As Decimal = 0
                Dim slPrice As Decimal = 0
                Dim targetPrice As Decimal = 0
                Dim quantity As Integer = _quantity
                Dim supporting1 As String = potentialHighEntry
                Dim supporting2 As String = potentialLowEntry

                If runningPayload.Date = tradingDate.Date Then
                    If oncePerDay Then
                        If camarillaPivot IsNot Nothing Then
                            potentialHighEntry = Math.Round(camarillaPivot.H2, 2)
                            potentialLowEntry = Math.Round(camarillaPivot.L2, 2)
                            supporting1 = potentialHighEntry
                            supporting2 = potentialLowEntry
                        End If
                        If Not potentialHighEntry = 0 AndAlso Not potentialLowEntry = 0 Then
                            If _inputPayload(runningPayload).Open > potentialHighEntry Then
                                signal = -1
                                entryPrice = potentialHighEntry
                                slPrice = entryPrice + entryPrice * 0.2 / 100
                                targetPrice = entryPrice - entryPrice * 0.3 / 100
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            ElseIf _inputPayload(runningPayload).Open < potentialLowEntry Then
                                signal = 1
                                entryPrice = potentialLowEntry
                                slPrice = entryPrice - entryPrice * 0.2 / 100
                                targetPrice = entryPrice + entryPrice * 0.3 / 100
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            End If
                        End If
                        oncePerDay = False
                    Else
                        If Not potentialHighEntry = 0 AndAlso Not potentialLowEntry = 0 Then
                            If _inputPayload(runningPayload).PreviousCandlePayload.Close > potentialHighEntry AndAlso
                            _inputPayload(runningPayload).Low < potentialHighEntry Then
                                signal = -1
                                entryPrice = potentialHighEntry
                                slPrice = entryPrice + entryPrice * 0.2 / 100
                                targetPrice = entryPrice - entryPrice * 0.3 / 100
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            ElseIf _inputPayload(runningPayload).PreviousCandlePayload.Close < potentialLowEntry AndAlso
                            _inputPayload(runningPayload).High > potentialLowEntry Then
                                signal = 1
                                entryPrice = potentialLowEntry
                                slPrice = entryPrice - entryPrice * 0.2 / 100
                                targetPrice = entryPrice + entryPrice * 0.3 / 100
                                potentialHighEntry = 0
                                potentialLowEntry = 0
                            End If
                        End If
                    End If
                End If

                'If _inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                outputSignalPayload.Add(runningPayload.AddMinutes(-5), signal)
                If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
                outputEntryPayload.Add(runningPayload.AddMinutes(-5), entryPrice)
                If outputTargetPayload Is Nothing Then outputTargetPayload = New Dictionary(Of Date, Decimal)
                outputTargetPayload.Add(runningPayload.AddMinutes(-5), targetPrice)
                If outputStoplossPayload Is Nothing Then outputStoplossPayload = New Dictionary(Of Date, Decimal)
                outputStoplossPayload.Add(runningPayload.AddMinutes(-5), slPrice)
                If outputQuantityPayload Is Nothing Then outputQuantityPayload = New Dictionary(Of Date, Integer)
                outputQuantityPayload.Add(runningPayload.AddMinutes(-5), quantity)
                If outputSupporting1Payload Is Nothing Then outputSupporting1Payload = New Dictionary(Of Date, String)
                outputSupporting1Payload.Add(runningPayload.AddMinutes(-5), supporting1)
                If outputSupporting2Payload Is Nothing Then outputSupporting2Payload = New Dictionary(Of Date, String)
                outputSupporting2Payload.Add(runningPayload.AddMinutes(-5), supporting2)
                'End If
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
