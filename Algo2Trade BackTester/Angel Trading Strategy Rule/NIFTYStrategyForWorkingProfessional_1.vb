Imports Algo2TradeBLL
Imports System.Threading

Public Class NIFTYStrategyForWorkingProfessional_1
    Inherits StrategyRule
    Implements IDisposable

    Private _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _totalPL As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal,
                   ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common,
                   ByVal tradingDate As Date,
                   ByVal totalPL As Decimal)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
        _tradingDate = tradingDate
        _totalPL = totalPL
    End Sub

    Public Overrides Sub CalculateRule(ByRef outputPayload As Dictionary(Of String, Object))
        If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputQuantityPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputSupporting1Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim previousTradingDay As Date = _cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Futures, "NIFTY", _tradingDate)
                Dim eodPayload As Dictionary(Of Date, Payload) = _cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Futures, _inputPayload.LastOrDefault.Value.TradingSymbol, previousTradingDay, previousTradingDay)

                If eodPayload IsNot Nothing AndAlso eodPayload.Count > 0 Then
                    Dim previousDayNiftyPayload As Payload = eodPayload.FirstOrDefault.Value
                    Dim firstCandleOfTheTradingDay As Boolean = True
                    Dim takeTrade As Boolean = True
                    For Each runningPayload In _inputPayload.Keys
                        Dim signal As Integer = 0
                        Dim entryPrice As Decimal = 0
                        Dim slPrice As Decimal = 0
                        Dim targetPrice As Decimal = 0
                        Dim quantity As Integer = _quantity
                        Dim supporting1 As String = previousDayNiftyPayload.CandleColor.ToString
                        'Dim supporting2 As String = Nothing
                        'Dim supporting3 As String = Nothing
                        'Dim supporting4 As String = Nothing
                        'Dim supporting5 As String = Nothing
                        Dim quantityToBeAdded As Integer = 0

                        If runningPayload.Date = _tradingDate.Date Then
                            If Not firstCandleOfTheTradingDay Then
                                If takeTrade AndAlso previousDayNiftyPayload.CandleColor = Color.Red Then
                                    signal = 1
                                    entryPrice = _inputPayload(runningPayload).Open
                                    targetPrice = entryPrice + 10000
                                    slPrice = entryPrice - 50
                                    If _totalPL > 0 Then
                                        quantityToBeAdded = Strategy.CalculateQuantityFromInvestment(_quantity, _totalPL, entryPrice, Trade.TypeOfStock.Futures)
                                        If quantityToBeAdded > 0 Then
                                            quantity += quantityToBeAdded
                                        End If
                                    End If
                                    takeTrade = False
                                ElseIf takeTrade AndAlso previousDayNiftyPayload.CandleColor = Color.Green Then
                                    signal = -1
                                    entryPrice = _inputPayload(runningPayload).Open
                                    targetPrice = entryPrice - 10000
                                    slPrice = entryPrice + 50
                                    If _totalPL > 0 Then
                                        quantityToBeAdded = Strategy.CalculateQuantityFromInvestment(_quantity, _totalPL, entryPrice, Trade.TypeOfStock.Futures)
                                        If quantityToBeAdded > 0 Then
                                            quantity += quantityToBeAdded
                                        End If
                                    End If
                                    takeTrade = False
                                End If
                            Else
                                firstCandleOfTheTradingDay = False
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
                    If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
                    'If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
                    'If outputSupporting3Payload IsNot Nothing Then outputPayload.Add("Supporting3", outputSupporting3Payload)
                    'If outputSupporting4Payload IsNot Nothing Then outputPayload.Add("Supporting4", outputSupporting4Payload)
                    'If outputSupporting5Payload IsNot Nothing Then outputPayload.Add("Supporting5", outputSupporting5Payload)
                End If
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
