Imports Algo2TradeBLL
Imports System.Threading

Public Class CrudeoilCandleRangeBreakoutStrategyRule
    Inherits StrategyRule
    Implements IDisposable

    Private ReadOnly _HigherTailPercentage As Decimal = 50
    Private ReadOnly _LowerTailPercentage As Decimal = 25
    Private ReadOnly _MinimumTargetPoint As Decimal = 10
    Private ReadOnly _MaximumStoplossPoint As Decimal = 10
    Private ReadOnly _MinimumStoplossPoint As Decimal = 5

    Private ReadOnly _cmn As Common = Nothing
    Private ReadOnly _tradingDate As Date
    Private ReadOnly _timeframe As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal tickSize As Decimal, ByVal quantity As Integer,
                   ByVal canceller As CancellationTokenSource,
                   ByVal cmn As Common,
                   ByVal tradingDate As Date,
                   ByVal timeframe As Integer)
        MyBase.New(inputPayload, tickSize, quantity, canceller)
        _cmn = cmn
        _tradingDate = tradingDate
        _timeframe = timeframe
    End Sub

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
            Dim outputSupporting2Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting3Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting4Payload As Dictionary(Of Date, String) = Nothing
            'Dim outputSupporting5Payload As Dictionary(Of Date, String) = Nothing

            If _inputPayload.LastOrDefault.Key.Date = _tradingDate.Date Then
                Dim firstCandleOfTheTradingDay As Boolean = True
                Dim lastSignal As Integer = 0
                Dim tradingSymbol As String = _inputPayload.LastOrDefault.Value.TradingSymbol
                For Each runningPayload In _inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim slPrice As Decimal = 0
                    Dim targetPrice As Decimal = 0
                    Dim quantity As Integer = _quantity

                    'Dim modifySLPrice As ModifyStoploss = Nothing
                    'Dim modifyTargetPrice As Decimal = 0
                    Dim supporting1 As String = Nothing
                    Dim supporting2 As String = Nothing
                    'Dim supporting3 As String = Nothing
                    'Dim supporting4 As String = Nothing
                    'Dim supporting5 As String = Nothing
                    If runningPayload.Date = _tradingDate.Date Then
                        If Not firstCandleOfTheTradingDay Then
                            If (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Bottom / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100 >= _HigherTailPercentage AndAlso
                                (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Top / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100 <= _LowerTailPercentage Then
                                Dim potentialEntryPrice As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.High
                                potentialEntryPrice += CalculateBuffer(potentialEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                Dim potentialStoploss As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.Low
                                potentialStoploss -= CalculateBuffer(potentialStoploss, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                If (potentialEntryPrice - potentialStoploss) >= _MinimumStoplossPoint AndAlso
                                    (potentialEntryPrice - potentialStoploss) <= _MaximumStoplossPoint Then
                                    If _inputPayload(runningPayload).High >= potentialEntryPrice Then
                                        If _inputPayload(runningPayload).Open >= potentialEntryPrice Then
                                            potentialEntryPrice = _inputPayload(runningPayload).Open
                                        End If
                                        signal = 1
                                        entryPrice = potentialEntryPrice
                                        slPrice = potentialStoploss
                                        If Math.Abs(entryPrice - slPrice) * 2 < _MinimumTargetPoint Then
                                            targetPrice = entryPrice + _MinimumTargetPoint
                                        Else
                                            targetPrice = entryPrice + Math.Abs(entryPrice - slPrice) * 2
                                        End If
                                        supporting1 = (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Top / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100
                                        supporting2 = (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Bottom / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100
                                    End If
                                End If
                            ElseIf (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Top / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100 >= _HigherTailPercentage AndAlso
                                (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Bottom / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100 <= _LowerTailPercentage Then
                                Dim potentialEntryPrice As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.Low
                                potentialEntryPrice -= CalculateBuffer(potentialEntryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                Dim potentialStoploss As Decimal = _inputPayload(runningPayload).PreviousCandlePayload.High
                                potentialStoploss += CalculateBuffer(potentialStoploss, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                                If (potentialStoploss - potentialEntryPrice) >= _MinimumStoplossPoint AndAlso
                                    (potentialStoploss - potentialEntryPrice) <= _MaximumStoplossPoint Then
                                    If _inputPayload(runningPayload).Low <= potentialEntryPrice Then
                                        If _inputPayload(runningPayload).Open <= potentialEntryPrice Then
                                            potentialEntryPrice = _inputPayload(runningPayload).Open
                                        End If
                                        signal = -1
                                        entryPrice = potentialEntryPrice
                                        slPrice = potentialStoploss
                                        If Math.Abs(entryPrice - slPrice) * 2 < _MinimumTargetPoint Then
                                            targetPrice = entryPrice - _MinimumTargetPoint
                                        Else
                                            targetPrice = entryPrice - Math.Abs(entryPrice - slPrice) * 2
                                        End If
                                        supporting1 = (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Top / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100
                                        supporting2 = (_inputPayload(runningPayload).PreviousCandlePayload.CandleWicks.Bottom / _inputPayload(runningPayload).PreviousCandlePayload.CandleRange) * 100
                                    End If
                                End If
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
                If outputEntryPayload IsNot Nothing Then outputPayload.Add("Entry", outputEntryPayload)
                If outputTargetPayload IsNot Nothing Then outputPayload.Add("Target", outputTargetPayload)
                If outputStoplossPayload IsNot Nothing Then outputPayload.Add("Stoploss", outputStoplossPayload)
                If outputQuantityPayload IsNot Nothing Then outputPayload.Add("Quantity", outputQuantityPayload)
                'If outputModifyStoplossPayload IsNot Nothing Then outputPayload.Add("ModifyStoploss", outputModifyStoplossPayload)
                'If outputModifyTargetPayload IsNot Nothing Then outputPayload.Add("ModifyTarget", outputModifyTargetPayload)
                If outputSupporting1Payload IsNot Nothing Then outputPayload.Add("Supporting1", outputSupporting1Payload)
                If outputSupporting2Payload IsNot Nothing Then outputPayload.Add("Supporting2", outputSupporting2Payload)
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
