Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module VWAPTouchOneSidedOpenClose
        Public Sub CalculateRule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPayload As Dictionary(Of Date, Decimal), ByRef outputNextCandleExecutionPayload As Dictionary(Of Date, Boolean))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim VWAPPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.VWAP.CalculateVWAP(inputPayload, VWAPPayload)
                Dim previousSignal As Integer = 0
                Dim previousEntryPrice As Decimal = 0
                Dim signalCandle As Payload = Nothing
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim nextCandleExecution As Boolean = False
                    If inputPayload(runningPayload).Low < VWAPPayload(runningPayload) AndAlso
                        inputPayload(runningPayload).Open > VWAPPayload(runningPayload) AndAlso
                        inputPayload(runningPayload).Close > VWAPPayload(runningPayload) Then
                        signal = 1
                        entryPrice = inputPayload(runningPayload).High
                        signalCandle = inputPayload(runningPayload)
                    ElseIf inputPayload(runningPayload).High > VWAPPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Open < VWAPPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Close < VWAPPayload(runningPayload) Then
                        signal = -1
                        entryPrice = inputPayload(runningPayload).Low
                        signalCandle = inputPayload(runningPayload)
                    Else
                        If previousSignal <> 0 Then
                            If inputPayload(runningPayload).High > signalCandle.High OrElse
                                inputPayload(runningPayload).Low < signalCandle.Low Then
                                signal = 0
                                entryPrice = 0
                                signalCandle = Nothing
                            Else
                                signal = previousSignal
                                entryPrice = previousEntryPrice
                            End If
                        End If
                    End If
                    previousSignal = signal
                    previousEntryPrice = entryPrice
                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                    outputSignalPayload.Add(runningPayload, signal)
                    If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
                    outputEntryPayload.Add(runningPayload, entryPrice)
                    If outputNextCandleExecutionPayload Is Nothing Then outputNextCandleExecutionPayload = New Dictionary(Of Date, Boolean)
                    outputNextCandleExecutionPayload.Add(runningPayload, nextCandleExecution)
                Next
            End If
        End Sub
    End Module
End Namespace
