Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module TIIBasedRuleOnHA
        Public Sub CalculateRule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPayload As Dictionary(Of Date, Decimal), ByRef outputNextCandleExecutionPayload As Dictionary(Of Date, Boolean))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim TIIPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim SignalLinePayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.TrendIntensityIndex.CalculateTII(Payload.PayloadFields.Close, 14, 9, inputPayload, TIIPayload, SignalLinePayload)
                Dim previousSignal As Integer = 0
                Dim previousEntryPrice As Decimal = 0
                Dim signalCandle As Signals = Signals.None
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim nextCandleExecution As Boolean = False
                    If TIIPayload(runningPayload) < 20 AndAlso TIIPayload(runningPayload) < SignalLinePayload(runningPayload) Then
                        signalCandle = Signals.Buy
                    ElseIf TIIPayload(runningPayload) > 80 AndAlso TIIPayload(runningPayload) > SignalLinePayload(runningPayload) Then
                        signalCandle = Signals.Sell
                    End If
                    If signalCandle = Signals.Buy Then
                        If inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bearish Then
                            signal = 1
                            entryPrice = inputPayload(runningPayload).High
                        ElseIf previousSignal = 1 AndAlso inputPayload(runningPayload).High > previousEntryPrice Then
                            signal = 0
                            entryPrice = 0
                            signalCandle = Signals.None
                        ElseIf previousSignal = 1 Then
                            signal = previousSignal
                            entryPrice = previousEntryPrice
                        End If
                    ElseIf signalCandle = Signals.Sell Then
                        If inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bullish Then
                            signal = -1
                            entryPrice = inputPayload(runningPayload).Low
                        ElseIf previousSignal = -1 AndAlso inputPayload(runningPayload).Low < previousEntryPrice Then
                            signal = 0
                            entryPrice = 0
                            signalCandle = Signals.None
                        ElseIf previousSignal = -1 Then
                            signal = previousSignal
                            entryPrice = previousEntryPrice
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
        Public Enum Signals
            Buy = 1
            Sell
            None
        End Enum
    End Module
End Namespace
