Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module HeikenAshiReversalRule
        Public Sub CalculateHeikenAshiReversalRule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPayload As Dictionary(Of Date, Decimal), ByRef outputNextCandleExecutionPayload As Dictionary(Of Date, Boolean))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim nextCandleExecution As Boolean = False

                    If inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bearish Then
                        signal = 1
                        entryPrice = inputPayload(runningPayload).High
                    ElseIf inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bullish Then
                        signal = -1
                        entryPrice = inputPayload(runningPayload).Low
                    End If
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
