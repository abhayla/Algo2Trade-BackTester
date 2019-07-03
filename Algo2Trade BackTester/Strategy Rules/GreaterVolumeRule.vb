Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module GreaterVolumeRule
        Public Sub CalculateGreaterVolumeRule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPayload As Dictionary(Of Date, Decimal), ByRef outputNextCandleExecutionPayload As Dictionary(Of Date, Boolean))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim oncePerDay As Boolean = False
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim nextCandleExecution As Boolean = False
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                        runningPayload.Date <> inputPayload(runningPayload).PayloadDate.Date Then
                        oncePerDay = False
                    End If

                    If Not oncePerDay AndAlso inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                        inputPayload(runningPayload).CandleColor = inputPayload(runningPayload).PreviousCandlePayload.CandleColor AndAlso
                        inputPayload(runningPayload).VolumeColor = inputPayload(runningPayload).PreviousCandlePayload.VolumeColor AndAlso
                        inputPayload(runningPayload).Volume > inputPayload(runningPayload).PreviousCandlePayload.Volume Then
                        If inputPayload(runningPayload).VolumeColor = Color.Green Then
                            signal = 1
                        ElseIf inputPayload(runningPayload).VolumeColor = Color.Red Then
                            signal = -1
                        End If
                        entryPrice = inputPayload(runningPayload).Close
                        nextCandleExecution = True
                        oncePerDay = True
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