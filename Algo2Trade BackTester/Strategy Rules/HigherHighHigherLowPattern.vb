Imports Algo2TradeBLL

Namespace StrategyRules
    Module HigherHighHigherLowPattern
        Public Sub CalculateHigherHighHigherLowPattern(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef signalPayload As Dictionary(Of Date, Integer))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim highSignal1 As Integer = Integer.MaxValue
                    Dim lowSignal1 As Integer = Integer.MaxValue
                    Dim highSignal2 As Integer = Integer.MaxValue
                    Dim lowSignal2 As Integer = Integer.MaxValue
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                        If inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = runningPayload.Date Then
                            GetSignal(inputPayload(runningPayload), highSignal1, lowSignal1)
                            If (highSignal1 = 1 AndAlso lowSignal1 = -1) OrElse (highSignal1 = -1 AndAlso lowSignal1 = 1) Then
                                GoTo label
                            Else
                                GetSignal(inputPayload(runningPayload).PreviousCandlePayload, highSignal2, lowSignal2)
                                If (highSignal2 = 1 AndAlso lowSignal2 = -1) OrElse (highSignal2 = -1 AndAlso lowSignal2 = 1) Then
                                    GoTo label
                                Else
                                    If highSignal1 = 1 OrElse lowSignal1 = 1 Then
                                        If highSignal2 = 1 OrElse lowSignal2 = 1 OrElse (highSignal2 = 0 AndAlso lowSignal2 = 0) Then
                                            signal = 1
                                        End If
                                    ElseIf highSignal1 = -1 OrElse lowSignal1 = -1 Then
                                        If highSignal2 = -1 OrElse lowSignal2 = -1 OrElse (highSignal2 = 0 AndAlso lowSignal2 = 0) Then
                                            signal = -1
                                        End If
                                    ElseIf highSignal1 = 0 AndAlso lowSignal1 = 0 Then
                                        If highSignal2 = 1 OrElse lowSignal2 = 1 Then
                                            signal = 1
                                        ElseIf highSignal2 = -1 OrElse lowSignal2 = -1 Then
                                            signal = -1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
label:              If signalPayload Is Nothing Then signalPayload = New Dictionary(Of Date, Integer)
                    signalPayload.Add(runningPayload, signal)
                Next
            End If
        End Sub
        Private Sub GetSignal(ByVal currentPayload As Payload, ByRef highSignal As Integer, ByRef lowSignal As Integer)
            highSignal = Integer.MaxValue
            lowSignal = Integer.MaxValue
            If currentPayload IsNot Nothing AndAlso currentPayload.PreviousCandlePayload IsNot Nothing Then
                If currentPayload.High > currentPayload.PreviousCandlePayload.High Then
                    highSignal = -1
                ElseIf currentPayload.High < currentPayload.PreviousCandlePayload.High Then
                    highSignal = 1
                Else
                    highSignal = 0
                End If
                If currentPayload.Low > currentPayload.PreviousCandlePayload.Low Then
                    lowSignal = -1
                ElseIf currentPayload.Low < currentPayload.PreviousCandlePayload.Low Then
                    lowSignal = 1
                Else
                    lowSignal = 0
                End If
            End If
        End Sub
    End Module
End Namespace
