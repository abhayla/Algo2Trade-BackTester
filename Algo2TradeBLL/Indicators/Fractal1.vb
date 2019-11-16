Namespace Indicator
    Public Module Fractal1
        Public Sub CalculateFractal1(ByVal wickPercentage As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Integer))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
                Fractals.CalculateFractal(5, inputPayload, fractalHighPayload, fractalLowPayload)
                For Each runningPayload In inputPayload.Keys
                    Dim signalValue As Integer = 0

                    If inputPayload(runningPayload).CandleStrength = Payload.StrongCandle.Bullish Then
                        If inputPayload(runningPayload).Close <= fractalHighPayload(runningPayload) AndAlso
                                inputPayload(runningPayload).High >= fractalHighPayload(runningPayload) AndAlso
                                inputPayload(runningPayload).CandleWicks.Top > inputPayload(runningPayload).CandleRange * wickPercentage / 100 Then
                            signalValue = 1
                        End If
                    ElseIf inputPayload(runningPayload).CandleStrength = Payload.StrongCandle.Bearish Then
                        If inputPayload(runningPayload).Close >= fractalLowPayload(runningPayload) AndAlso
                                inputPayload(runningPayload).Low <= fractalLowPayload(runningPayload) AndAlso
                                inputPayload(runningPayload).CandleWicks.Bottom > inputPayload(runningPayload).CandleRange * wickPercentage / 100 Then
                            signalValue = -1
                        End If
                    End If
                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Integer)
                    outputPayload.Add(runningPayload, signalValue)
                Next
            End If
        End Sub
    End Module
End Namespace
