Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module FractalRetracementEntryRule
        Public Sub CalculateFractalRetracementEntry(ByVal smaPeriod As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPricePayload As Dictionary(Of Date, Double))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim fractalWithSMASignalPayload As Dictionary(Of Date, Color) = Nothing
                Dim dummyPayload1 As Dictionary(Of Date, Double) = Nothing
                Dim dummyPayload2 As Dictionary(Of Date, Double) = Nothing
                FractalWithSMA.CalculateFractalWithSMA(smaPeriod, inputPayload, fractalWithSMASignalPayload, dummyPayload1, dummyPayload2)
                Dim entryPrice As Double = Nothing
                Dim firstCandle As Boolean = True
                For Each runningPayload In inputPayload.Keys
                    Dim signalValue As Double = 0
                    If Not firstCandle Then
                        If (fractalWithSMASignalPayload(runningPayload) <> fractalWithSMASignalPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)) OrElse
                        (runningPayload.Date <> inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date) Then
                            entryPrice = If(fractalWithSMASignalPayload(runningPayload) = Color.Green, Double.MaxValue, If(fractalWithSMASignalPayload(runningPayload) = Color.Red, Double.MinValue, 0))
                            GoTo label
                        End If
                    End If
                    If fractalWithSMASignalPayload(runningPayload) = Color.Green Then
                        If inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bearish Then
                            entryPrice = Math.Min(inputPayload(runningPayload).High, entryPrice)
                            signalValue = 1
                        End If
                    ElseIf fractalWithSMASignalPayload(runningPayload) = Color.Red Then
                        If inputPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bullish Then
                            entryPrice = Math.Max(inputPayload(runningPayload).Low, entryPrice)
                            signalValue = -1
                        End If
                    End If
label:              If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                    outputSignalPayload.Add(runningPayload, signalValue)
                    If outputEntryPricePayload Is Nothing Then outputEntryPricePayload = New Dictionary(Of Date, Double)
                    outputEntryPricePayload.Add(runningPayload, If(signalValue = 1, entryPrice, If(signalValue = -1, entryPrice, 0)))
                    firstCandle = False
                Next
            End If
        End Sub
    End Module
End Namespace
