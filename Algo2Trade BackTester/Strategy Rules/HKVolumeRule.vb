Imports Algo2TradeBLL
Imports Utilities.Numbers

Namespace StrategyRules
    Public Module HKVolumeRule
        Public Sub CalculateRule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPayload As Dictionary(Of Date, Decimal), ByRef outputTargetPayload As Dictionary(Of Date, Decimal), ByRef outputStoplossPayload As Dictionary(Of Date, Decimal))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim emaPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, inputPayload, emaPayload)
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    Dim entryPrice As Decimal = 0
                    Dim targetPrice As Decimal = 0
                    Dim slPrice As Decimal = 0

                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                        If inputPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bullish AndAlso
                            NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Close, 2) >= NumberManipulation.RoundEX(emaPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Open, 2) <= NumberManipulation.RoundEX(emaPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(inputPayload(runningPayload).Low, 2) < NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Low, 2) AndAlso
                            inputPayload(runningPayload).PreviousCandlePayload.Volume > inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume * 2 Then
                            signal = -1
                            entryPrice = NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Low, 2)
                            slPrice = NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.High, 2)
                            targetPrice = If((slPrice / entryPrice) - 1 <= 0.003, entryPrice * 0.997, entryPrice - (slPrice - entryPrice))
                        ElseIf inputPayload(runningPayload).PreviousCandlePayload.CandleStrengthHK = Payload.StrongCandle.Bearish AndAlso
                            NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Open, 2) >= NumberManipulation.RoundEX(emaPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Close, 2) <= NumberManipulation.RoundEX(emaPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate), 2) AndAlso
                            NumberManipulation.RoundEX(inputPayload(runningPayload).High, 2) > NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.High, 2) AndAlso
                            inputPayload(runningPayload).PreviousCandlePayload.Volume > inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Volume * 2 Then
                            signal = 1
                            entryPrice = NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.High, 2)
                            slPrice = NumberManipulation.RoundEX(inputPayload(runningPayload).PreviousCandlePayload.Low, 2)
                            targetPrice = If(Math.Abs((slPrice / entryPrice) - 1) <= 0.003, entryPrice * 1.003, entryPrice + (entryPrice - slPrice))
                        End If
                    End If

                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                        If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                        outputSignalPayload.Add(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, signal)
                        If outputEntryPayload Is Nothing Then outputEntryPayload = New Dictionary(Of Date, Decimal)
                        outputEntryPayload.Add(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, entryPrice)
                        If outputTargetPayload Is Nothing Then outputTargetPayload = New Dictionary(Of Date, Decimal)
                        outputTargetPayload.Add(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, targetPrice)
                        If outputStoplossPayload Is Nothing Then outputStoplossPayload = New Dictionary(Of Date, Decimal)
                        outputStoplossPayload.Add(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate, slPrice)
                    End If
                Next
            End If
        End Sub
    End Module
End Namespace
