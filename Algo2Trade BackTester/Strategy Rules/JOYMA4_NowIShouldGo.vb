Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module JOYMA4_NowIShouldGo
        Public Sub CalculateJOYMA4_NowIShouldGo(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef buyPayload As Dictionary(Of Date, Decimal), ByRef sellPayload As Dictionary(Of Date, Decimal), ByRef typePayload As Dictionary(Of Date, Integer), ByRef targetPayload As Dictionary(Of Date, Decimal), ByRef slPayload As Dictionary(Of Date, Decimal), ByRef remarksPayload As Dictionary(Of Date, String))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim atrPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATR.CalculateATR(14, inputPayload, atrPayload)
                Dim buyPrice As Decimal = Nothing
                Dim sellPrice As Decimal = Nothing
                Dim targetPoint As Decimal = Nothing
                Dim slPoint As Decimal = Nothing
                Dim remarks As String = Nothing
                Dim previousType As Integer = 0
                For Each runningPayload In inputPayload.Keys
                    Dim type As Integer = 0
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing Then
                        If inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date <> runningPayload.Date Then
                            buyPrice = 0
                            sellPrice = 0
                            targetPoint = 0
                            slPoint = 0
                        Else
                            Dim currentCandleVolume As Decimal = Math.Round(inputPayload(runningPayload).Volume / 1000, 1)
                            Dim previousCandleVolume As Decimal = Math.Round(inputPayload(runningPayload).PreviousCandlePayload.Volume / 1000, 1)
                            If currentCandleVolume >= 2 * previousCandleVolume OrElse currentCandleVolume <= 0.5 * previousCandleVolume Then
                                If inputPayload(runningPayload).CandleRange > 3 Then
                                    If inputPayload(runningPayload).DojiCandle OrElse
                                        inputPayload(runningPayload).CandleWicks.Top >= inputPayload(runningPayload).CandleRange * 50 / 100 OrElse
                                        inputPayload(runningPayload).CandleWicks.Bottom >= inputPayload(runningPayload).CandleRange * 50 / 100 Then
                                        Dim allowableCR As Decimal = atrPayload(runningPayload) * 66.7 / 100
                                        If inputPayload(runningPayload).CandleRange <= allowableCR Then
                                            type = 1
                                            buyPrice = inputPayload(runningPayload).High
                                            sellPrice = inputPayload(runningPayload).Low
                                            previousType = type
                                        ElseIf inputPayload(runningPayload).CandleRange <= atrPayload(runningPayload) Then
                                            type = 2
                                            buyPrice = inputPayload(runningPayload).High
                                            sellPrice = inputPayload(runningPayload).Low
                                            previousType = type
                                        Else
                                            remarks = "Greater Than ATR"
                                        End If
                                    Else
                                        remarks = "Doji/Wick range Not Satisfying"
                                    End If
                                Else
                                    remarks = "Candle range Not Satisfying"
                                End If
                            Else
                                remarks = "Volume Not Satisfying"
                            End If
                            If type = 1 Then
                                targetPoint = atrPayload(runningPayload) + 2
                                remarks = ""
                            ElseIf type = 2 Then
                                targetPoint = atrPayload(runningPayload) * 1.5 + 2
                                remarks = ""
                            Else
                                If previousType = 1 Then
                                    targetPoint = atrPayload(runningPayload) + 2
                                    remarks = ""
                                ElseIf previousType = 2 Then
                                    targetPoint = atrPayload(runningPayload) * 1.5 + 2
                                    remarks = ""
                                End If
                            End If
                            slPoint = Math.Abs(buyPrice - sellPrice)
                        End If
                    End If
                    If buyPayload Is Nothing Then buyPayload = New Dictionary(Of Date, Decimal)
                    buyPayload.Add(runningPayload, buyPrice)
                    If sellPayload Is Nothing Then sellPayload = New Dictionary(Of Date, Decimal)
                    sellPayload.Add(runningPayload, sellPrice)
                    If typePayload Is Nothing Then typePayload = New Dictionary(Of Date, Integer)
                    typePayload.Add(runningPayload, type)
                    If targetPayload Is Nothing Then targetPayload = New Dictionary(Of Date, Decimal)
                    targetPayload.Add(runningPayload, targetPoint)
                    If slPayload Is Nothing Then slPayload = New Dictionary(Of Date, Decimal)
                    slPayload.Add(runningPayload, slPoint)
                    If remarksPayload Is Nothing Then remarksPayload = New Dictionary(Of Date, String)
                    remarksPayload.Add(runningPayload, remarks)
                Next
            End If
        End Sub
    End Module
End Namespace
