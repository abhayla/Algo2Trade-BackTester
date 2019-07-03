Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module MRWithOneFractalAndMA
        Public Sub CalculateMRWithOneFractalAndMA(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.Fractals.CalculateFractal(5, inputPayload, fractalHighPayload, fractalLowPayload)
                Dim MAPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.SMA.CalculateSMA(50, Payload.PayloadFields.Close, inputPayload, MAPayload)
                Dim firstFractalHigh As Decimal = 0
                Dim secondFractalHigh As Decimal = 0
                Dim thirdFractalHigh As Decimal = 0
                Dim firstFractalLow As Decimal = 0
                Dim secondFractalLow As Decimal = 0
                Dim thirdFractalLow As Decimal = 0
                Dim firstFractalHighTime As Date = Nothing
                Dim secondFractalHighTime As Date = Nothing
                Dim thirdFractalHighTime As Date = Nothing
                Dim firstFractalLowTime As Date = Nothing
                Dim secondFractalLowTime As Date = Nothing
                Dim thirdFractalLowTime As Date = Nothing

                Dim firstCandle As Boolean = True
                For Each runningPayload In inputPayload.Keys
                    Dim signal As Integer = 0
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso runningPayload.Date <> inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date Then
                        firstFractalHigh = 0
                        secondFractalHigh = 0
                        thirdFractalHigh = 0
                        firstFractalLow = 0
                        secondFractalLow = 0
                        thirdFractalLow = 0
                    End If
                    If Not firstCandle Then
                        If fractalHighPayload(runningPayload) <> thirdFractalHigh Then
                            firstFractalHigh = secondFractalHigh
                            secondFractalHigh = thirdFractalHigh
                            thirdFractalHigh = fractalHighPayload(runningPayload)
                            firstFractalHighTime = secondFractalHighTime
                            secondFractalHighTime = thirdFractalHighTime
                            thirdFractalHighTime = runningPayload
                        End If
                        If fractalLowPayload(runningPayload) <> thirdFractalLow Then
                            firstFractalLow = secondFractalLow
                            secondFractalLow = thirdFractalLow
                            thirdFractalLow = fractalLowPayload(runningPayload)
                            firstFractalLowTime = secondFractalLowTime
                            secondFractalLowTime = thirdFractalLowTime
                            thirdFractalLowTime = runningPayload
                        End If

                        If firstFractalHigh <> 0 AndAlso thirdFractalHigh > secondFractalHigh AndAlso secondFractalHigh < firstFractalHigh Then
                            If fractalHighPayload(runningPayload) = thirdFractalHigh Then
                                If thirdFractalHigh = firstFractalHigh Then
                                    If Not CheckCrossOver(1, inputPayload, MAPayload, firstFractalHigh, secondFractalHigh, thirdFractalHigh, secondFractalHighTime, thirdFractalHighTime) Then
                                        signal = 2
                                    End If
                                ElseIf thirdFractalHigh > firstFractalHigh Then
                                    If Not CheckCrossOver(1, inputPayload, MAPayload, firstFractalHigh, secondFractalHigh, thirdFractalHigh, secondFractalHighTime, thirdFractalHighTime) Then
                                        signal = 1
                                    End If
                                End If
                            End If
                        End If
                        If firstFractalLow <> 0 AndAlso thirdFractalLow < secondFractalLow AndAlso secondFractalLow > firstFractalLow Then
                            If fractalLowPayload(runningPayload) = thirdFractalLow Then
                                If thirdFractalLow = firstFractalLow Then
                                    If Not CheckCrossOver(-1, inputPayload, MAPayload, firstFractalLow, secondFractalLow, thirdFractalLow, secondFractalLowTime, thirdFractalLowTime) Then
                                        signal = -2
                                    End If
                                ElseIf thirdFractalLow < firstFractalLow Then
                                    If Not CheckCrossOver(-1, inputPayload, MAPayload, firstFractalLow, secondFractalLow, thirdFractalLow, secondFractalLowTime, thirdFractalLowTime) Then
                                        signal = -1
                                    End If
                                End If
                            End If
                        End If
                    End If
                    firstCandle = False
                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                    outputSignalPayload.Add(runningPayload, signal)
                Next
            End If
        End Sub
        Public Function CheckCrossOver(ByVal signalDirection As Integer, ByVal inputPayload As Dictionary(Of Date, Payload),
                                       ByVal MAPayload As Dictionary(Of Date, Decimal), ByVal firstFractal As Decimal, ByVal secondFractal As Decimal,
                                       ByVal thirdFractal As Decimal, ByVal secondFractalTime As Date, ByVal thirdFractalTime As Date) As Boolean
            Dim ret As Boolean = True
            Dim ctr As Integer = 0
            'Dim firstPosition As Boolean = False
            'Dim secondPosition As Boolean = False
            'Dim thirdPosition As Boolean = False
            'Dim fourthPosition As Boolean = False
            If signalDirection = 1 Then
                If firstFractal > MAPayload(inputPayload(secondFractalTime).PreviousCandlePayload.PayloadDate) Then
                    ctr += 1
                End If
                If secondFractal > MAPayload(secondFractalTime) Then
                    ctr += 1
                End If
                If secondFractal > MAPayload(inputPayload(thirdFractalTime).PreviousCandlePayload.PayloadDate) Then
                    ctr += 1
                End If
                If thirdFractal > MAPayload(thirdFractalTime) Then
                    ctr += 1
                End If
                'If Not (firstPosition And secondPosition And thirdPosition And fourthPosition) Then
                ret = ctr >= 4 Or ctr = 0
                'End If
            ElseIf signalDirection = -1 Then
                If firstFractal < MAPayload(inputPayload(secondFractalTime).PreviousCandlePayload.PayloadDate) Then
                    ctr += 1
                End If
                If secondFractal < MAPayload(secondFractalTime) Then
                    ctr += 1
                End If
                If secondFractal < MAPayload(inputPayload(thirdFractalTime).PreviousCandlePayload.PayloadDate) Then
                    ctr += 1
                End If
                If thirdFractal < MAPayload(thirdFractalTime) Then
                    ctr += 1
                End If
                'If Not (firstPosition And secondPosition And thirdPosition And fourthPosition) Then
                ret = ctr >= 4 Or ctr = 0
                'End If
            End If
            Return ret
        End Function
    End Module
End Namespace
