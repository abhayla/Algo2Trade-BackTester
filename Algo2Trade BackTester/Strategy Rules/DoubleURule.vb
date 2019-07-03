Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module DoubleURule
        Public Sub CalculateDoubleURule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.Fractals.CalculateFractal(5, inputPayload, fractalHighPayload, fractalLowPayload)
                Dim firstFractalHigh As Decimal = 0
                Dim secondFractalHigh As Decimal = 0
                Dim thirdFractalHigh As Decimal = 0
                Dim firstFractalLow As Decimal = 0
                Dim secondFractalLow As Decimal = 0
                Dim thirdFractalLow As Decimal = 0
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
                        firstCandle = True
                    End If
                    If Not firstCandle Then
                        If fractalHighPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) <> thirdFractalHigh Then
                            firstFractalHigh = secondFractalHigh
                            secondFractalHigh = thirdFractalHigh
                            thirdFractalHigh = fractalHighPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                        End If
                        If fractalLowPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) <> thirdFractalLow Then
                            firstFractalLow = secondFractalLow
                            secondFractalLow = thirdFractalLow
                            thirdFractalLow = fractalLowPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                        End If

                        If thirdFractalHigh > secondFractalHigh AndAlso secondFractalHigh < firstFractalHigh AndAlso
                        thirdFractalLow > secondFractalLow AndAlso secondFractalLow < firstFractalLow Then
                            If fractalHighPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = thirdFractalHigh AndAlso fractalLowPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = thirdFractalLow Then
                                If inputPayload(runningPayload).High >= fractalHighPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                    signal = 1
                                End If
                            End If
                        End If
                        If thirdFractalHigh < secondFractalHigh AndAlso secondFractalHigh > firstFractalHigh AndAlso
                            thirdFractalLow < secondFractalLow AndAlso secondFractalLow > firstFractalLow Then
                            If fractalHighPayload(runningPayload) = thirdFractalHigh AndAlso fractalLowPayload(runningPayload) = thirdFractalLow Then
                                If inputPayload(runningPayload).Low < fractalLowPayload(runningPayload) Then
                                    signal = -1
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
    End Module
End Namespace
