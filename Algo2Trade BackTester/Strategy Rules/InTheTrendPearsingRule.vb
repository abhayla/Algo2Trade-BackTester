Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module InTheTrendPearsingRule
        Public Sub CalculateInTheTrendPearsingRule(ByVal atrShift As Integer, ByVal atrPeriod As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef upperBandPayload As Dictionary(Of Date, Decimal), ByRef lowerBandPayload As Dictionary(Of Date, Decimal), ByRef signalPayload As Dictionary(Of Date, Integer))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim vwapPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.VWAP.CalculateVWAP(inputPayload, vwapPayload)
                Dim highATRBandPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim lowATRBandPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATRBands.CalculateATRBands(atrShift, atrPeriod, Payload.PayloadFields.Close, inputPayload, highATRBandPayload, lowATRBandPayload)
                Dim upperBand As Decimal = Decimal.MaxValue
                Dim lowerBand As Decimal = Decimal.MinValue
                Dim signalValue As Integer = 0
                Dim firstSignal As Boolean = True
                For Each runningPayload In inputPayload.Keys
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                        runningPayload.Date <> inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date Then
                        signalValue = 0
                    End If

                    If inputPayload(runningPayload).High >= highATRBandPayload(runningPayload) AndAlso
                        inputPayload(runningPayload).Low <= lowATRBandPayload(runningPayload) AndAlso
                        inputPayload(runningPayload).High >= vwapPayload(runningPayload) AndAlso
                        inputPayload(runningPayload).Low <= vwapPayload(runningPayload) Then
                        If highATRBandPayload(runningPayload) >= vwapPayload(runningPayload) AndAlso
                            lowATRBandPayload(runningPayload) >= vwapPayload(runningPayload) Then
                            If signalValue <= 0 Then
                                upperBand = Decimal.MaxValue
                                lowerBand = Decimal.MinValue
                            End If
                            signalValue = 1
                        ElseIf highATRBandPayload(runningPayload) <= vwapPayload(runningPayload) AndAlso
                                lowATRBandPayload(runningPayload) <= vwapPayload(runningPayload) Then
                            If signalValue >= 0 Then
                                upperBand = Decimal.MaxValue
                                lowerBand = Decimal.MinValue
                            End If
                            signalValue = -1
                        Else
                            signalValue = 0
                        End If
                    ElseIf inputPayload(runningPayload).High >= lowATRBandPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Low <= lowATRBandPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).High >= vwapPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Low <= vwapPayload(runningPayload) Then
                        If signalValue <= 0 Then
                            upperBand = Decimal.MaxValue
                            lowerBand = Decimal.MinValue
                        End If
                        signalValue = 1
                    ElseIf inputPayload(runningPayload).High >= highATRBandPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Low <= highATRBandPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).High >= vwapPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Low <= vwapPayload(runningPayload) Then
                        If signalValue >= 0 Then
                            upperBand = Decimal.MaxValue
                            lowerBand = Decimal.MinValue
                        End If
                        signalValue = -1
                    End If
                    If signalValue < 0 AndAlso highATRBandPayload(runningPayload) >= vwapPayload(runningPayload) AndAlso
                        lowATRBandPayload(runningPayload) >= vwapPayload(runningPayload) Then
                        signalValue = 0
                    ElseIf signalValue > 0 AndAlso highATRBandPayload(runningPayload) <= vwapPayload(runningPayload) AndAlso
                            lowATRBandPayload(runningPayload) <= vwapPayload(runningPayload) Then
                        signalValue = 0
                    End If

                    If signalValue > 0 OrElse signalValue < 0 Then
                        If highATRBandPayload(runningPayload) > vwapPayload(runningPayload) Then
                            upperBand = Math.Min(upperBand, highATRBandPayload(runningPayload))
                        Else
                            upperBand = vwapPayload(runningPayload) + (highATRBandPayload(runningPayload) - lowATRBandPayload(runningPayload))
                        End If

                        If lowATRBandPayload(runningPayload) < vwapPayload(runningPayload) Then
                            lowerBand = Math.Max(lowerBand, lowATRBandPayload(runningPayload))
                        Else
                            lowerBand = vwapPayload(runningPayload) - (highATRBandPayload(runningPayload) - lowATRBandPayload(runningPayload))
                        End If
                    Else
                        upperBand = Decimal.MaxValue
                        lowerBand = Decimal.MinValue
                    End If

                    firstSignal = False
                    If upperBandPayload Is Nothing Then upperBandPayload = New Dictionary(Of Date, Decimal)
                    upperBandPayload.Add(runningPayload, If(upperBand = Decimal.MaxValue, 0, upperBand))
                    If lowerBandPayload Is Nothing Then lowerBandPayload = New Dictionary(Of Date, Decimal)
                    lowerBandPayload.Add(runningPayload, If(lowerBand = Decimal.MinValue, 0, lowerBand))
                    If signalPayload Is Nothing Then signalPayload = New Dictionary(Of Date, Integer)
                    signalPayload.Add(runningPayload, signalValue)
                Next
            End If
        End Sub
    End Module
End Namespace
