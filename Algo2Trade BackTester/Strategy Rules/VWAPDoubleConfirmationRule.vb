Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module VWAPDoubleConfirmationRule
        Public Sub CalculateNaughtyBoyVWAPRule(ByVal ATRShift As Integer, ByVal ATRPeriod As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputStoplossPricePayload As Dictionary(Of Date, Double))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim VWAPOutputPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.VWAP.CalculateVWAP(inputPayload, VWAPOutputPayload)
                Dim highATRBand As Dictionary(Of Date, Decimal) = Nothing
                Dim lowATRBand As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATRBands.CalculateATRBands(ATRShift, ATRPeriod, Payload.PayloadFields.Close, inputPayload, highATRBand, lowATRBand)

                Dim buyStoplossPrice As Double = Nothing
                Dim sellStoplossPrice As Double = Nothing
                Dim firstCandle As Boolean = True
                Dim signalValue As Integer = 0
                For Each runningPayload In inputPayload.Keys
                    Dim vwapValue As Decimal = VWAPOutputPayload(runningPayload)

                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso runningPayload.Date <> inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date Then
                        firstCandle = True
                        signalValue = 0
                        buyStoplossPrice = 0
                        sellStoplossPrice = 0
                    End If

                    If VWAPOutputPayload(runningPayload) > lowATRBand(runningPayload) Then
                        buyStoplossPrice = lowATRBand(runningPayload)
                    End If
                    If VWAPOutputPayload(runningPayload) < highATRBand(runningPayload) Then
                        sellStoplossPrice = highATRBand(runningPayload)
                    End If

                    If Not firstCandle Then
                        If inputPayload(runningPayload).PreviousCandlePayload.High > VWAPOutputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            inputPayload(runningPayload).PreviousCandlePayload.Low < VWAPOutputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                            If inputPayload(runningPayload).PreviousCandlePayload.Close > VWAPOutputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                inputPayload(runningPayload).Close > VWAPOutputPayload(inputPayload(runningPayload).PayloadDate) AndAlso
                                inputPayload(runningPayload).High > VWAPOutputPayload(inputPayload(runningPayload).PayloadDate) AndAlso
                                inputPayload(runningPayload).Low > VWAPOutputPayload(inputPayload(runningPayload).PayloadDate) Then

                                signalValue = 1

                            ElseIf inputPayload(runningPayload).PreviousCandlePayload.Close < VWAPOutputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                                inputPayload(runningPayload).Close < VWAPOutputPayload(inputPayload(runningPayload).PayloadDate) AndAlso
                                inputPayload(runningPayload).High < VWAPOutputPayload(inputPayload(runningPayload).PayloadDate) AndAlso
                                inputPayload(runningPayload).Low < VWAPOutputPayload(inputPayload(runningPayload).PayloadDate) Then

                                signalValue = -1

                            Else
                                signalValue = 0
                            End If
                        Else
                            signalValue = 0
                        End If
                    End If

                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                    outputSignalPayload.Add(runningPayload, signalValue)
                    If outputStoplossPricePayload Is Nothing Then outputStoplossPricePayload = New Dictionary(Of Date, Double)
                    outputStoplossPricePayload.Add(runningPayload, If(signalValue > 0, buyStoplossPrice, If(signalValue < 0, sellStoplossPrice, 0)))
                    firstCandle = False
                Next
            End If
        End Sub
    End Module
End Namespace
