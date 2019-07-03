Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module InTheTrendRule
        Public Sub CalculateInTheTrendRule(ByVal atrShift As Integer, ByVal atrPeriod As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef vwapPayload As Dictionary(Of Date, Decimal), ByRef highEntryPayload As Dictionary(Of Date, Decimal), ByRef lowEntryPayload As Dictionary(Of Date, Decimal), ByRef signalPayload As Dictionary(Of Date, Integer), ByRef pearsingCandleSignal As Dictionary(Of Date, Integer))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Indicator.VWAP.CalculateVWAP(inputPayload, vwapPayload)
                Dim highATRBandPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim lowATRBandPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.ATRBands.CalculateATRBands(atrShift, atrPeriod, Payload.PayloadFields.Close, inputPayload, highATRBandPayload, lowATRBandPayload)
                Dim highEntryPrice As Decimal = Decimal.MaxValue
                Dim lowEntryPrice As Decimal = Decimal.MinValue
                Dim signalValue As Integer = 0
                Dim pearsingSignal As Integer = 0
                Dim firstSignal As Boolean = True
                For Each runningPayload In inputPayload.Keys
                    signalValue = 0
                    pearsingSignal = 0
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                        runningPayload.Date <> inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date Then
                        highEntryPrice = Decimal.MaxValue
                        lowEntryPrice = Decimal.MinValue
                    End If

                    If highATRBandPayload(runningPayload) > vwapPayload(runningPayload) Then
                        If Not firstSignal AndAlso
                            lowATRBandPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) < vwapPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            lowATRBandPayload(runningPayload) > vwapPayload(runningPayload) Then
                            highEntryPrice = highATRBandPayload(runningPayload)
                        Else
                            highEntryPrice = Math.Min(highEntryPrice, highATRBandPayload(runningPayload))
                        End If
                    Else
                        highEntryPrice = vwapPayload(runningPayload) + (highATRBandPayload(runningPayload) - lowATRBandPayload(runningPayload))
                        signalValue = 1
                    End If
                    If lowATRBandPayload(runningPayload) < vwapPayload(runningPayload) Then
                        If Not firstSignal AndAlso
                            highATRBandPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) > vwapPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso
                            highATRBandPayload(runningPayload) < vwapPayload(runningPayload) Then
                            lowEntryPrice = lowATRBandPayload(runningPayload)
                        Else
                            lowEntryPrice = Math.Max(lowEntryPrice, lowATRBandPayload(runningPayload))
                        End If
                    Else
                        lowEntryPrice = vwapPayload(runningPayload) - (highATRBandPayload(runningPayload) - lowATRBandPayload(runningPayload))
                        signalValue = -1
                    End If

                    If inputPayload(runningPayload).High > lowATRBandPayload(runningPayload) AndAlso
                        inputPayload(runningPayload).Low < lowATRBandPayload(runningPayload) Then
                        pearsingSignal = 1
                    ElseIf inputPayload(runningPayload).High > highATRBandPayload(runningPayload) AndAlso
                            inputPayload(runningPayload).Low < highATRBandPayload(runningPayload) Then
                        pearsingSignal = -1
                    End If

                    firstSignal = False
                    If highEntryPayload Is Nothing Then highEntryPayload = New Dictionary(Of Date, Decimal)
                    highEntryPayload.Add(runningPayload, highEntryPrice)
                    If lowEntryPayload Is Nothing Then lowEntryPayload = New Dictionary(Of Date, Decimal)
                    lowEntryPayload.Add(runningPayload, lowEntryPrice)
                    If signalPayload Is Nothing Then signalPayload = New Dictionary(Of Date, Integer)
                    signalPayload.Add(runningPayload, signalValue)
                    If pearsingCandleSignal Is Nothing Then pearsingCandleSignal = New Dictionary(Of Date, Integer)
                    pearsingCandleSignal.Add(runningPayload, pearsingSignal)
                Next
            End If
        End Sub
    End Module
End Namespace