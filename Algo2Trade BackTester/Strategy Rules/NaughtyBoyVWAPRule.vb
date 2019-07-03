Imports Algo2TradeBLL

Namespace StrategyRules
    Public Module NaughtyBoyVWAPRule
        Public Sub CalculateNaughtyBoyVWAPRule(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPricePayload As Dictionary(Of Date, Double), ByRef outputStoplossPricePayload As Dictionary(Of Date, Double))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim VWAPOutputPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.VWAP.CalculateVWAP(inputPayload, VWAPOutputPayload)

                Dim entryPrice As Double = Nothing
                Dim stoplossPrice As Double = Nothing
                Dim firstCandle As Boolean = True
                Dim signalValue As Integer = 0
                For Each runningPayload In inputPayload.Keys
                    Dim vwapValue As Decimal = VWAPOutputPayload(runningPayload)

                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso runningPayload.Date <> inputPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date Then
                        firstCandle = True
                        signalValue = 0
                        entryPrice = 0
                        stoplossPrice = 0
                    End If

                    If Not firstCandle Then
                        If inputPayload(runningPayload).PreviousCandlePayload.High > VWAPOutputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) AndAlso inputPayload(runningPayload).PreviousCandlePayload.Low < VWAPOutputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                            If inputPayload(runningPayload).Close > vwapValue Then
                                signalValue = 1
                                entryPrice = inputPayload(runningPayload).High
                                stoplossPrice = vwapValue
                            ElseIf inputPayload(runningPayload).Close < vwapValue Then
                                signalValue = -1
                                entryPrice = inputPayload(runningPayload).Low
                                stoplossPrice = vwapValue
                            End If
                        Else
                            If signalValue > 0 Then
                                entryPrice = inputPayload(runningPayload).High
                                stoplossPrice = vwapValue
                                If stoplossPrice > entryPrice Then
                                    signalValue = -10
                                    entryPrice = inputPayload(runningPayload).Low
                                    stoplossPrice = vwapValue
                                End If
                            ElseIf signalValue < 0 Then
                                entryPrice = inputPayload(runningPayload).Low
                                stoplossPrice = vwapValue
                                If stoplossPrice < entryPrice Then
                                    signalValue = 10
                                    entryPrice = inputPayload(runningPayload).High
                                    stoplossPrice = vwapValue
                                End If
                            End If
                        End If
                    End If

                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                    outputSignalPayload.Add(runningPayload, signalValue)
                    If outputEntryPricePayload Is Nothing Then outputEntryPricePayload = New Dictionary(Of Date, Double)
                    outputEntryPricePayload.Add(runningPayload, If(signalValue > 0, entryPrice, If(signalValue < 0, entryPrice, 0)))
                    If outputStoplossPricePayload Is Nothing Then outputStoplossPricePayload = New Dictionary(Of Date, Double)
                    outputStoplossPricePayload.Add(runningPayload, If(signalValue > 0, stoplossPrice, If(signalValue < 0, stoplossPrice, 0)))
                    firstCandle = False
                Next
            End If
        End Sub
    End Module
End Namespace
