Imports Algo2TradeBLL
Public Module HARetracementEntry
    Public Sub CalculateHARetracementEntry(ByVal inputHAPayload As Dictionary(Of Date, Payload), ByVal signalDirection As SignalDirection, ByRef outputSignalPayload As Dictionary(Of Date, Integer), ByRef outputEntryPricePayload As Dictionary(Of Date, Double))
        If inputHAPayload IsNot Nothing AndAlso inputHAPayload.Count > 0 Then
            Dim entryPrice As Double = Nothing
            Dim firstCandle As Boolean = True
            For Each runningPayload In inputHAPayload.Keys
                Dim signalValue As Double = 0
                If Not firstCandle Then
                    If (runningPayload.Date <> inputHAPayload(runningPayload).PreviousCandlePayload.PayloadDate.Date) Then
                        entryPrice = If(signalDirection = SignalDirection.Buy, Double.MaxValue, If(signalDirection = SignalDirection.Sell, Double.MinValue, 0))
                        GoTo label
                    End If
                End If
                If signalDirection = SignalDirection.Buy Then
                    If inputHAPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bearish Then
                        entryPrice = Math.Min(inputHAPayload(runningPayload).High, entryPrice)
                        signalValue = 1
                    End If
                ElseIf signalDirection = SignalDirection.Sell Then
                    If inputHAPayload(runningPayload).CandleStrengthHK = Payload.StrongCandle.Bullish Then
                        entryPrice = Math.Max(inputHAPayload(runningPayload).Low, entryPrice)
                        signalValue = -1
                    End If
                End If
label:          If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Integer)
                outputSignalPayload.Add(runningPayload, signalValue)
                If outputEntryPricePayload Is Nothing Then outputEntryPricePayload = New Dictionary(Of Date, Double)
                outputEntryPricePayload.Add(runningPayload, If(signalValue = 1, entryPrice, If(signalValue = -1, entryPrice, 0)))
                firstCandle = False
            Next
        End If
    End Sub
End Module
Public Enum SignalDirection
    Buy = 1
    Sell
    None
End Enum
