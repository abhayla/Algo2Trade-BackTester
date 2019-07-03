Imports Algo2TradeBLL

Public Module CloseGap
    Public Function CalculateRule(ByVal currentStockPayload As Dictionary(Of Date, Payload), ByVal farStockPayload As Dictionary(Of Date, Payload)) As Dictionary(Of Date, Payload)
        Dim ret As Dictionary(Of Date, Payload) = Nothing
        If currentStockPayload IsNot Nothing AndAlso currentStockPayload.Count > 0 AndAlso
            farStockPayload IsNot Nothing AndAlso farStockPayload.Count > 0 Then
            Dim lastExistingPayload As Payload = Nothing
            For Each runningPayload In currentStockPayload
                Dim farContractRunningPayload As Payload = Nothing
                If farStockPayload.ContainsKey(runningPayload.Key) Then
                    farContractRunningPayload = farStockPayload(runningPayload.Key)
                Else
                    farContractRunningPayload = lastExistingPayload
                End If
                lastExistingPayload = farContractRunningPayload

                If farContractRunningPayload IsNot Nothing Then
                    If ret Is Nothing Then ret = New Dictionary(Of Date, Payload)
                    ret.Add(runningPayload.Key, New Payload(Payload.CandleDataSource.Calculated) With {.Close = (farContractRunningPayload.Close - runningPayload.Value.Close)})
                End If
            Next
        End If
        Return ret
    End Function
End Module