Imports System.Drawing
Namespace Indicator
    Public Module FractalWithSMA
        Public Sub CalculateFractalWithSMA(ByVal smaPeriod As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputSignalPayload As Dictionary(Of Date, Color), ByRef outputEntryPricePayload As Dictionary(Of Date, Double), ByRef outputStoplossPricePayload As Dictionary(Of Date, Double))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim SMAPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim FractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim FractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
                SMA.CalculateSMA(smaPeriod, Payload.PayloadFields.Close, inputPayload, SMAPayload)
                Fractals.CalculateFractal(5, inputPayload, FractalHighPayload, FractalLowPayload)
                Dim tempSignal As Boolean = False
                Dim mainSignal As Boolean = False
                Dim signalColor As Color = Color.Black
                Dim entryPrice As Double = 0
                Dim stopLossPrice As Double = 0
                Dim highPrice As Double = 0
                Dim lowPrice As Double = 0
                For Each runningPayload In inputPayload.Keys
                    If SMAPayload(runningPayload) > FractalHighPayload(runningPayload) AndAlso
                            SMAPayload(runningPayload) > FractalLowPayload(runningPayload) Then
                        signalColor = Color.Red
                    ElseIf SMAPayload(runningPayload) < FractalHighPayload(runningPayload) AndAlso
                            SMAPayload(runningPayload) < FractalLowPayload(runningPayload) Then
                        signalColor = Color.Green
                    End If

                    If FractalHighPayload(runningPayload) > SMAPayload(runningPayload) Then
                        highPrice = FractalHighPayload(runningPayload)
                    End If
                    If FractalLowPayload(runningPayload) < SMAPayload(runningPayload) Then
                        lowPrice = FractalLowPayload(runningPayload)
                    End If

                    If signalColor = Color.Green Then
                        If FractalHighPayload(runningPayload) < inputPayload(runningPayload).High Then
                            If outputSignalPayload IsNot Nothing AndAlso outputSignalPayload.Count > 0 AndAlso
                                    outputSignalPayload.ContainsKey(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                If outputSignalPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = Color.Green Then
                                    entryPrice = Math.Min(entryPrice, inputPayload(runningPayload).High)
                                Else
                                    entryPrice = inputPayload(runningPayload).High
                                End If
                            Else
                                entryPrice = inputPayload(runningPayload).High
                            End If
                        Else
                            entryPrice = FractalHighPayload(runningPayload)
                        End If
                        If FractalLowPayload(runningPayload) < SMAPayload(runningPayload) Then
                            stopLossPrice = FractalLowPayload(runningPayload)
                        Else
                            stopLossPrice = lowPrice
                        End If
                    ElseIf signalColor = Color.Red Then
                        If FractalLowPayload(runningPayload) > inputPayload(runningPayload).Low Then
                            If outputSignalPayload IsNot Nothing AndAlso outputSignalPayload.Count > 0 AndAlso
                                    outputSignalPayload.ContainsKey(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) Then
                                If outputSignalPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) = Color.Red Then
                                    entryPrice = Math.Max(entryPrice, inputPayload(runningPayload).Low)
                                Else
                                    entryPrice = inputPayload(runningPayload).Low
                                End If
                            Else
                                entryPrice = inputPayload(runningPayload).Low
                            End If
                        Else
                            entryPrice = FractalLowPayload(runningPayload)
                        End If
                        If FractalHighPayload(runningPayload) > SMAPayload(runningPayload) Then
                            stopLossPrice = FractalHighPayload(runningPayload)
                        Else
                            stopLossPrice = highPrice
                        End If
                    End If

                    If outputSignalPayload Is Nothing Then outputSignalPayload = New Dictionary(Of Date, Color)
                    outputSignalPayload.Add(runningPayload, signalColor)
                    If outputEntryPricePayload Is Nothing Then outputEntryPricePayload = New Dictionary(Of Date, Double)
                    outputEntryPricePayload.Add(runningPayload, entryPrice)
                    If outputStoplossPricePayload Is Nothing Then outputStoplossPricePayload = New Dictionary(Of Date, Double)
                    outputStoplossPricePayload.Add(runningPayload, stopLossPrice)
                Next
            End If
        End Sub
    End Module
End Namespace