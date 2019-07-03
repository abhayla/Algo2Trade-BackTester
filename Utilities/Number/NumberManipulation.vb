Namespace Numbers
    Public Module NumberManipulation
#Region "Enum"
        Public Enum RoundOfType
            Floor = 1
            Celing
        End Enum
#End Region

#Region "Public Function"
        Public Function ConvertFloorCeling(ByVal number As Double, ByVal tickSize As Double, ByVal upDown As RoundOfType) As Double
            Dim numberOfDigits As Integer = BitConverter.GetBytes(Decimal.GetBits(tickSize)(3))(2)
            Dim decimalPortion As Double = Math.Round(number - Math.Truncate(number), numberOfDigits)
            Dim normalizedDecimalPortion As Double = decimalPortion * Math.Pow(10, numberOfDigits)
            Dim normalizedConvertedDecimalPortion As Double = Math.Floor(normalizedDecimalPortion / (tickSize * Math.Pow(10, numberOfDigits))) * (tickSize * Math.Pow(10, numberOfDigits))
            Dim finalNumber As Double = Math.Truncate(number) + normalizedConvertedDecimalPortion / Math.Pow(10, numberOfDigits)

            Select Case upDown
                Case RoundOfType.Celing
                    Return (If(finalNumber < number, finalNumber + tickSize, finalNumber))
                Case RoundOfType.Floor
                    Return (If(finalNumber < tickSize, tickSize, finalNumber))
                Case Else
                    Throw New ApplicationException("Not Implemented")
            End Select
        End Function
        Public Function RoundEX(ByVal number As Decimal, ByVal places As Integer) As Decimal
            Return CType(Math.Round(number, places, MidpointRounding.AwayFromZero), Decimal)
        End Function
        Public Function RoundEX(ByVal number As Double, ByVal places As Integer) As Double
            Return CType(Math.Round(number, places, MidpointRounding.AwayFromZero), Double)
        End Function
#End Region
    End Module
End Namespace
