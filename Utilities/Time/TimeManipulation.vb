Imports NLog
Namespace Time
    Public Module TimeManipulation
#Region "Logging and Status Progress"
        Public logger As Logger = LogManager.GetCurrentClassLogger
#End Region

#Region "Private Attributes"
        Private INDIAN_ZONE As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time")
#End Region

#Region "Public Methods"
        Public Function IsDateTimeEqualTillMinutes(ByVal datetime1 As Date, ByVal datetime2 As Date) As Boolean
            Return datetime1.Date = datetime2.Date And
                    datetime1.Hour = datetime2.Hour And
                    datetime1.Minute = datetime2.Minute
        End Function
        Public Function GetCurrentISTTime() As DateTime
            Return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE)
        End Function
        Public Function UnixToDateTime(ByVal unixTimeStamp As UInt64) As DateTime
            Dim dateTime As DateTime = New DateTime(1970, 1, 1, 5, 30, 0, 0, DateTimeKind.Unspecified)
            dateTime = dateTime.AddSeconds(unixTimeStamp)
            Return dateTime
        End Function
        Public Function GetDateTimeTillMinutes(ByVal datetime1 As Date) As Date
            'logger.Debug("Converting datetime till minutes")
            Return New Date(datetime1.Date.Year, datetime1.Date.Month, datetime1.Date.Day,
                            datetime1.Hour, datetime1.Minute, 0)
        End Function
#End Region
    End Module
End Namespace