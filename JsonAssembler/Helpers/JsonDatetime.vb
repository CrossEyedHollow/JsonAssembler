Module JsonDatetime
    Property TimeFormat As String = "yyMMddHH"
    Property LongTimeFormat As String = "yyyy-MM-ddThh:mm:ssZ"
    Property DateFormat As String = "yyyy-MM-dd"

    Public Function GetTime() As String
        Return Date.UtcNow.ToString("yyMMddHH")
    End Function
    Public Function GetTime(time As Date) As String
        Return time.ToString("yyMMddHH")
    End Function

    Public Function GetTimeLong() As String
        Return Date.UtcNow.ToString("yyyy-MM-ddThh:mm:ssZ")
    End Function
    Public Function GetTimeLong(t As Date) As String
        Return t.ToString("yyyy-MM-ddThh:mm:ssZ")
    End Function

    Public Function ParseTime(time As String) As Date
        Return Date.ParseExact(time, TimeFormat, Nothing)
    End Function
    Public Function ParseTimeLong(time As String) As Date
        Return Date.ParseExact(time, LongTimeFormat, Nothing)
    End Function
    Public Function ParseDate(time As String) As Date
        Return Date.ParseExact(time, DateFormat, Nothing)
    End Function
End Module
