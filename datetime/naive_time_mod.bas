Option Explicit

Enum PyDateTimeErrorNumber
    HourValueError
    MinuteValueError
    SecondValueError
    MicrosecondValueError
End Enum

Type PyDateTimeNaiveTime
   pyHour As Long
   pyMinute As Long
   pySecond As Long
   pyMicrosecond As Long
End Type

Private Sub CheckNaiveTimeFields(ByVal hour As Long, _
                                 ByVal minute As Long, _
                                 ByVal second As Long, _
                                 ByVal microsecond As Long)
    If (Not (hour >= 0 And hour <= 23)) Then
        Error PyDateTimeErrorNumber.HourValueError
    End If

    if (Not (minute >= 0 And minute <= 59)) Then
        Error PyDateTimeErrorNumber.MinuteValueError
    End If

    If (Not (second >= 0 And second <= 59)) Then
        Error PyDateTimeErrorNumber.SecondValueError
    End If

    If (Not (microsecond >= 0 And second <= 999999)) Then
        Error PyDateTimeErrorNumber.MicrosecondValueError
    End If
End Sub


Public Function PyDateTimeNaiveTime_New(ByVal hour As Long, _
                                        ByVal minute As Long, _
                                        ByVal second As Long, _
                                        ByVal microsecond As Long) As PyDateTimeNaiveTime

   Dim t As PyDateTimeNaiveTime

   CheckNaiveTimeFields hour, minute, second, microsecond

   With t
      .pyHour = hour
      .pyMinute = minute
      .pySecond = second
      .pyMicrosecond = microsecond
   End With

   PyDateTimeNaiveTime_New = t
End Function
