Attribute VB_Name = "mDaylightTime"
'-----------------------------------------------------------------------------------------
' Copyright ©1996-2005 VBnet, Randy Birch. All Rights Reserved Worldwide.
'        Terms of use http://vbnet.mvps.org/terms/pages/terms.htm
'-----------------------------------------------------------------------------------------
'http://vbnet.mvps.org/index.html?code/locale/timezoneforecast.htm

Private Const TIME_ZONE_ID_UNKNOWN As Long = 1
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF

Private Const LB_SETTABSTOPS As Long = &H192
Private Const LB_FINDSTRING = &H18F

Public Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte  'unicode (0-based)
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte  'unicode (0-based)
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Public Enum DateFormats
   vbGeneralDate = 0
   vbLongDate = 1
   vbShortDate = 2
   vbLongTime = 3
   vbShortTime = 4
End Enum

Public Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Sub Command1_Click()

   Dim tzi As TIME_ZONE_INFORMATION
   Dim tziYear As Long
   
  'retrieve time zone info for the system
   Call GetTimeZoneInformation(tzi)
   
  'The .wYear parameter in the StandardTime
  'and DaylightTime members return as 0,
  'and we have to pass a valid year as well.
  'This shows historical dates since DST
  'formally began in 1966, continuing
  'for 100 years.
   For tziYear = 1966 To 2066
   
     'call method that uses the time zone info
     'returned to calculate the actual dates that
     'daylight/standard time changes
      List1.AddItem tziYear & vbTab & _
                    GetTimeZoneChangeDate(tzi.DaylightDate, _
                                          tziYear, _
                                          vbLongDate) & vbTab & _
                    GetTimeZoneChangeDate(tzi.StandardDate, _
                                          tziYear, _
                                          vbLongDate)
   Next
   
  'this just locates the current year
  'in the list, making it the top list item
   EnsureItemVisible CStr(Year(Now))
   
End Sub

Public Function GetTimeZoneChangeDate(tziDate As SYSTEMTIME, _
                                       ByVal tziYear As Long) As Long

  'thanks to Mathias Schiffer for this routine
  
   Dim tmp As Date
   Dim MonthFirstWeekday As Long

   With tziDate

      Select Case .wDay 'week in month

         Case 1 To 4:   'week 1 to week 4

           'Calculate the first day in the month,
           'and then calculate the appropriate day
           'that the time zone change will occur
            MonthFirstWeekday = Weekday(DateSerial(tziYear, .wMonth, 1)) - 1
            tmp = DateSerial(tziYear, _
                             .wMonth, _
                              (.wDayOfWeek - MonthFirstWeekday + _
                              .wDay * 7) Mod 7 + 1)



         Case 5:        'last week in month

           'Calculate the month's last day,
           'then work back to the appropriate
           'weekday
            tmp = DateSerial(tziYear, .wMonth + 1, 0)
            tmp = DateAdd("d", tmp, _
                          -(Weekday(tmp) - .wDayOfWeek + 7 - 1) Mod 7)

      End Select

   End With
   
  'Now that the date has been calculated,
  'return it in the string format requested
  'In VB6, you can use the FormatDateTime
  'function to return date in specified format
   'FormatDateTime(tmp, dwType)
   
   GetTimeZoneChangeDate = tmp
   
End Function

Public Function GetTimeZoneChangeTime(tzi As SYSTEMTIME, _
                                       ByVal dwType As DateFormats) As Long
                                       
   Dim tmp As Date
   
   tmp = TimeSerial(tzi.wHour, tzi.wMinute, tzi.wSecond)
   'GetTimeZoneChangeTime = FormatDateTime(tmp, dwType)
   GetTimeZoneChangeTime = tmp
   
End Function

Public Function GetDaylightStartDate(ByVal Year As Long) As Long

    Dim tzi As TIME_ZONE_INFORMATION
    Dim tziYear As Long
    
    Call GetTimeZoneInformation(tzi)
    GetDaylightStartDate = GetTimeZoneChangeDate(tzi.DaylightDate, Year)

End Function

Public Function GetDaylightEndDate(ByVal Year As Long) As Long

    Dim tzi As TIME_ZONE_INFORMATION
    Dim tziYear As Long
    
    Call GetTimeZoneInformation(tzi)
    GetDaylightEndDate = GetTimeZoneChangeDate(tzi.StandardDate, Year)

End Function

