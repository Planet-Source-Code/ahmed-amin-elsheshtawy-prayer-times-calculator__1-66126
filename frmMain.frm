VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=Copyright Infomation
'==========================================================
'Program Name     : PrayerTimes
'Program Author   : Ahmed Amin Elsheshtawy
'Home Page        : http://www.islamware.com
'Contact Email    : support@islamware.com
'Copyrights © 2006 IslamWare. All rights reserved.
'==========================================================
'This is a module that computes prayer times and sunrise.
'Require PrayerTimes.dll library from www.islamware.com free
'==========================================================
Option Explicit

Private Sub Form_Load()

On Error Resume Next

Dim longtud As Double, latud As Double, Zonh As Double
Dim i As Integer, X As Integer
Dim Param(0 To 11) As Double
Dim lst(0 To 15) As Double
Dim yg As Long, mg As Long, dg As Double
Dim yg1 As Long, mg1 As Long, dg1 As Long
Dim Hour As Integer, Min As Integer, Sec As Integer
    
    longtud = 31.25
    latud = 30.05
    Zonh = 2
    
    Param(0) = 0.016388 ' /* 59 seconds, safety time */
    Param(1) = longtud * DtoR '  /* Longitude in radians */
    Param(2) = latud * DtoR '
    Param(3) = 23#
    Param(4) = 23#
    Param(5) = Zonh '     /* Time Zone difference from GMT S.A. +3*/
    Param(6) = 19 * DtoR '; /* Fajer Angle */
    Param(7) = 0 '; /* Isha Angle  */
    Param(8) = 1 ';    /* Aser=1,2 /
    Param(9) = 45 * DtoR ';   /* Reference Angle suggested by Rabita */
    Param(10) = 1.5 ';  /* Isha fixed time from sunset */
    Param(11) = 4.2 * DtoR '; /* Eid Prayer Time   */
    
    'i = PrayTimes(yg1, mg1, dg1, Param(), lst())
    i = PrayerTimes(yg, mg, dg, Param(0), lst(0))
    'lst(x):
    '1:   Fajer
    '2:   Sunrise
    '3:   Zohar
    '4:   Aser
    '5:   Magreb
    '6:   Isha
    '7:   Fajer using exact Rabita method for places > 48
    '8:   Isha using exact Rabita method for places > 48
    '9:   Eid Prayer Time
    
    FormatTime lst(1), Hour, Min, Sec
    Debug.Print "Fajer: " & Hour & ":" & Min & ":" & Sec
    
    FormatTime lst(3), Hour, Min, Sec
    Debug.Print "Zohr: " & Hour & ":" & Min & ":" & Sec
    
    FormatTime lst(4), Hour, Min, Sec
    Debug.Print "Aser: " & Hour & ":" & Min & ":" & Sec
    
    FormatTime lst(5), Hour, Min, Sec
    Debug.Print "Magreb: " & Hour & ":" & Min & ":" & Sec
    
    FormatTime lst(6), Hour, Min, Sec
    Debug.Print "Isha: " & Hour & ":" & Min & ":" & Sec
    
    FormatTime lst(2), Hour, Min, Sec
    Debug.Print "Subrise: " & Hour & ":" & Min & ":" & Sec
    
    'For X = 1 To 9
        'FormatTime lst(X), Hour, Min, Sec
        'Debug.Print "Time " & X & "= " & Hour & ":" & Min & ":" & Sec, lst(X)
    'Next
    '--------------------------------------------------------------------
    ' Daylight/Standard time
    'Debug.Print "Daylight Time Starts: " & FormatDateTime(GetDaylightStartDate(Year(Now)))
    'Debug.Print "Standard Time Starts: " & FormatDateTime(GetDaylightEndDate(Year(Now)))
    '--------------------------------------------------------------------
    'FreeLibrary (PrayerTimes)
End Sub

Private Sub FormatTime(ByVal Hours As Double, Hour As Integer, Min As Integer, Sec As Integer)

    Dim Minutes As Double
    
    Hour = Int(Hours)
    
    Minutes = 60 * Abs(Hours - Hour)
    
    Min = Fix(Minutes)
    
    If (Min = 60) Then
       Hour = Hour + 1
       Min = 0
    End If
    
    Sec = Fix(60 * (Minutes - Min))
    
    If (Sec = 60) Then
        Min = Min + 1
        Sec = 0
        If (Min = 60) Then
           Hour = Hour + 1
           Min = 0
        End If
    End If

End Sub
