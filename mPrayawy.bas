Attribute VB_Name = "mPrayawy"
Option Explicit

Public Const pi As Double = 3.14159265358979
Public Const DtoR As Double = (pi / 180#)
Public Const RtoD As Double = (180# / pi)

'GetRatior(int yg,int mg,int dg,double param[],double *IshRt,double *FajrRt);
'  Function to obtain the ratio of the start time of Isha and Fajr at
'  a referenced latitude (45deg suggested by Rabita) to the night length
Public Declare Sub GetRatior Lib "Prayawy.dll" (ByVal yg As Long, ByVal mg As Long, ByVal dg As Long, ByRef Param As Double, ByRef IshRt As Double, ByRef FajrRt As Double)

'double WINAPI atanxy(double x,double y);
Public Declare Function atanxy Lib "Prayawy.dll" (ByVal X As Double, ByVal y As Double) As Double

'void WINAPI EclipToEquator(double lmdr,double betar,double &alph, double &dltr);
Public Declare Sub EclipToEquator Lib "Prayawy.dll" (ByVal lmdr As Double, ByVal betar As Double, ByRef alph As Double, ByRef dltr As Double)

'double WINAPI RoutinR2(double M,double e)
'    Calculate the value of E
'    p.91, Peter Duffett-Smith book
Public Declare Function RoutinR2 Lib "Prayawy.dll" (ByVal M As Double, ByVal e As Double) As Double

'double WINAPI GCalendarToJD(int yg,int mg, double dg );
'/****************************************************************************/
'/* Name:    GCalendarToJD                           */
'/* Type:    Function                                                        */
'/* Purpose: convert Gdate(year,month,day) to Julian Day                 */
'/* Arguments:                                                               */
'/* Input : Gregorian date: year:yy, month:mm, day:dd                        */
'/* Output:  The Julian Day: JD                                              */
'/****************************************************************************/
Public Declare Function GCalendarToJD Lib "Prayawy.dll" (ByVal yg As Long, ByVal mg As Long, ByVal dg As Double) As Double

'double WINAPI SunParamr(int yg,int mg, int dg,double ObsLon,double ObsLat, double TimeZone,
'        double *Rise,double *Transit,double *Setting,double *RA,double *Decl,int *RiseSetFlags)
' p.99 of the Peter Duffett-Smith book
Public Declare Function SunParamr Lib "Prayawy.dll" (ByVal yg As Long, ByVal mg As Long, _
    ByVal dg As Long, ByVal ObsLon As Double, ByVal ObsLat As Double, ByVal TimeZone As Double, _
    ByRef Rise As Double, ByRef Transit As Double, ByRef Setting As Double, _
    ByRef RA As Double, ByRef Decl As Double, ByRef RiseSetFlags As Long) As Double

'int WINAPI OmAlQrahr(int yg,int mg,int dg, double param[], double lst[]);
'/*
'  For international prayer times see Islamic Fiqah Council of the Muslim
'  World League:  Saturday 12 Rajeb 1406H, concerning prayer times and fasting
'  times for countries of high latitudes.
'  This program is based on the above.
'*/
'/*****************************************************************************/
'/* Name:    OmAlQrah                                                         */
'/* Type:    Procedure                                                        */
'/* Purpose: Compute prayer times and sunrise                                 */
'/* Arguments:                                                                */
'/*   yg,mg,dg : Date in Greg                                                 */
'/*   param[0]: Safety time  in hours should be 0.016383h                     */
'/*   longtud,latud: param[1],[2] : The place longtude and latitude in radians*/
'/*   HeightdifW : param[3]: The place western herizon height difference in meters */
'/*   HeightdifE : param[4]: The place eastern herizon height difference in meters */
'/*   Zonh :param[5]: The place zone time dif. from GMT  West neg and East pos*/
'/*          in decimal hours                                                 */
'/*  fjrangl: param[6]: The angle (radian) used to compute                    */
'/*            Fajer prayer time (OmAlqrah  -19 deg.)                         */
'/*  ashangl: param[7]: The angle (radian) used to compute Isha  prayer time  */
'/*          ashangl=0 then use  (OmAlqrah: ash=SunSet+1.5h)                  */
'/*  asr  : param[8]: The Henfy (asr=2) Shafi (asr=1, Omalqrah asr=1)         */
'/*  param[9]: latude (radian) that should be used for places above -+65.5    */
'/*            should be 45deg as suggested by Rabita                         */
'/*   param[10]: The Isha fixed time from Sunset                              */
'/*  Output:                                                                  */
'/*  lst[]: lst[n], 1:Fajer 2:Sunrise 3:Zohar 4:Aser  5:Magreb  6:Ishe        */
'/*                 7:Fajer using exact Rabita method for places >48          */
'/*                 8:Ash   using exact Rabita method for places >48          */
'/*                 9: Eid Prayer Time                                        */
'/*          for places above 48 lst[1] and lst[6] use a modified version of  */
'/*          Rabita method that tries to eliminate the discontinuity          */
'/*         all in 24 decimal hours                                           */
'/*         returns flag:0 if there are problems, flag:1 no problems          */
'/*****************************************************************************/
Public Declare Function OmAlQrahr Lib "Prayawy.dll" (ByVal yg As Long, _
        ByVal mg As Long, ByVal dg As Long, ByRef Param As Double, _
        ByRef lst As Double) As Long

'/****************************************************************************/
'/* Name:    BH2GA                                                            */
'/* Type:    Procedure                                                       */
'/* Purpose: Finds Gdate(year,month,day) for Hdate(year,month,day=1)         */
'/* Arguments:                                                               */
'/* Input: Hijrah  date: year:yh, month:mh                                   */
'/* Output: Gregorian date: year:yg, month:mg, day:dg , day of week:dayweek  */
'/*       and returns flag found:1 not found:0                               */
'/****************************************************************************/
Public Declare Function BH2GA Lib "Prayawy.dll" (yh As Long, mh As Long, yg As Long, mg As Long, dg As Long, DayWeek As Long) As Long

' int  WINAPI G2HA(int yg,int mg, int dg,int *yh,int *mh,int *dh,int *dayweek)
'/****************************************************************************/
'/* Name:    G2HA                                                            */
'/* Type:    Procedure                                                       */
'/* Purpose: convert Gdate(year,month,day) to Hdate(year,month,day)          */
'/* Arguments:                                                               */
'/* Input: Gregorian date: year:yg, month:mg, day:dg                         */
'/* Output: Hijrah  date: year:yh, month:mh, day:dh, day of week:dayweek     */
'/*       and returns flag found:1 not found:0                               */
'/****************************************************************************/
Public Declare Function G2HA Lib "Prayawy.dll" (yg As Long, mg As Long, dg As Long, yh As Long, mh As Long, dh As Long, DayWeek As Long) As Long

'int  WINAPI H2GA(int *yh,int *mh,int *dh, int *yg,int *mg, int *dg,int *dayweek)
'/****************************************************************************/
'/* Name:    H2GA                                                            */
'/* Type:    Procedure                                                       */
'/* Purpose: convert Hdate(year,month,day) to Gdate(year,month,day)          */
'/* Arguments:                                                               */
'/* Input/Ouput: Hijrah  date: year:yh, month:mh, day:dh                     */
'/* Output: Gregorian date: year:yg, month:mg, day:dg , day of week:dayweek  */
'/*       and returns flag found:1 not found:0                               */
'/* Note: The function will correct Hdate if day=30 and the month is 29 only */
'/****************************************************************************/
Public Declare Function H2GA Lib "Prayawy.dll" (yh As Long, mh As Long, dh As Long, yg As Long, mg As Long, dg As Long, DayWeek As Long) As Long

'double WINAPI JDToGCalendar(double JD, int *yy,int *mm, int *dd);
'/****************************************************************************/
'/* Name:    JDToGCalendar                           */
'/* Type:    Procedure                                                       */
'/* Purpose: convert Julian Day  to Gdate(year,month,day)                    */
'/* Arguments:                                                               */
'/* Input:  The Julian Day: JD                                               */
'/* Output: Gregorian date: year:yy, month:mm, day:dd                        */
'/****************************************************************************/
Public Declare Function JDToGCalendar Lib "Prayawy.dll" (ByVal JD As Double, _
        ByRef yy As Long, ByRef mm As Long, ByRef dd As Long) As Double
        
'int WINAPI GLeapYear(int year);
'/****************************************************************************/
'/* Name:    GLeapYear                                   */
'/* Type:    Function                                                        */
'/* Purpose: Determines if  Gdate(year) is leap or not                   */
'/* Arguments:                                                               */
'/* Input : Gregorian date: year                                 */
'/* Output:  0:year not leap   1:year is leap                                */
'/****************************************************************************/
Public Declare Function GLeapYear Lib "Prayawy.dll" (ByVal year As Long) As Long

'void WINAPI GDateAjust(int *yg,int *mg,int *dg);
'/****************************************************************************/
'/* Name:    GDateAjust                              */
'/* Type:    Procedure                                                       */
'/* Purpose: Adjust the G Dates by making sure that the month lengths        */
'/*      are correct if not so take the extra days to next month or year */
'/* Arguments:                                                               */
'/* Input: Gregorian date: year:yg, month:mg, day:dg                         */
'/* Output: corrected Gregorian date: year:yg, month:mg, day:dg              */
'/****************************************************************************/
Public Declare Sub GDateAjust Lib "Prayawy.dll" (ByRef yg As Long, ByRef mg As Long, ByRef dg As Long)

'int WINAPI DayWeek(long JulianD);
'/*
'  The day of the week is obtained as
'  Dy=(Julian+1)%7
'  Dy=0 Sunday
'  Dy=1 Monday
'  ...
'  Dy=6 Saturday
'*/
Public Declare Function DayWeek Lib "Prayawy.dll" (ByVal JulianD As Long) As Long

'int WINAPI DayinYear(int yh,int mh,int dh)
'/****************************************************************************/
'/* Name:    DayInYear                               */
'/* Type:    Function                                                        */
'/* Purpose: Obtains the day number in the yea                           */
'/* Arguments:                                                               */
'/* Input : Hijrah  date: year:yh, month:mh  day:dh                          */
'/* Output:  Day number in the Year                      */
'/****************************************************************************/
Public Declare Function DayinYear Lib "Prayawy.dll" (ByVal yh As Long, ByVal mh As Long, ByVal dh As Long) As Long

'int WINAPI HYearLength(int yh)
'/****************************************************************************/
'/* Name:    HYearLength                                 */
'/* Type:    Function                                                        */
'/* Purpose: Obtains the year length                                     */
'/* Arguments:                                                               */
'/* Input : Hijrah  date: year:yh                                        */
'/* Output:  Year Length                                                     */
'/****************************************************************************/
Public Declare Function HYearLength Lib "Prayawy.dll" (ByVal yh As Long) As Long

'void WINAPI JDToHCalendar(double JD,int *yh,int *mh,int *dh);
'/****************************************************************************/
'/* Name:    JDToHCalendar                           */
'/* Type:    Procedure                                                       */
'/* Purpose: convert Julian Day to estimated Hdate(year,month,day)       */
'/* Arguments:                                                               */
'/* Input:  The Julian Day: JD                                               */
'/* Output : Hijrah date: year:yh, month:mh, day:dh                          */
'/****************************************************************************/
Public Declare Sub JDToHCalendar Lib "Prayawy.dll" (ByVal JD As Double, _
                ByRef yh As Long, ByRef mh As Long, ByRef dh As Long)

'void WINAPI JDToHACalendar(double JD,int *yh,int *mh,int *dh);
'/****************************************************************************/
'/* Name:    JDToHACalendar                          */
'/* Type:    Procedure                                                       */
'/* Purpose: convert Julian Day to  Hdate(year,month,day)                 */
'/* Arguments:                                                               */
'/* Input:  The Julian Day: JD                                               */
'/* Output : Hijrah date: year:yh, month:mh, day:dh                          */
'/****************************************************************************/
Public Declare Sub JDToHACalendar Lib "Prayawy.dll" (ByVal JD As Double, _
        ByRef yh As Long, ByRef mh As Long, ByRef dh As Long)

'double WINAPI HCalendarToJD(int yh,int mh,int dh);
'/****************************************************************************/
'/* Name:    HCalendarToJDA                          */
'/* Type:    Function                                                        */
'/* Purpose: convert Hdate(year,month,day) to Exact Julian Day           */
'/* Arguments:                                                               */
'/* Input : Hijrah  date: year:yh, month:mh, day:dh                          */
'/* Output:  The Exact Julian Day: JD                                        */
'/****************************************************************************/
Public Declare Function HCalendarToJD Lib "Prayawy.dll" (ByVal yh As Long, _
    ByVal mh As Long, ByVal dh As Long) As Double
    
' double WINAPI HCalendarToJDA(int yh,int mh,int dh);
'/****************************************************************************/
'/* Name:    HCalendarToJDA                          */
'/* Type:    Function                                                        */
'/* Purpose: convert Hdate(year,month,day) to Exact Julian Day           */
'/* Arguments:                                                               */
'/* Input : Hijrah  date: year:yh, month:mh, day:dh                          */
'/* Output:  The Exact Julian Day: JD                                        */
'/****************************************************************************/
Public Declare Function HCalendarToJDA Lib "Prayawy.dll" (ByVal yh As Long, _
    ByVal mh As Long, ByVal dh As Long) As Double

'int WINAPI IsValid(int yh, int mh, int dh)
'/* Purpose: returns 0 for incorrect Hijri date and 1 for correct date      */
Public Declare Function IsValid Lib "Prayawy.dll" (ByVal yh As Long, _
    ByVal mh As Long, ByVal dh As Long) As Long

'int WINAPI HMonthLength(int yh,int mh);
'/****************************************************************************/
'/* Name:    HMonthLength                            */
'/* Type:    Function                                                        */
'/* Purpose: Obtains the month length                                */
'/* Arguments:                                                               */
'/* Input : Hijrah  date: year:yh, month:mh                                  */
'/* Output:  Month Length                                                    */
'/****************************************************************************/
Public Declare Function HMonthLength Lib "Prayawy.dll" (ByVal yh As Long, ByVal mh As Long) As Long

'--------------------------------------------------------------------
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long


Public Function PrayTimes(ByVal yg As Long, _
        ByVal mg As Long, ByVal dg As Long, ByRef Param() As Double, _
        ByRef lst() As Double) As Long
    
    PrayTimes = OmAlQrahr(yg, mg, dg, Param(0), lst(0))
    
End Function

