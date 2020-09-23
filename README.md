<div align="center">

## Calculate InternetTime


</div>

### Description

The function InternetTime() calculates the internettime, the new time standard from Swatch. You only have to call the function.
 
### More Info
 
Copy and Paste all of these code in one single module.

This function returns a value containing the internettime. If you want to convert it to a small string, you could use the following: Mid(Format(MyTime, "000.0000000"), 1, 3)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hiu\-Hong Hau](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hiu-hong-hau.md)
**Level**          |Unknown
**User Rating**    |3.0 (6 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hiu-hong-hau-calculate-internettime__1-2131/archive/master.zip)

### API Declarations

```
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*****************************************************************'
'* The function InternetTime() returns the current time     *'
'* in beats. The code to determine the timezone has been written *'
'* by Dror Saddan (drors@ietusa.com).              *'
'* The code to calculate the internettime has been written by  *'
'* Swatch. I have ported it to Visual Basic.           *'
'*                                *'
'* Written by Hiu-Hong Hau (hhhau@dds.nl)            *'
'* Date: June 20th 1999                     *'
'* Website: http://www.supervisie.nl/qlaunch           *'
'* Website: http://www.supervisie.nl/gemini           *'
'*                                *'
'* Take a look at QuickLaunch, a skinnable application launcher, *'
'* completely written in Visual Basic.              *'
'*****************************************************************'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type SYSTEMTIME ' 16 Bytes
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
Public Type TIME_ZONE_INFORMATION
  Bias As Long
  StandardName(31) As Integer
  StandardDate As SYSTEMTIME
  StandardBias As Long
  DaylightName(31) As Integer
  DaylightDate As SYSTEMTIME
  DaylightBias As Long
End Type
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation _
  As TIME_ZONE_INFORMATION) As Long
```


### Source Code

```

Public Function GetTimeZone(Optional ByRef strTZName As String) As Long
  Dim objTimeZone As TIME_ZONE_INFORMATION
  Dim lngResult As Long
  Dim i As Long
  lngResult = GetTimeZoneInformation&(objTimeZone)
  Select Case lngResult
   Case 0&, 1& 'use standard time
   GetTimeZone = -(objTimeZone.Bias + objTimeZone.StandardBias) 'into minutes
   For i = 0 To 31
     If objTimeZone.StandardName(i) = 0 Then Exit For
     strTZName = strTZName & Chr(objTimeZone.StandardName(i))
   Next
   Case 2& 'use daylight savings time
   GetTimeZone = -(objTimeZone.Bias + objTimeZone.DaylightBias) 'into minutes
   For i = 0 To 31
     If objTimeZone.DaylightName(i) = 0 Then Exit For
     strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
   Next
  End Select
End Function
Public Function InternetTime()
  Dim tmpH
  Dim tmpS
  Dim tmpM
  Dim itime
  Dim tmpZ
  Dim testtemp As String
  tmpH = Hour(Time)
  tmpM = Minute(Time)
  tmpS = Second(Time)
  tmpZ = GetTimeZone
  itime = ((tmpH * 3600 + ((tmpM - tmpZ + 60) * 60) + tmpS) * 1000 / 86400)
  If itime > 1000 Then
   itime = itime - 1000
  ElseIf itime < 0 Then
   itime = itime + 1000
  End If
  InternetTime = itime
End Function
```

