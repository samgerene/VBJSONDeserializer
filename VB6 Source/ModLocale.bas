Attribute VB_Name = "ModLocale"
Option Explicit

Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ICOUNTRY = &H5
Public Const LOCALE_ICURRDIGITS = &H19
Public Const LOCALE_ICURRENCY = &H1B
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_IDAYLZERO = &H26
Public Const LOCALE_IDEFAULTCODEPAGE = &HB
Public Const LOCALE_IDEFAULTCOUNTRY = &HA
Public Const LOCALE_IDEFAULTLANGUAGE = &H9
Public Const LOCALE_IDIGITS = &H11
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ILANGUAGE = &H1
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_IMEASURE = &HD
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_INEGCURR = &H1C
Public Const LOCALE_INEGSEPBYSPACE = &H57
Public Const LOCALE_INEGSIGNPOSN = &H53
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ITLZERO = &H25
Public Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Const LOCALE_S1159 = &H28
Public Const LOCALE_S2359 = &H29
Public Const LOCALE_SABBREVCTRYNAME = &H7
Public Const LOCALE_SABBREVDAYNAME1 = &H31
Public Const LOCALE_SABBREVDAYNAME2 = &H32
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SABBREVLANGNAME = &H3
Public Const LOCALE_SABBREVMONTHNAME1 = &H44
Public Const LOCALE_SCOUNTRY = &H6
Public Const LOCALE_SCURRENCY = &H14
Public Const LOCALE_SDATE = &H1D
Public Const LOCALE_SDAYNAME1 = &H2A
Public Const LOCALE_SDAYNAME2 = &H2B
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SGROUPING = &H10
Public Const LOCALE_SINTLSYMBOL = &H15
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SLIST = &HC
Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONGROUPING = &H18
Public Const LOCALE_SMONTHNAME1 = &H38
Public Const LOCALE_SMONTHNAME10 = &H41
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHOUSANDSEP = &H17
Public Const LOCALE_SNATIVECTRYNAME = &H8
Public Const LOCALE_SNATIVEDIGITS = &H13
Public Const LOCALE_SNATIVELANGNAME = &H4
Public Const LOCALE_SNEGATIVESIGN = &H51
Public Const LOCALE_SPOSITIVESIGN = &H50
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_STIME = &H1E
Public Const LOCALE_STIMEFORMAT = &H1003

Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Function GetRegionalSettings(ByVal regionalsetting As Long) As String ' Retrieve the regional setting

On Error GoTo errorHandler

Dim Locale As Long
Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim Pos As Integer
      
Locale = GetUserDefaultLCID()

iRet1 = GetLocaleInfo(Locale, regionalsetting, lpLCDataVar, 0)
Symbol = String$(iRet1, 0)
iRet2 = GetLocaleInfo(Locale, regionalsetting, Symbol, iRet1)
Pos = InStr(Symbol, Chr$(0))
If Pos > 0 Then
    Symbol = Left$(Symbol, Pos - 1)
End If
      
errorHandler:
GetRegionalSettings = Symbol
Select Case Err.Number
Case 0
Case Else
    Err.Raise 123, "GetRegionalSetting", "GetRegionalSetting: " & regionalsetting
End Select
End Function




