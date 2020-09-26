Attribute VB_Name = "Module1"


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Const SRCCOPY = &HCC0020

Global iTPPY As Long
Global iTPPX As Long



  
Public nep_year As Variant
Public nep_month As Variant
Public nep_date As Variant
Public nep_day As Variant



Public eng_year As Variant
Public eng_month As Variant
Public eng_date As Variant
Public eng_day As Variant
Public bs(90) As Variant


Public listofevents(20) As Variant


Public debug_info
  Public Sub checkevents()
 
  
  
'  listofevents(0) = Array("New year in BS", "bs", 1, 1)
'    listofevents(1) = Array("Loktantra diwas", "bs", 1, 11)
'  listofevents(2) = Array("World Labour Day", "ad", 5, 1)
'  listofevents(3) = Array("Ganatantra diwas", "bs", 2, 15)
'  listofevents(4) = Array("Guru nanak day", "ad", 11, 11)
'  listofevents(5) = Array("World AIDS day", "ad", 12, 1)
'  listofevents(6) = Array("World handicapped day", "ad", 12, 3)
'  listofevents(7) = Array("Christmas day", "ad", 12, 25)
'  listofevents(8) = Array("Maghi parva", "bs", 10, 1)
'  listofevents(9) = Array("Shahid diwas", "bs", 10, 16)
'  listofevents(10) = Array("Democracy day", "bs", 11, 7)
'  listofevents(11) = Array("World women day", "ad", 3, 8)
'  listofevents(12) = Array("New year in AD", "ad", 1, 1)
'  listofevents(13) = Array("Valentine day", "ad", 2, 14)
'    listofevents(14) = Array("L. Birendra Birthday", "bs", 9, 14)
'    listofevents(15) = Array("Tamu Lhosar", "bs", 9, 15)
'
'  listofevents(16) = Array("END", "end", 0, 0)
  i = 0
 Open App.Path & "\tools\events.set" For Input As #1
  While Not EOF(1)
   Input #1, a
   Input #1, b
   Input #1, c
   Input #1, d

   listofevents(i) = Array(a, b, c, d)

i = i + 1

Wend
Close #1
  listofevents(i) = Array("END", "end", 0, 0)

  End Sub
  
  
  
Public Sub showdate()
Call eng_to_nep(Right(Date$, 4), Val(Left(Date$, 2)), Mid(Date$, 4, 2))
devcalender.eventbar.Visible = False

devcalender.nepaliyear = nep_year
devcalender.nepalimonth = get_nepali_month(nep_month)
devcalender.nepalidate = nep_date
devcalender.nepaliday = get_day_of_week(nep_day)


Call checkevents
devcalender.showtodaysevent = ""
 i = 0
Do
If listofevents(i)(1) = "bs" And listofevents(i)(2) = nep_month And listofevents(i)(3) = nep_date Then
devcalender.showtodaysevent.Caption = listofevents(i)(0)
devcalender.showtodaysevent.ToolTipText = listofevents(i)(0)

End If
If listofevents(i)(1) = "ad" And listofevents(i)(2) = Val(Left(Date$, 2)) And listofevents(i)(3) = Val(Mid(Date$, 4, 2)) Then
devcalender.showtodaysevent.Caption = listofevents(i)(0)
devcalender.showtodaysevent.ToolTipText = listofevents(i)(0)

End If




i = i + 1


Loop While Not listofevents(i)(0) = "END"

If Len(devcalender.showtodaysevent.Caption) <> 0 Then devcalender.eventbar.Visible = 1
End Sub

Public Sub initilizeClass()


bs(0) = Array(2000, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(1) = Array(2001, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(2) = Array(2002, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(3) = Array(2003, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(4) = Array(2004, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(5) = Array(2005, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(6) = Array(2006, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(7) = Array(2007, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(8) = Array(2008, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)

bs(9) = Array(2009, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(10) = Array(2010, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(11) = Array(2011, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(12) = Array(2012, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)

bs(13) = Array(2013, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(14) = Array(2014, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(15) = Array(2015, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(16) = Array(2016, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)

bs(17) = Array(2017, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(18) = Array(2018, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(19) = Array(2019, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(20) = Array(2020, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)

bs(21) = Array(2021, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(22) = Array(2022, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)

bs(23) = Array(2023, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(24) = Array(2024, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)

bs(25) = Array(2025, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(26) = Array(2026, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(27) = Array(2027, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(28) = Array(2028, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(29) = Array(2029, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30)

bs(30) = Array(2030, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(31) = Array(2031, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(32) = Array(2032, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(33) = Array(2033, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(34) = Array(2034, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(35) = Array(2035, 30, 32, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)

bs(36) = Array(2036, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(37) = Array(2037, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(38) = Array(2038, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(39) = Array(2039, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)

bs(40) = Array(2040, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(41) = Array(2041, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(42) = Array(2042, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(43) = Array(2043, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)

bs(44) = Array(2044, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(45) = Array(2045, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(46) = Array(2046, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(47) = Array(2047, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)

bs(48) = Array(2048, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(49) = Array(2049, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)

bs(50) = Array(2050, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(51) = Array(2051, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)

bs(52) = Array(2052, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(53) = Array(2053, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)

bs(54) = Array(2054, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(55) = Array(2055, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(56) = Array(2056, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30)

bs(57) = Array(2057, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(58) = Array(2058, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(59) = Array(2059, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(60) = Array(2060, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(61) = Array(2061, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(62) = Array(2062, 30, 32, 31, 32, 31, 31, 29, 30, 29, 30, 29, 31)

bs(63) = Array(2063, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(64) = Array(2064, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(65) = Array(2065, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(66) = Array(2066, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)

bs(67) = Array(2067, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(68) = Array(2068, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(69) = Array(2069, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(70) = Array(2070, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)

bs(71) = Array(2071, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(72) = Array(2072, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)

bs(73) = Array(2073, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)

bs(74) = Array(2074, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)

bs(75) = Array(2075, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(76) = Array(2076, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)

bs(77) = Array(2077, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)

bs(78) = Array(2078, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)

bs(79) = Array(2079, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)

bs(80) = Array(2080, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)

bs(81) = Array(2081, 31, 31, 32, 32, 31, 30, 30, 30, 29, 30, 30, 30)

bs(82) = Array(2082, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)

bs(83) = Array(2083, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30)

bs(84) = Array(2084, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30)

bs(85) = Array(2085, 31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30)

bs(86) = Array(2086, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)

bs(87) = Array(2087, 31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30)

bs(88) = Array(2088, 30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30)

bs(89) = Array(2089, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)

bs(90) = Array(2090, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)

'Set nep_date = New Nepali_Calender
'nep_date.Add "year", 0
'
'nep_date.Add "month", 0
'
'nep_date.Add "nmonth", 0
'
'nep_date.Add "num_day", 0
'
'Set eng_date = New Dictionary
'
'eng_date.Add "year", 0
'
'eng_date.Add "month", 0
'
'eng_date.Add "date", 0
'
'eng_date.Add "day", 0
'
'eng_date.Add "emonth", 0
'
'eng_date.Add "num_day", 0

End Sub

'‘ /**
'
'‘ * Calculates wheather english year is leap year or not
'
'‘ *
'
'‘ * @param integer $year
'
'‘ * @return boolean
'
'‘ */’

Public Function is_leap_year(year) As Boolean

Dim a, returnVal

a = year

If (a Mod 100 = 0 And a Mod 400 = 0) Or a Mod 4 = 0 Then


returnVal = True

Else

returnVal = False

End If


is_leap_year = returnVal

End Function

Public Function get_nepali_month(m) As String

Dim n_month As String

n_month = False

Select Case m

Case 1:

n_month = "Baisakh"

Case 2:

n_month = "Jestha"

Case 3:

n_month = "Asadh"

Case 4:

n_month = "Shrawn"

Case 5:

n_month = "Bhadra"

Case 6:

n_month = "Ashwin"

Case 7:

n_month = "Kartik"

Case 8:

n_month = "Mangsir"

Case 9:

n_month = "Poush"

Case 10:

n_month = "Magh"

Case 11:

n_month = "Falgun"

Case 12:

n_month = "Chaitra"

End Select

get_nepali_month = n_month

End Function

Public Function get_english_month(m)

Dim eMonth

eMonth = False

Select Case m

Case 1

eMonth = "January"

Case 2:

eMonth = "February"

Case 3:

eMonth = "March"

Case 4:

eMonth = "April"

Case 5:

eMonth = "May"

Case 6:

eMonth = "June"

Case 7:

eMonth = "July"

Case 8:

eMonth = "August"

Case 9:

eMonth = "September"

Case 10:

eMonth = "October"

Case 11:

eMonth = "November"

Case 12:

eMonth = "December"

End Select

get_english_month = eMonth

End Function

Public Function get_day_of_week(d) As String
On Error Resume Next

Dim day

day = False

Select Case d

Case 1:

day = "Sunday"

Case 2:

day = "Monday"

Case 3:

day = "Tuesday"

Case 4:

day = "Wednesday"

Case 5:

day = "Thursday"

Case 6:

day = "Friday"

Case 7:

day = "Saturday"

End Select

get_day_of_week = day

End Function

Public Function is_range_eng(yy, mm, dd)

Dim returnVal

returnVal = True

If (yy < 1944 Or yy > 2033) Then

debug_info = "Supported only between 1944-2022?"

returnVal = False

End If

If (mm < 1 Or mm > 12) Then

debug_info = "Error! value 1-12 only"

returnVal = False

End If

If (dd < 1 Or dd > 31) Then

debug_info = "Error! value 1-31 only"

returnVal = False

End If

is_range_eng = returnVal

End Function

Public Function is_range_nep(yy, mm, dd)

Dim returnVal

returnVal = True

If (yy < 2000 Or yy > 2089) Then

debug_info = "Supported only between 2000-2089?"

returnVal = False

End If

If (mm < 1 Or mm > 12) Then

debug_info = "Error! value 1-12 only"

returnVal = False

End If

If (dd < 1 Or dd > 32) Then

debug_info = "Error! value 1-32 only"

returnVal = False

End If

is_range_nep = returnVal

End Function

Public Sub eng_to_nep(yy, mm, dd)


Dim month, lmonth


Dim def_eyy, def_nyy, def_nmm, def_ndd
Dim total_eDays, total_nDays, a, day
Dim m, Y, i, j
Dim numDay
On Error Resume Next

Call initilizeClass
month = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
lmonth = Array(31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

def_eyy = 1944
def_nyy = 2000: def_nmm = 9: def_ndd = 17 - 1
total_eDays = 0: total_nDays = 0: a = 0: day = 7 - 1
m = 0: Y = 0: i = 0: j = 0
numDay = 0


For i = 0 To yy - def_eyy - 1
 If (is_leap_year(def_eyy + i) = True) Then
  For j = 0 To 12 - 1
   total_eDays = total_eDays + lmonth(j)
  Next j
Else
For j = 0 To 12 - 1
total_eDays = total_eDays + month(j)
Next j
End If
Next i

'‘// count total no. of days in-terms of month

For i = 0 To ((mm - 1) - 1)
If (is_leap_year(yy) = True) Then
total_eDays = total_eDays + lmonth(i)
Else
total_eDays = total_eDays + month(i)
End If
Next i

'‘ // count total no. of days in-terms of date

total_eDays = total_eDays + dd
i = 0: j = def_nmm
total_nDays = def_ndd
m = def_nmm
Y = def_nyy

'‘// count nepali date from array

Do While (total_eDays <> 0)
a = bs(i)(j)
total_nDays = total_nDays + 1
day = day + 1
If (total_nDays > a) Then
m = m + 1
total_nDays = 1
j = j + 1
End If
If (day > 7) Then
day = 1
End If
If (m > 12) Then
Y = Y + 1
m = 1
End If
If (j > 12) Then
j = 1: i = i + 1
End If
total_eDays = total_eDays - 1
Loop




nep_year = Y
nep_month = m
nep_date = total_nDays
nep_day = day

End Sub


Public Sub nep_to_eng(yy, mm, dd)
Dim def_eyy, def_emm, def_edd

Dim def_nyy, def_nmm, def_ndd

Dim total_eDays, total_nDays, a, dayy

Dim m, Y, i, j
Dim k, numDay

Dim month, lmonth
Call initilizeClass









def_eyy = 1943: def_emm = 4: def_edd = 14 - 1
def_nyy = 2000: def_nmm = 1: def_ndd = 1
total_eDays = 0: total_nDays = 0: a = 0: dayy = 4 - 1
m = 0: Y = 0: i = 0

k = 0: numDay = 0

month = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

lmonth = Array(0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)



For i = 0 To (yy - def_nyy - 1)

For j = 1 To 12

total_nDays = total_nDays + bs(i)(j)

Next j

k = k + 1

Next i


For j = 1 To mm - 1

total_nDays = total_nDays + bs(k)(j)

Next j


total_nDays = total_nDays + dd


total_eDays = def_edd

m = def_emm
Y = def_eyy

Do While (total_nDays <> 0)

If (is_leap_year(Y) = True) Then

a = lmonth(m)

Else

a = month(m)

End If

total_eDays = total_eDays + 1

dayy = dayy + 1

If (total_eDays > a) Then

m = m + 1

total_eDays = 1
End If

If (m > 12) Then

Y = Y + 1

m = 1


End If


If (dayy > 7) Then
dayy = 1
End If

total_nDays = total_nDays - 1



Loop

eng_year = Y
eng_month = m
eng_date = total_eDays
eng_day = dayy

End Sub
 Public Sub adjustlocation()
  On Error Resume Next
If devcalender.Left > Screen.Width - devcalender.Width Then devcalender.Left = Screen.Width - devcalender.Width
If devcalender.Left < 0 Then devcalender.Left = 0
If devcalender.Top < 0 Then devcalender.Top = 0
If devcalender.Top > Screen.Height - devcalender.Height - 400 Then devcalender.Top = Screen.Height - devcalender.Height - 400


On Error Resume Next

Open App.Path & "\tools\location.set" For Output As #1
Write #1, devcalender.Top
Write #1, devcalender.Left
Close #1
 End Sub


