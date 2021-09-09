Attribute VB_Name = "Module1"
Option Explicit

Global Const dNoError = 0
Global Const dDayError = 1
Global Const dMonthError = 2
Global Const dYearError = 4

Global Month(12) As String
Global Week(6) As String

Function ExtraYear(Year As Long) As Boolean
    ExtraYear = (Year Mod 400 = 0) Or ((Year Mod 100 <> 0) And (Year Mod 4 = 0))
End Function

Sub InitMonth()
    Month(1) = "January"
    Month(2) = "February"
    Month(3) = "March"
    Month(4) = "April"
    Month(5) = "May"
    Month(6) = "June"
    Month(7) = "July"
    Month(8) = "August"
    Month(9) = "September"
    Month(10) = "October"
    Month(11) = "November"
    Month(12) = "December"
End Sub

Sub InitWeek()
    Week(0) = "Sunday"
    Week(1) = "Monday"
    Week(2) = "Tuesday"
    Week(3) = "Wednesday"
    Week(4) = "Thursday"
    Week(5) = "Friday"
    Week(6) = "Saturday"
End Sub

Function Validate(Day As Byte, Month As Byte, Year As Long) As Byte
Dim ErrorCode As Byte
    ErrorCode = dNoError
    If (Month < 1) Or (Month > 12) Then
        ErrorCode = ErrorCode + dMonthError
    End If
    Select Case Month
        Case 1, 3, 5, 7, 8, 10, 12
            If (Day < 1) Or (Day > 31) Then
                ErrorCode = ErrorCode + dDayError
            End If
        Case 4, 6, 9, 11
            If (Day < 1) Or (Day > 30) Then
                ErrorCode = ErrorCode + dDayError
            End If
        Case 2
            If ExtraYear(Year) Then
                If (Day < 1) Or (Day > 29) Then
                    ErrorCode = ErrorCode + dMonthError
                End If
            Else
                If (Day < 1) Or (Day > 28) Then
                    ErrorCode = ErrorCode + dMonthError
                End If
            End If
    End Select
    Validate = ErrorCode
End Function

Function CalculateWeekday(Day As Byte, Month As Byte, Year As Long) As Byte
Dim C As Integer
Dim X As Long
Dim S As Integer
Dim W As Integer
Dim FebDay As Byte
    X = Year
    C = 0
    If ExtraYear(X) Then
        FebDay = 29
    Else
        FebDay = 28
    End If
    Select Case Month
        Case 1
            C = Day
        Case 2
            C = 31 + Day
        Case 3
            C = 31 + FebDay + Day
        Case 4
            C = 31 + FebDay + 31 + Day
        Case 5
            C = 31 + FebDay + 31 + 30 + Day
        Case 6
            C = 31 + FebDay + 31 + 30 + 31 + Day
        Case 7
            C = 31 + FebDay + 31 + 30 + 31 + 30 + Day
        Case 8
            C = 31 + FebDay + 31 + 30 + 31 + 30 + 31 + Day
        Case 9
            C = 31 + FebDay + 31 + 30 + 31 + 30 + 31 + 31 + Day
        Case 10
            C = 31 + FebDay + 31 + 30 + 31 + 30 + 31 + 31 + 30 + Day
        Case 11
            C = 31 + FebDay + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + Day
        Case 12
            C = 31 + FebDay + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + Day
    End Select
    S = X - 1 + Fix((X - 1) / 4) - Fix((X - 1) / 100) + Fix((X - 1) / 400) + C
    W = S Mod 7
    If W >= 0 Then
        CalculateWeekday = W
    Else
        CalculateWeekday = 7 + W
    End If
End Function
