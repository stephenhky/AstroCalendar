Attribute VB_Name = "Module2"
Option Explicit
Type ABC
    A As Integer
    B As Single
    C As Double
End Type
Global Const Spring = 1
Global Const Summer = 2
Global Const Autumn = 3
Global Const Winter = 4
Global JDE As Double
Dim JDE0, T, W, Lamda, S1 As Double
Dim S(24) As ABC

Private Function JDEN(Year As Long, Season As Byte) As Double
Dim Y As Double
    If (Year > -1000) And (Year < 1000) Then
        Y = Year / 1000
        Select Case Season
            Case Spring
                JDEN = 1721139.29189 + 365242.1374 * Y + 0.06134 * Y ^ 2 + 0.00111 * Y ^ 3 - 0.00071 * Y ^ 4
            Case Summer
                JDEN = 1721233.25401 + 365241.72562 * Y - 0.05323 * Y ^ 2 + 0.00907 * Y ^ 3 + 0.00025 * Y ^ 4
            Case Autumn
                JDEN = 1721325.70455 + 365242.49558 * Y - 0.11677 * Y ^ 2 - 0.00297 * Y ^ 3 + 0.00074 * Y ^ 4
            Case Winter
                JDEN = 1721414.39987 + 365242.88257 * Y - 0.00769 * Y ^ 2 - 0.00933 * Y ^ 3 - 0.00006 * Y ^ 4
        End Select
    ElseIf (Year >= 1000) And (Year < 3000) Then
        Y = (Year - 2000) / 1000
        Select Case Season
            Case Spring
                JDEN = 2451623.80984 + 365242.37404 * Y + 0.05169 * Y ^ 2 - 0.00411 * Y ^ 3 - 0.00057 * Y ^ 4
            Case Summer
                JDEN = 2451716.56767 + 365241.62603 * Y + 0.00325 * Y ^ 2 + 0.00888 * Y ^ 3 - 0.0003 * Y ^ 4
            Case Autumn
                JDEN = 2451810.21715 + 365242.01767 * Y - 0.11575 * Y ^ 2 + 0.00337 * Y ^ 3 + 0.00078 * Y ^ 4
            Case Winter
                JDEN = 2451900.05952 + 365242.74049 * Y - 0.06223 * Y ^ 2 - 0.00823 * Y ^ 3 + 0.00032 * Y ^ 4
        End Select
    End If
End Function

Private Sub CalculateTWL(Year As Long, Season As Byte)
Dim Pi As Double
    Pi = 4 * Atn(1)
    JDE0 = JDEN(Year, Season)
    T = (JDE0 - 2451545) / 36525
    W = 35999.373 * T - 2.47
    Lamda = 1 + 0.0334 * Cos(W * Pi / 180) + 0.0007 * Cos(2 * W * Pi / 180)
End Sub

Sub InitS()
    S(1).A = 485
    S(1).B = 324.96
    S(1).C = 1934.136
    S(2).A = 203
    S(2).B = 337.23
    S(2).C = 32964.467
    S(3).A = 199
    S(3).B = 342.08
    S(3).C = 20.186
    S(4).A = 182
    S(4).B = 27.85
    S(4).C = 445267.112
    S(5).A = 156
    S(5).B = 73.14
    S(5).C = 45036.886
    S(6).A = 136
    S(6).B = 171.52
    S(6).C = 22518.443
    S(7).A = 77
    S(7).B = 222.54
    S(7).C = 65928.934
    S(8).A = 74
    S(8).B = 296.72
    S(8).C = 3034.906
    S(9).A = 70
    S(9).B = 243.58
    S(9).C = 9037.513
    S(10).A = 58
    S(10).B = 119.81
    S(10).C = 33718.147
    S(11).A = 52
    S(11).B = 297.17
    S(11).C = 150.678
    S(12).A = 50
    S(12).B = 21.02
    S(12).C = 2281.226
    S(13).A = 45
    S(13).B = 247.54
    S(13).C = 29929.562
    S(14).A = 44
    S(14).B = 325.15
    S(14).C = 31555.956
    S(15).A = 29
    S(15).B = 60.93
    S(15).C = 4443.417
    S(16).A = 18
    S(16).B = 155.12
    S(16).C = 67555.328
    S(17).A = 17
    S(17).B = 288.79
    S(17).C = 4562.452
    S(18).A = 16
    S(18).B = 198.04
    S(18).C = 62894.029
    S(19).A = 14
    S(19).B = 199.76
    S(19).C = 31436.921
    S(20).A = 12
    S(20).B = 95.39
    S(20).C = 14577.848
    S(21).A = 12
    S(21).B = 287.11
    S(21).C = 31931.756
    S(22).A = 12
    S(22).B = 320.81
    S(22).C = 34777.259
    S(23).A = 9
    S(23).B = 227.73
    S(23).C = 1222.114
    S(24).A = 8
    S(24).B = 15.45
    S(24).C = 16859.074
End Sub

Private Sub CalculateS()
Dim I As Byte
Dim Pi As Double
    Pi = 4 * Atn(1)
    S1 = 0
    For I = 1 To 24
        S1 = S1 + S(I).A * Cos((S(I).B + S(I).C * T) * Pi / 180)
    Next I
End Sub

Sub CalculateJDE(Year As Long, Season As Byte)
    CalculateTWL Year, Season
    CalculateS
    JDE = JDE0 + 0.00001 * S1 / Lamda
End Sub

Sub JDtoCD(ByVal JD As Double, Y As Long, M As Byte, Day As Single)
Dim A, B, C, D, E As Integer
Dim F As Single
Dim Z As Long
Dim alpha As Double
    Z = Int(JD + 0.5)
    F = JD + 0.5 - Z
    If Z < 2299161 Then
        A = Z
    Else
        alpha = Int((Z - 1867216.25) / 36524.25)
        A = Z + 1 + alpha - Int(alpha / 4)
    End If
    B = A + 1524
    C = Int((B - 122.1) / 365.25)
    D = Int(365.25 * C)
    E = Int((B - D) / 30.6001)
    
    Day = B - D - Int(30.6001 * E) + F
    
    If (E < 14) Then
        M = E - 1
    ElseIf (E = 14) Or (E = 15) Then
        M = E - 13
    End If
    
    If (M > 2) Then
        Y = C - 4716
    ElseIf (M = 1) Or (M = 2) Then
        Y = C - 4715
    End If
End Sub

Sub FracToTime(ByVal Frac, Hour As Byte, Min As Byte, Sec As Single)
Dim F As Single
    F = Frac - Int(Frac)
    
    F = F * 24
    Hour = Int(F)
    F = F - Int(F)
    
    F = F * 60
    Min = Int(F)
    F = F - Int(F)
    
    F = F * 60
    Sec = F
End Sub
