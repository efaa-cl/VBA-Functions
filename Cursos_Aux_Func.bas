Attribute VB_Name = "Cursos_Aux_Func"
Public Function thirdWednesday(fecha As Date)

    thirdWednesday = DateSerial(Year(fecha), Month(fecha), 8) - _
                     Weekday(DateSerial(Year(fecha), Month(fecha), 4)) + 14
    
End Function

Function interest(startDate, endDate, rate, typeOfRate, notional)

    typeOfRate = LCase(typeOfRate)
    aux = Len(typeOfRate) - 4

    If Left(typeOfRate, 3) = "lin" Then
        interest = notional * rate * YearFraction(startDate, endDate, Right(typeOfRate, aux))
    Else
        interest = notional * ((1 + rate) ^ YearFraction(startDate, endDate, Right(typeOfRate, aux)) - 1)
    End If
    
End Function

Function Indicador(a, b)

    If a = b Then Indicador = 1 Else Indicador = 0
    
End Function

Function Ten(Tn)
'Esto es un comentario
    If UCase(Right(Tn, 1)) = "C" Then
        Tn = Left(Tn, Len(Tn) - 1)
    End If
    If Tn = "Spot" Then
        Ten = 0
    ElseIf UCase(Right(Tn, 1)) = "D" Then
        Ten = 1
    ElseIf UCase(Right(Tn, 1)) = "W" Then
        Ten = Val(Tn) * 7
    Else
        If UCase(Right(Tn, 1)) = "M" Then Ten = 1 Else Ten = 12
        Ten = Val(Tn) * Ten
    End If
End Function
Function localTenor(startDate, Tn)
    aux = Ten(Tn)
    aux1 = AddMonths(startDate, aux)
    localTenor = DateSerial(Year(aux1), Month(aux1), 9)
End Function

Function BussDay(ByVal a)
    If Weekday(a) = 1 Then a = a + 1
    If Weekday(a) = 7 Then a = a + 2
    BussDay = a
End Function
Function lag(ByVal startDate, step)
    Select Case step
        Case 0
            lag = Prev2(startDate)
        Case Else
            lag = Prev2(startDate)
            For i = 1 To step
                lag = Prev2(lag - 1)
            Next i
    End Select
End Function
Function shift(ByVal startDate, step)
    Select Case step
        Case 0
            shift = BussDay(startDate)
        Case Else
            shift = BussDay(startDate)
            For i = 1 To step
                shift = BussDay(shift + 1)
            Next i
    End Select
End Function

Function ModBussDay(ByVal a)
    m1 = Month(a)
    b = BussDay(a)
    m2 = Month(b)
    If m2 <> m1 Then
        b = b - 1
        w = Weekday(b)
        m2 = Month(b)
        While (w = 1 Or w = 7) Or m2 <> m1
            b = b - 1
            w = Weekday(b)
            m2 = Month(b)
        Wend
    End If
    ModBussDay = b
End Function

Function Prev2(ByVal a)
    If Application.Weekday(a) = 1 Then a = a - 2
    If Application.Weekday(a) = 7 Then a = a - 1
    Prev2 = a
End Function

Function Max(a, b)
    Max = a
    If b > a Then Max = b
End Function


Function Min(a, b)
    Min = a
    If b < a Then Min = b
End Function

Function CountDays(t1, t2, Basis)
    'Count days between dates t1 and t2 accordingly to the day count convention Basis
    'Valid Basis arguments are: Act/360, Act/365 and 30/360
    Select Case LCase(Basis)
    Case "act/360", "act/365"
        CountDays = (t2 - t1)
    Case "30/360"
        d1 = Day(t1): d2 = Day(t2): m1 = Month(t1): m2 = Month(t2): y1 = Year(t1): y2 = Year(t2)
        If d1 = 31 Then d1 = 30
        If d2 = 31 And d1 = 30 Then d2 = 30
        CountDays = (d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1)
        'CountDays = (Max(0, 30 - Day(t1)) + Min(30, Day(t2)) + 30 * (Month(t2) - Month(t1) - 1) + 360 * (Year(t2) - Year(t1)))
    Case Else   'Default
        CountDays = (t2 - t1)
    End Select
End Function


Function YearFraction(ByVal t1, ByVal t2, ByVal Basis) As Double
    'Computes the fraction of year between dates t1 and t2 accordingly to the day count convention Basis
    'Valid Basis arguments are: Act/360, Act/365 and 30/360
    Select Case LCase(Basis)
    Case "act/360"
        YearFraction = (t2 - t1) / (3.6 * 100)
    Case "act/365"
        YearFraction = (t2 - t1) / (3.65 * 100)
    Case "30/360"
        d1 = Day(t1): d2 = Day(t2): m1 = Month(t1): m2 = Month(t2): y1 = Year(t1): y2 = Year(t2)
        If d1 = 31 Then d1 = 30
        If d2 = 31 And d1 = 30 Then d2 = 30
        YearFraction = ((d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1)) / (3.6 * 100)
    Case Else   'Default
        YearFraction = (t2 - t1) / (3.65 * 100)
    End Select
End Function

Function AddMonths(Dat, ByVal Mon)
    'equivalent to Excel Function EDATE()
    'add/sustracts Mon number of months to date Dat for positive/negative values of Mon
    Dim dm(12) As Integer
    Dim D, m, y As Integer
    dm(1) = 31: dm(2) = 28: dm(3) = 31: dm(4) = 30:   dm(5) = 31:   dm(6) = 30
    dm(7) = 31: dm(8) = 31: dm(9) = 30: dm(10) = 31: dm(11) = 30: dm(12) = 31
    D = Day(Dat):    m = Month(Dat):    y = Year(Dat)

    If D = dm(m) And m <> 2 Then eom = True  'Detect if the date correspond to the end of a month
    Mon = Mon + m
    m = Mon - 12 * Int(Mon / 12): If m = 0 Then m = 12
    y = y + Int((Mon - 1) / 12)
    If y < 0 Then y = 1900
    If D > dm(m) Then eom = True

    If eom = True Then   'if the date correspond to the end of a month,
        D = dm(m)        'the resulting date is forced to be the end of a month
        If (Int(y / 4) = y / 4) And m = 2 Then D = 29
    End If
    AddMonths = DateSerial(y, m, D)
End Function
Function AddMonthsC(Dat, ByVal Mon)
    'equivalent to Excel Function EDATE()
    'add/sustracts Mon number of months to date Dat for positive/negative values of Mon
    Dim dm(12) As Integer
    Dim D, m, y As Integer
    dm(1) = 31: dm(2) = 28: dm(3) = 31: dm(4) = 30:   dm(5) = 31:   dm(6) = 30
    dm(7) = 31: dm(8) = 31: dm(9) = 30: dm(10) = 31: dm(11) = 30: dm(12) = 31
    D = Day(Dat):    m = Month(Dat):    y = Year(Dat)

    If D = dm(m) And m <> 2 Then eom = True  'Detect if the date correspond to the end of a month
    Mon = Mon + m
    m = Mon - 12 * Int(Mon / 12): If m = 0 Then m = 12
    y = y + Int((Mon - 1) / 12)
    If y < 0 Then y = 1900
    If D > dm(m) Then eom = True

    If eom = True Then   'if the date correspond to the end of a month,
        D = dm(m)        'the resulting date is forced to be the end of a month
        If (Int(y / 4) = y / 4) And m = 2 Then D = 29
    End If
    AddMonthsC = DateSerial(y, m, 9)
End Function

Function YearFracActualActual(inicial, final)
    dia_inicial = Day(inicial)
    mes_inicial = Month(inicial)
    a_o_inicial = Year(inicial)
    dia_final = Day(final)
    mes_final = Month(final)
    a_o_final = Year(final)
    If mes_final > mes_inicial Then
        aux1 = DateAdd("m", (a_o_final - a_o_inicial) * 12, inicial)
    ElseIf mes_final < mes_inicial Then
        aux1 = DateAdd("m", (a_o_final - a_o_inicial - 1) * 12, inicial)
    Else
        If dia_inicial <= dia_final Then
            aux1 = DateAdd("m", (a_o_final - a_o_inicial) * 12, inicial)
        Else
            aux1 = DateAdd("m", (a_o_final - a_o_inicial - 1) * 12, inicial)
        End If
    End If
    a_o1 = Year(aux1) - a_o_inicial
    aux2 = DateAdd("m", 12, aux1)
    a_o2 = (final - aux1) / (aux2 - aux1)
    YearFracActualActual = a_o1 + a_o2
End Function

Function RedondeoParcial(numero)
    entero = Int(numero)
    mantisa = numero - entero
    If mantisa >= 0 And mantisa < 0.25 Then
        RedondeoParcial = 0
    ElseIf mantisa >= 0.25 And mantisa < 0.75 Then
        RedondeoParcial = 0.5
    ElseIf mantisa >= 0.75 Then
         RedondeoParcial = 1
    End If
    RedondeoParcial = RedondeoParcial + entero
End Function

