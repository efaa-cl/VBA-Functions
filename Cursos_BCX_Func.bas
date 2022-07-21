Attribute VB_Name = "Cursos_BCX_Func"
Function BCX_MatDate(nemo)
    BCX_MatDate = DateSerial(2000 + CInt(Mid(nemo, 9, 2)), Mid(nemo, 7, 2), 1)
End Function

Function BCX_Cupon(nemo)
    BCX_Cupon = Val(CInt(Mid(nemo, 5, 2))) / 1000
End Function

Function BCX_ValorPar(Fecha_Val, FechaMat, Cupon, Tera)
    If Fecha_Val > FechaMat Then BCX_ValorPar = 0: Exit Function
    npagos1 = CInt(DateDiff("m", Fecha_Val, FechaMat) / 6) + 2
    ReDim fechapago(1 To npagos1) As Date
    ReDim flujopago(1 To npagos1) As Double
    fechapago(1) = FechaMat
    flujopago(1) = 100 * (1 + Cupon / 2)
    npagos = 1
    For i = 1 To npagos1 + 1
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then fechapago(i) = DateAdd("m", -(i - 1) * 6, FechaMat)
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then flujopago(i) = Nom * Cupon / 2
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then npagos = i
    Next i
'Calculo del VPar
    VPar = 100 * (1 + Tera) ^ YearFraction(DateAdd("m", -6, fechapago(npagos)), Fecha_Val, "Act/365")
    BCX_ValorPar = VPar
End Function

Function BCX_PvFromTirAndNemo(Fecha_Val, Tir, nemo)

    FechaMat = BCX_MatDate(nemo)
    Cupon = BCX_Cupon(nemo)
    
    BCX_PvFromTirAndNemo = BCX_PvFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    
End Function

Function BCX_PvFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    If Fecha_Val > FechaMat Then BCX_PvFromTir = 0: Exit Function
    npagos1 = CInt(DateDiff("m", Fecha_Val, FechaMat) / 6) + 2
    ReDim fechapago(1 To npagos1) As Date
    ReDim flujopago(1 To npagos1) As Double
    fechapago(1) = FechaMat
    flujopago(1) = 100 * (1 + Cupon / 2)
    npagos = 1
    For i = 2 To npagos1 + 1
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then fechapago(i) = DateAdd("m", -(i - 1) * 6, FechaMat)
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then flujopago(i) = 100 * Cupon / 2
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then npagos = i
    Next i
'Calculo Valor Presente
    i = npagos
    BCX_PvFromTir = 0
    Do
        BCX_PvFromTir = BCX_PvFromTir + flujopago(i) / (1 + Tir) ^ (YearFraction(Fecha_Val, fechapago(i), "Act/365"))
        i = i - 1
    Loop Until i < 1
End Function

Function BCX_DurationFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    If Fecha_Val > FechaMat Then BCX_DurationFromTir = 0: Exit Function
    npagos1 = CInt(DateDiff("m", Fecha_Val, FechaMat) / 6) + 2
    ReDim fechapago(1 To npagos1) As Date
    ReDim flujopago(1 To npagos1) As Double
    fechapago(1) = FechaMat
    flujopago(1) = 100 * (1 + Cupon / 2)
    npagos = 1
    For i = 2 To npagos1 + 1
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then fechapago(i) = DateAdd("m", -(i - 1) * 6, FechaMat)
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then flujopago(i) = 100 * Cupon / 2
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then npagos = i
    Next i
'Calculo Valor Presente
    i = npagos
    BCX_DurationFromTir = 0
    aux_pv = 0
    Do
        yf = YearFraction(Fecha_Val, fechapago(i), "Act/365")
        vp_cupon = flujopago(i) / (1 + Tir) ^ (yf)
        BCX_DurationFromTir = BCX_DurationFromTir + yf * vp_cupon
        aux_pv = aux_pv + vp_cupon
        i = i - 1
    Loop Until i < 1
    BCX_DurationFromTir = BCX_DurationFromTir / aux_pv
End Function

Function BCX_ConvexityFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    If Fecha_Val > FechaMat Then BCX_ConvexityFromTir = 0: Exit Function
    npagos1 = CInt(DateDiff("m", Fecha_Val, FechaMat) / 6) + 2
    ReDim fechapago(1 To npagos1) As Date
    ReDim flujopago(1 To npagos1) As Double
    fechapago(1) = FechaMat
    flujopago(1) = 100 * (1 + Cupon / 2)
    npagos = 1
    For i = 2 To npagos1 + 1
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then fechapago(i) = DateAdd("m", -(i - 1) * 6, FechaMat)
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then flujopago(i) = 100 * Cupon / 2
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then npagos = i
    Next i
'Calculo Valor Presente
    i = npagos
    BCX_ConvexityFromTir = 0
    aux_pv = 0
    Do
        yf = YearFraction(Fecha_Val, fechapago(i), "Act/365")
        vp_cupon = flujopago(i) / (1 + Tir) ^ (yf)
        BCX_ConvexityFromTir = BCX_ConvexityFromTir + yf * (yf + 1) * vp_cupon
        aux_pv = aux_pv + vp_cupon
        i = i - 1
    Loop Until i < 1
    BCX_ConvexityFromTir = BCX_ConvexityFromTir / (aux_pv * (1 + Tir) ^ 2)
End Function

Function BCX_PvDurConvFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    If Fecha_Val > FechaMat Then BCX_PvDurConvFromTir = 0: Exit Function
    npagos1 = CInt(DateDiff("m", Fecha_Val, FechaMat) / 6) + 2
    ReDim fechapago(1 To npagos1) As Date
    ReDim flujopago(1 To npagos1) As Double
    Dim output(1 To 3) As Double
    fechapago(1) = FechaMat
    flujopago(1) = 100 * (1 + Cupon / 2)
    npagos = 1
    For i = 2 To npagos1 + 1
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then fechapago(i) = DateAdd("m", -(i - 1) * 6, FechaMat)
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then flujopago(i) = 100 * Cupon / 2
        If DateAdd("m", -(i - 1) * 6, FechaMat) > Fecha_Val Then npagos = i
    Next i
    
    i = npagos
    present_value = 0
    duracion = 0
    convexidad = 0
    Do
        yf = YearFraction(Fecha_Val, fechapago(i), "Act/365")
        vp_cupon = flujopago(i) / (1 + Tir) ^ (yf)
        present_value = present_value + vp_cupon
        duracion = duracion + yf * vp_cupon
        convexidad = convexidad + yf * (yf + 1) * vp_cupon
        i = i - 1
    Loop Until i < 1
    output(1) = present_value
    output(2) = duracion / present_value
    output(3) = convexidad / (present_value * (1 + Tir) ^ 2)
    BCX_PvDurConvFromTir = output
End Function

Function BCX_QuoteFromTir(Fecha_Val, FechaMat, Cupon, Tir, Tera)
    If Fecha_Val > FechaMat Then BCX_QuoteFromTir = 0: Exit Function
    
    'Calculo del VPar
    VPar = BCX_ValorPar(Fecha_Val, FechaMat, Cupon, Tera)
    
    'Ca lculo Valor Presente
    BCX = BCX_PvFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    
    'Calculo del precio
    BCX_QuoteFromTir = Int(BCX / VPar * 1000000 + 0.5) / 1000000
End Function

Function BCX_ValorCompraFromTir(Nominal, Fecha_Val, FechaMat, Cupon, Tir, Tera)
    If Fecha_Val > FechaMat Then BCX_ValorCompraFromTir = 0: Exit Function
    
    'Calculo del VPar
    VPar = BCX_ValorPar(Fecha_Val, FechaMat, Cupon, Tera)
    
    'Ca lculo Valor Presente
    BCX = BCX_PvFromTir(Fecha_Val, FechaMat, Cupon, Tir)
    
    'Calculo del precio
    precio = Int(BCX / VPar * 1000000 + 0.5) / 1000000
    
    BCX_ValorCompraFromTir = precio * VPar * Nominal / 100
    
End Function


Function BCX_FindTirFromPV(FechaVal, FechaMat, Cupon, PV)
    'Return the Dirty price of a bond paying Coupon (expresed as a percentage)
    'Freq times per year. BasisCount is the number of days of the year
    If FechaVal > FechaMat Then BCX_FindTirFromPV = 0: Exit Function
    Tir = 0.02: diff = 1: dp = 1E+20
        While Abs(diff) > 1E-06
            Tir = Tir - diff / dp
            p = BCX_PvFromTir(FechaVal, FechaMat, Cupon, Tir)
            dp = (BCX_PvFromTir(FechaVal, FechaMat, Cupon, Tir + 0.01) - p) / 0.01
            diff = p - PV
        Wend
    BCX_FindTirFromPV = Tir
End Function

Function BCX_Moneda(nemo)

    If Left(nemo, 3) = "BCP" Or Left(nemo, 3) = "BTP" Then
        BCX_Moneda = "CLP"
    Else
        BCX_Moneda = "CLF"
    End If
    
End Function

Function BCX_ValorCompra(nemo, precio, valor_par, posicion, uf)

    If BCX_Moneda(nemo) = "CLP" Then
        BCX_ValorCompra = Round(precio * valor_par * posicion / 100, 0)
    Else
        BCX_ValorCompra = Round(uf * precio * valor_par * posicion / 100, 0)
    End If
    
End Function
