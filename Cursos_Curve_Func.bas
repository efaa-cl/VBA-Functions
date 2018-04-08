Attribute VB_Name = "Cursos_Curve_Func"
Function SearchIndexLinInterpol(Vector, x)

Size = Application.CountA(Vector)

Maxim = Size
Minim = 1
Dim Centro As Integer
Centro = Int((Maxim + Minim) / 2)

' Va al entero anterior por ser Integer

If Vector(Minim) > Vector(Maxim) Then SearchIndexLinInterpol = " Vector debe ser creciente": Exit Function

Dim i As Integer


i = 1
While Minim <= Maxim Or i < 10

If x = Vector(Centro) Then
    
    SearchIndexLinInterpol = Centro: Exit Function

Else
        
    If x < Vector(Centro) Then
    Maxim = Centro - 1
        
    Else
    Minim = Centro + 1
    
    End If
    
 End If
 
    Centro = Int((Maxim + Minim) / 2)
    i = i + 1
    Wend
    'solo por control
    If Centro < 1 Then SearchIndexLinInterpol = 1: Exit Function
    
    SearchIndexLinInterpol = Centro
       
End Function
'-----------------------------------------------------------

Function linInterpol(plazos, tasas, Plazo)
    TaM = Application.CountA(plazos)
    n = Application.CountA(tasas)
    If TaM <> n Then Exit Function
    If Plazo >= plazos(TaM) Then linInterpol = tasas(TaM): Exit Function
    If Plazo <= plazos(1) Then linInterpol = tasas(1): Exit Function
    
    Dim Indice As Integer
    Dim Slope As Double
    
    Indice = SearchIndexLinInterpol(plazos, Plazo)
    
    Slope = (tasas(Indice + 1) - tasas(Indice)) / (plazos(Indice + 1) - plazos(Indice))
    
    linInterpol = Slope * (Plazo - plazos(Indice)) + tasas(Indice)

End Function
Function loglinInterpol(plazos, factores, Plazo)
    TaM = Application.CountA(plazos)
    n = Application.CountA(factores)
    If TaM <> n Then Exit Function
    If Plazo >= plazos(TaM) Then loglinInterpol = factores(TaM): Exit Function
    If Plazo <= plazos(1) Then loglinInterpol = factores(1): Exit Function
    
    Dim Indice As Integer
    Dim Slope As Double
    
    Indice = SearchIndexLinInterpol(plazos, Plazo)
    
    Slope = (Log(factores(Indice + 1)) - Log(factores(Indice))) / (plazos(Indice + 1) - plazos(Indice))
    
    loglinInterpol = Exp(Slope * (Plazo - plazos(Indice)) + Log(factores(Indice)))

End Function
    
Function GetDiscountFactorFromCurve(Days, Tenors, Rates, Basis, Compound)
'Compund =1 es Lineal, =2 es compuesto, otra cosa es compuesto
    Dim rate As Variant
            rate = linInterpol(Tenors, Rates, Days)
            Select Case Compound
                Case 1
                    GetDiscountFactorFromCurve = 1 / (1 + rate * Days / Basis)
                Case 2
                    GetDiscountFactorFromCurve = 1 / (1 + rate) ^ (Days / Basis)
                Case 3
                    GetDiscountFactorFromCurve = Exp(-rate * Days / Basis)
                Case Else
                    GetDiscountFactorFromCurve = 1 / (1 + rate) ^ (Days / Basis)
            End Select
End Function

Function GetDiscountFactorFwdFromCurve(Days1, Days2, Tenors, Rates, Basis, Compound)
'compound = 1 es Lineal, compound = 2 es Compuesto, compound = 3 es Exponencial
'compound = otro valor es compuesto

    Rate1 = linInterpol(Tenors, Rates, Days1)
    Rate2 = linInterpol(Tenors, Rates, Days2)
    
    Select Case Compound
        Case 1
            aux1 = 1 / (1 + Rate1 * Days1 / Basis)
            aux2 = 1 / (1 + Rate2 * Days2 / Basis)
        Case 2
            aux1 = 1 / (1 + Rate1) ^ (Days1 / Basis)
            aux2 = 1 / (1 + Rate2) ^ (Days2 / Basis)
        Case 3
            aux1 = Exp(-Rate1 * Days1 / Basis)
            aux2 = Exp(-Rate2 * Days2 / Basis)
        Case Else
            aux1 = 1 / (1 + Rate1) ^ (Days1 / Basis)
            aux2 = 1 / (1 + Rate2) ^ (Days2 / Basis)
        End Select
        
    GetDiscountFactorFwdFromCurve = aux2 / aux1

End Function

