Attribute VB_Name = "Cursos_Black_Func"

Function CP_Black(tipo, t, strike, Forward, sigma, df)
    tipo = UCase(tipo)
    If tipo = "C" Then factor = 1 Else factor = -1
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    d2 = d1 - Sqr(t) * sigma
    CP_Black = df * (factor * Forward * Cumnorm(factor * d1) - factor * strike * Cumnorm(factor * d2))
End Function

Function CP_Black_Delta(tipo, t, strike, Forward, sigma, df_usd)
    tipo = UCase(tipo)
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    If tipo = "C" Then
        CP_Black_Delta = df_usd * CND(d1)
    Else
        CP_Black_Delta = df_usd * (CND(d1) - 1)
    End If
End Function
Function CP_Black_DeltaFwd(tipo, t, strike, Forward, sigma, df_clp)
    tipo = UCase(tipo)
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    If tipo = "C" Then
        CP_Black_DeltaFwd = df_clp * CND(d1)
    Else
        CP_Black_DeltaFwd = df_clp * (CND(d1) - 1)
    End If
End Function
Function CP_Black_Gamma(t, strike, spot, Forward, sigma, df_usd)
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    aux = 1 / (Sqr(2 * Pi)) * Exp(-0.5 * d1 ^ 2)
    CP_Black_Gamma = aux * df_usd / (spot * sigma * Sqr(t))
End Function
Function CP_Black_Rho_Dom(tipo, t, strike, df_dom, Forward, sigma)
    tipo = UCase(tipo)
    If tipo = "C" Then factor = 1 Else factor = -1
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    d2 = d1 - Sqr(t) * sigma
    CP_Black_Rho_Dom = factor * strike * t * df_dom * CND(factor * d2)
End Function
Function CP_Black_Rho_For(tipo, t, strike, df_for, spot, Forward, sigma)
    tipo = UCase(tipo)
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    If tipo = "C" Then
        CP_Black_Rho_For = -t * df_for * spot * CND(d1)
    Else
        CP_Black_Rho_For = t * df_for * spot * CND(-d1)
    End If
End Function
Function CP_Black_Theta(tipo, t, strike, r_dom, r_for, spot, Forward, sigma)
    tipo = UCase(tipo)
    If tipo = "C" Then factor = 1 Else factor = -1
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    d2 = d1 - Sqr(t) * sigma
    aux = 1 / (Sqr(2 * Pi)) * Exp(-0.5 * d1 ^ 2)
    df_dom = Exp(-r_dom * t)
    df_for = Exp(-r_for * t)
    CP_Black_Theta = -spot * aux * sigma * df_for / (2# * Sqr(t)) + factor * r_for * spot * CND(factor * d1) * df_for - factor * r_dom * strike * df_dom * CND(factor * d2)
End Function
Function CP_Black_Vega(t, strike, spot, Forward, sigma, df_usd)
    d1 = (Log(Forward / strike) + 0.5 * t * sigma ^ 2) / (Sqr(t) * sigma)
    aux = 1 / (Sqr(2 * Pi)) * Exp(-0.5 * d1 ^ 2)
    CP_Black_Vega = spot * Sqr(t) * aux * df_usd
End Function
Function VolSmile(atm, rr1, vwb1, rr2, vwb2)
    Dim aux(1 To 5) As Double
    aux(1) = atm + vwb2 - 0.5 * rr2
    aux(2) = atm + vwb1 - 0.5 * rr1
    aux(3) = atm
    aux(4) = atm + vwb1 + 0.5 * rr1
    aux(5) = atm + vwb2 + 0.5 * rr2
    VolSmile = aux
End Function

Function FwdStrike(fwd, sigma, t, tipo, delta, rd)
    tipo = UCase(tipo)
    If tipo = "C" Then cpf = -1 Else cpf = 1
    alpha = ICND(delta * Exp(rd * t))
    FwdStrike = fwd * Exp(cpf * alpha * sigma * Sqr(t) + 0.5 * t * sigma ^ 2)

End Function
Function StrikeSmile(Plazo, put2, put1, atm, call1, call2, spot, tenorsClp, ratesClp, tenorsCld, ratesCld, delta1, delta2)
    Dim aux(1 To 5) As Double
    rd = GetDiscountFactorFromCurve(Plazo, tenorsClp, ratesClp, 360, 2)
    fwd = spot / rd
    fwd = fwd * GetDiscountFactorFromCurve(Plazo, tenorsCld, ratesCld, 360, 2)
    rd = -365 / Plazo * Log(rd)
    t = Plazo / 365
    
    aux(1) = FwdStrike(fwd, put2 / 100, t, "p", delta2, rd)
    aux(2) = FwdStrike(fwd, put1 / 100, t, "p", delta1, rd)
    aux(3) = fwd * Exp(0.5 * t * (atm / 100) ^ 2)
    aux(4) = FwdStrike(fwd, call1 / 100, t, "C", delta1, rd)
    aux(5) = FwdStrike(fwd, call2 / 100, t, "C", delta2, rd)
    
    StrikeSmile = aux
    
End Function


Function fxForward(Plazo, spot, tenorsClp, ratesClp, tenorsusd, ratesusd, Basis, Compound)
    dfclp = GetDiscountFactorFromCurve(Plazo, tenorsClp, ratesClp, Basis, Compound)
    dfusd = GetDiscountFactorFromCurve(Plazo, tenorsusd, ratesusd, Basis, Compound)
    fxForward = spot * dfusd / dfclp
    
End Function

