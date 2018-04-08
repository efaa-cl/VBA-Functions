Attribute VB_Name = "Distributions"
Option Base 1
Global Const Pi = 3.14159265358979

Function Cumnorm(x As Double) As Double
  XAbs = Abs(x)
  If XAbs > 37 Then
    Cumnorm = 0
  Else
    Exponential = Exp(-XAbs ^ 2 / 2)
    If XAbs < 7.07106781186547 Then
      Build = 3.52624965998911E-02 * XAbs + 0.700383064443688
      Build = Build * XAbs + 6.37396220353165
      Build = Build * XAbs + 33.912866078383
      Build = Build * XAbs + 112.079291497871
      Build = Build * XAbs + 221.213596169931
      Build = Build * XAbs + 220.206867912376
      Cumnorm = Exponential * Build
      Build = 8.83883476483184E-02 * XAbs + 1.75566716318264
      Build = Build * XAbs + 16.064177579207
      Build = Build * XAbs + 86.7807322029461
      Build = Build * XAbs + 296.564248779674
      Build = Build * XAbs + 637.333633378831
      Build = Build * XAbs + 793.826512519948
      Build = Build * XAbs + 440.413735824752
      Cumnorm = Cumnorm / Build
    Else
      Build = XAbs + 0.65
      Build = XAbs + 4 / Build
      Build = XAbs + 3 / Build
      Build = XAbs + 2 / Build
      Build = XAbs + 1 / Build
      Cumnorm = Exponential / Build / 2.506628274631
    End If
End If
  If x > 0 Then Cumnorm = 1 - Cumnorm
End Function

Public Function ND(ByVal z As Double) As Double
    'Normal Distribution
    ND = 1 / Sqr(2 * Pi) * Exp(-z ^ 2 / 2)
End Function


Function CND(ByVal z As Double) As Double
Attribute CND.VB_ProcData.VB_Invoke_Func = " \n14"
    'Cumulative Normal Distribution
    a = Array(0.31938153, -0.356563782, 1.781477937, -1.821255978, 1.330274429)
    y = 1 / (1 + 0.2316419 * Abs(z))
    CND = (1 / Sqr(2 * Pi) * Exp(-0.5 * z ^ 2)) * y * (a(1) + y * (a(2) + y * (a(3) + y * (a(4) + y * a(5)))))
    If z > 0 Then CND = 1 - CND
End Function


Function ICND(u)
Attribute ICND.VB_ProcData.VB_Invoke_Func = " \n14"
    'Inverse of the Cumulative Normal Distribution
    'Beasley&Springer
    a = Array(2.50662823884, -18.61500062529, 41.39119773534, -25.44106049637)
    b = Array(-8.4735109309, 23.08336743743, -21.06224101826, 3.13082909833)
    c = Array(0.337475482272615, 0.976169019091719, 0.160797971491821, 2.76438810333863E-02, 3.8405729373609E-03, 3.951896511919E-04, 3.21767881768E-05, 2.888167364E-07, 3.960315187E-07)
     
    x = u - 0.5
    If Abs(x) <= 0.42 Then
        r = x ^ 2
        r = x * (a(1) + r * (a(2) + r * (a(3) + a(4) * r))) / (1 + r * (b(1) + r * (b(2) + r * (b(3) + r * b(4)))))
    Else
        r = u
        If x > 0 Then r = 1 - u
        r = Log(-Log(r))
        r = c(1) + r * (c(2) + r * (c(3) + r * (c(4) + r * (c(5) + r * (c(6) + r * (c(7) + r * (c(8) + r * c(9))))))))
        If x < 0 Then r = -r
    End If
        ICND = r
End Function


Function CBND(a As Double, b As Double, rho As Double) As Double
Attribute CBND.VB_ProcData.VB_Invoke_Func = " \n14"
'Cumulative Bivariate Normal Distribution
    Dim x As Variant, y As Variant
    Dim rho1 As Double, rho2 As Double, delta As Double
    Dim a1 As Double, b1 As Double, Sum As Double
    Dim i As Integer, j As Integer
    
    x = Array(0.24840615, 0.39233107, 0.21141819, 0.03324666, 0.00082485334)
    y = Array(0.10024215, 0.48281397, 1.0609498, 1.7797294, 2.6697604)
    
    a1 = a / Sqr(2 * (1 - rho ^ 2))
    b1 = b / Sqr(2 * (1 - rho ^ 2))
    
    If a <= 0 And b <= 0 And rho <= 0 Then
        Sum = 0
        For i = 1 To 5
            For j = 1 To 5
                Sum = Sum + x(i) * x(j) * Exp(a1 * (2 * y(i) - a1) + b1 * (2 * y(j) - b1) + 2 * rho * (y(i) - a1) * (y(j) - b1))
            Next
        Next
        CBND = Sqr(1 - rho ^ 2) / Pi * Sum
    ElseIf a <= 0 And b >= 0 And rho >= 0 Then
        CBND = CND(a) - CBND(a, -b, -rho)
    ElseIf a >= 0 And b <= 0 And rho >= 0 Then
        CBND = CND(b) - CBND(-a, b, -rho)
    ElseIf a >= 0 And b >= 0 And rho <= 0 Then
        CBND = CND(a) + CND(b) - 1 + CBND(-a, -b, rho)
    ElseIf a * b * rho > 0 Then
        rho1 = (rho * a - b) * Sgn(a) / Sqr(a ^ 2 - 2 * rho * a * b + b ^ 2)
        rho2 = (rho * b - a) * Sgn(b) / Sqr(a ^ 2 - 2 * rho * a * b + b ^ 2)
        delta = (1 - Sgn(a) * Sgn(b)) / 4
        CBND = CBND(a, 0, rho1) + CBND(b, 0, rho2) - delta
    End If
End Function

