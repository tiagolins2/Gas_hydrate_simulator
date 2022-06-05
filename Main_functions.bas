Attribute VB_Name = "Module1"
Sub hfcalc()
'this subroutine reads all important variables and constants, calculates chemical potential
'and returns the equilibrium pressure at a fixed temperature and an initial guess for pressure


'reads main variables (T,P, and ideal gas constant) that will be used in the program
R = 8.314
T = Range("b3").Value
P = Range("b5").Value

'reading the Peng-Robinson parameters a and b from the spreadsheet
Dim a(1 To 2)
Dim b(1 To 2)
a(1) = Range("e7").Value
a(2) = Range("f7").Value
b(1) = Range("e6").Value
b(2) = Range("f6").Value

'eps is the error between the chemical potentials
eps = 1
kount = 1


'reads in the values for the parameters that are used to estimate the
'chemical potentials of the hydrate and of the liquid phases,
'based on the value of temperature
If T >= 273.15 Then
Dho = Range("f10").Value
Cpo = Range("g10").Value
be = Range("h10").Value
Dmu = Range("e10").Value
Dv = Range("i10").Value
vms = Range("f13").Value
vml = Range("f14").Value
Aml = Range("f18").Value
Ams = Range("f17").Value
Bml = Range("g18").Value
Bms = Range("g17").Value

End If

If T < 273.15 Then
Dho = Range("f11").Value
Cpo = Range("g11").Value
be = Range("h11").Value
Dmu = Range("e11").Value
Dv = Range("i11").Value
vms = Range("f13").Value
vml = Range("f14").Value
Aml = Range("f18").Value
Ams = Range("f17").Value
Bml = Range("g18").Value
Bms = Range("g17").Value
End If


'start of the loop where pressure is updated until eps is close to zero
Do Until eps <= 0.000001


'calculates the liquid and vapour molecular volumes (using H2O data for liquid and CH4 data for vapour)
vv = vapourvolume(R, a(1), b(1), T, P)


'calculates fugacity of the vapour that is used to
fugv = fugacity(a(1), b(1), P, T, vv, R)
Range("k5").Value = fugv

'DmuMTL is the difference in chemical potential of the hypothetical empty hydrate and the liquid water

DmuMTL = DmuMT_L(T, Dho, R, Cpo, be, Dmu, Dv, P)

'DmuMTH is the difference in chemical potential of the hypothetical empty hydrate and the hydrate
DmuMTH = DmuMT_H(vms, vml, Aml, Ams, Bml, Bms, fugv, T, P)

'P is found by using a function that performs the secant's method
P = DmuMT_L_H(T, Dho, R, Cpo, be, Dmu, Dv, vms, vml, Aml, Ams, Bml, Bms, fugv, P)

'pressure, the number of iterations and the error (eps) are reported back into the spreadsheet
Range("k3").Value = P
eps = (DmuMTL / DmuMTH - 1) ^ 2
Range("k4").Value = kount
Range("k2").Value = eps
kount = kount + 1

Loop

End Sub


Sub hfcalcmix()
'this function is very similar to hfcalc, but uses the fugacity found from flash calculation
R = 8.314
T = Range("b3").Value
P = Range("b5").Value

Dim a(1 To 2)
Dim b(1 To 2)
a(1) = Range("e7").Value
a(2) = Range("f7").Value
b(1) = Range("e6").Value
b(2) = Range("f6").Value
eps = 1
kount = 1
Range("k3").Value = Range("b4").Value


If T >= 273.15 Then
Dho = Range("f10").Value
Cpo = Range("g10").Value
be = Range("h10").Value
Dmu = Range("e10").Value
Dv = Range("i10").Value
vms = Range("f13").Value
vml = Range("f14").Value
Aml = Range("f18").Value
Ams = Range("f17").Value
Bml = Range("g18").Value
Bms = Range("g17").Value

End If

If T < 273.15 Then
Dho = Range("f11").Value
Cpo = Range("g11").Value
be = Range("h11").Value
Dmu = Range("e11").Value
Dv = Range("i11").Value
vms = Range("f13").Value
vml = Range("f14").Value
Aml = Range("f18").Value
Ams = Range("f17").Value
Bml = Range("g18").Value
Bms = Range("g17").Value
End If


Do Until eps <= 0.000001


mixture

fugl = Range("b14").Value
fugv = Range("b15").Value


DmuMTL = DmuMT_L(T, Dho, R, Cpo, be, Dmu, Dv, P)
DmuMTH = DmuMT_H(vms, vml, Aml, Ams, Bml, Bms, fugv, T, P)


P = DmuMT_L_H(T, Dho, R, Cpo, be, Dmu, Dv, vms, vml, Aml, Ams, Bml, Bms, fugv, P)


Range("k3").Value = P
eps = (DmuMTL / DmuMTH - 1) ^ 2
Range("k4").Value = kount
Range("k2").Value = eps
kount = kount + 1

Loop

End Sub

Function fugacitymix(b_i, b_mix, aik, a_mix, Z, BB, AA, y_in, P)

sum_a = 2 * aik
func1 = -Log(Z - BB) + (Z - 1) * b_i / b_mix
func2 = -AA / (2 * 2 ^ 0.5 * BB) * (sum_a / a_mix - b_i / b_mix)
func3 = Log((Z + (2 ^ 0.5 + 1) * BB) / (Z - (2 ^ 0.5 - 1) * BB))
fugacitymix = Exp(func1 + func2 * func3) * y_in


End Function


Function apure(T, Tc, omega, Rg, Pc)
'calculates the value for a for pure component given T(temperature of the flash), Tc(critical temperature, omega(accentric factor)
'R(ideal gas constant), and Pc(critical pressure)
alpha = (1 + (1 - (T / Tc) ^ (1 / 2)) * (0.37 + 1.54 * omega - 0.27 * omega ^ 2)) ^ (2)
apure = 0.45724 * Rg ^ 2 * Tc ^ 2 * alpha / Pc
End Function



Function liquidvolume(R, a, b, T, P)
'calculates liquid volume given a, b, T, and P



z_i = 0.001
AA = a * P / (R ^ 2 * T ^ 2)
BB = b * P / (R * T)
eps = 1

Do Until eps < 0.0001
Z = z_i - z_zero(z_i, AA, BB) / dz(z_i, AA, BB)
eps = Abs(Z - z_i)
z_i = Z
Loop

liquidvolume = Z * R * T / P


End Function


Function vapourvolume(R, a, b, T, P)
'calculates vapour volume given a, b, T, and P


z_i = 0.99
AA = a * P / (R ^ 2 * T ^ 2)
BB = b * P / (R * T)
eps = 1

'function estimates a value for compressibility and uses it to find volume
Do Until eps < 0.0001
Z = z_i - z_zero(z_i, AA, BB) / dz(z_i, AA, BB)
eps = Abs(Z - z_i)
z_i = Z


Loop

vapourvolume = Z * R * T / P


End Function


Function z_zero(Z, AA, BB)
'returns the
z_zero = Z ^ 3 + (BB - 1) * Z ^ 2 + (AA - 2 * BB - 3 * BB ^ 2) * Z + (BB ^ 3 + BB ^ 2 - AA * BB)
End Function
Function dz(Z, AA, BB)
dz = 3 * Z ^ 2 + (BB - 1) * 2 * Z + (AA - 2 * BB - 3 * BB ^ 2)
End Function


Function fugacity(a, b, P, T, v, R)
'calculates fugacity given a,b, pressure, temperature, molar volume and gas constant
Z = P * v / (R * T)
AA = a * P / (R * T) ^ 2
BB = b * P / (R * T)
fugcoeff = Exp((Z - 1) - Log(Z - BB) - AA / (2 * 2 ^ 0.5 * BB) * Log((Z + (1 + 2 ^ 0.5) * BB) / (Z + (1 - 2 ^ 0.5) * BB)))
fugacity = fugcoeff

End Function


Function DmuMT_L(T, Dho, R, Cpo, beta, Dmu, Dv, P)
'returns the difference in chemical potential in the empty lattice and and liquid phase

'calculating the temperature correction
Tcorrection = (1 / 273.15 - 1 / T) * (Dho / R - Cpo * 273.15 / R + beta * 273.15 ^ 2 / (2 * R)) + Log(T / 273.15) * (Cpo / R - beta * 273.15 / R) + beta / (2 * R) * (T - 273.15)
DmuMT_L = (Dmu / (R * 273.15) + Dv * (P - 0) / (R * T) - Tcorrection)


End Function

Function DmuMT_H(vms, vml, Aml, Ams, Bml, Bms, fug, T, P)
'returns the difference in chemical potential in the empty lattice and and hydrate phase

Cms = Ams / T * Exp(Bms / T) * 1 / 101325
Cml = Aml / T * Exp(Bml / T) * 1 / 101325
DmuMT_H = vms * Log(1 + Cms * fug * P) + vml * Log(1 + Cml * fug * P)


End Function

Function DmuMT_L_H(T, Dho, R, Cpo, beta, Dmu, Dv, vms, vml, Aml, Ams, Bml, Bms, fug, P)
'returns the pressure at which difference in chemical potential between the liquid phase
'and the hydrate phase is zero, holding all other variables constant

P_ii = P
P_i = P - 1
eps = 1
i = 1
Do Until i > 20 Or eps < 0.0000001
h_ii = (DmuMT_L(T, Dho, R, Cpo, beta, Dmu, Dv, P_ii) - DmuMT_H(vms, vml, Aml, Ams, Bml, Bms, fug, T, P_ii))
h_i = (DmuMT_L(T, Dho, R, Cpo, beta, Dmu, Dv, P_i) - DmuMT_H(vms, vml, Aml, Ams, Bml, Bms, fug, T, P_i))
P = P_ii - h_ii * (P_ii - P_i) / (h_ii - h_i)
eps = Abs(P - P_ii)
P_i = P_ii
P_ii = P



i = i + 1
Loop
DmuMT_L_H = P



End Function

Sub populate()
cltable
i = 0
b = 0

Do Until i > Range("n7").Value - 1
Range("m" & (i + 9)).Value = 20 * i / (Range("n7").Value - 1) + 260
Range("b3").Value = Range("m" & (i + 9)).Value


hfcalc

Range("b4").Value = Range("k3").Value
Range("n" & (i + 9)).Value = Range("k3").Value
Range("o" & (i + 9)).Value = Range("k4").Value
Range("p" & (i + 9)).Value = Range("k5").Value * Range("k3").Value
Range("q" & (i + 9)).Value = Range("k2").Value
i = i + 1
Loop

End Sub

Sub populatemix()
cltable
i = 0
b = 0

Do Until i > Range("n7").Value - 1
Range("m" & (i + 9)).Value = 20 * i / (Range("n7").Value - 1) + 260
Range("b3").Value = Range("m" & (i + 9)).Value

hfcalcmix

Range("b4").Value = Range("k3").Value
Range("n" & (i + 9)).Value = Range("k3").Value
Range("o" & (i + 9)).Value = Range("k4").Value
Range("p" & (i + 9)).Value = Range("k5").Value * Range("k3").Value
Range("q" & (i + 9)).Value = Range("k2").Value
i = i + 1
Loop

End Sub


