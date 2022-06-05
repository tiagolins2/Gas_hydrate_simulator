Attribute VB_Name = "Module2"
Sub mixture()
z1 = Range("b8").Value
z2 = Range("b9").Value
P = Range("k3").Value
T = Range("b3").Value
Rg = 8.314
omega1 = Range("e5").Value
omega2 = Range("f5").Value
Pc1 = Range("e4").Value
Pc2 = Range("f4").Value
Tc1 = Range("e3").Value
Tc2 = Range("f3").Value
kk1 = 100
kk2 = 0.001

b1 = Range("e6").Value
b2 = Range("f6").Value
a1 = Range("e7").Value
a2 = Range("f7").Value
be = beta(z1, z2, kk1, kk2)

F = 1

kount = 1


Do Until kount > 50
be = beta(z1, z2, kk1, kk2)
v = be * F
L = F - v
x1 = z1 * F / (L + kk1 * v)
x2 = z2 * F / (L + kk2 * v)
y1 = x1 * kk1
y2 = x2 * kk2

b_vmix = b1 * y1 + b2 * y2
b_lmix = b1 * x1 + b2 * x2
a_vmix = y1 * y2 * (a1 * a2) ^ (1 / 2) + y1 * y1 * (a1 * a1) ^ (1 / 2) + y2 * y2 * (a2 * a2) ^ (1 / 2)
a_lmix = x1 * x2 * (a1 * a2) ^ (1 / 2) + x1 * x1 * (a1 * a1) ^ (1 / 2) + x2 * x2 * (a2 * a2) ^ (1 / 2)
Zl = liquidz(Rg, a_lmix, b_lmix, T, P)
Zv = vapourz(Rg, a_vmix, b_vmix, T, P)
AAv = a_vmix * P / (Rg ^ 2 * T ^ 2)
AAl = a_lmix * P / (Rg ^ 2 * T ^ 2)
BBv = b_vmix * P / (Rg * T)
BBl = b_lmix * P / (Rg * T)

aikL1 = (a1 * a1) ^ 0.5 * x1 + (a1 * a2) ^ 0.5 * x2
aikV1 = (a1 * a1) ^ 0.5 * y1 + (a1 * a2) ^ 0.5 * y2
aikL2 = (a1 * a2) ^ 0.5 * x1 + (a2 * a2) ^ 0.5 * x2
aikV2 = (a1 * a2) ^ 0.5 * y1 + (a2 * a2) ^ 0.5 * y2

FL1 = comp_fg(x1, P, b1, b_lmix, Zl, BBl, AAl, aikL1, a_lmix, T)
FV1 = comp_fg(y1, P, b1, b_vmix, Zv, BBv, AAv, aikV1, a_vmix, T)
FL2 = comp_fg(x2, P, b2, b_lmix, Zl, BBl, AAl, aikL2, a_lmix, T)
FV2 = comp_fg(y2, P, b2, b_vmix, Zv, BBv, AAv, aikV2, a_vmix, T)

theta = y1 * Log(FV1 / FL1) + y2 * Log(FV2 / FL2)

kk1 = kk1 * (FL1 / FV1)
kk2 = kk2 * (FL2 / FV2)
Range("b11").Value = kount
kount = kount + 1
Loop
Range("b10").Value = y1
Range("b11").Value = y2
Range("b12").Value = x1
Range("b13").Value = x2

Range("b14").Value = FL1
Range("b15").Value = FV1
Range("b16").Value = FL2
Range("b17").Value = FV2



End Sub

Function ideal_ki(T, P, Tc, Pc, omega)
'returns the value for V/F using Wilson's equation
ideal_ki = Pc / P * Exp(5.373 * (1 + omega) * (1 - Tc / T))
End Function

Function comp_fg(yin, P, b_in, b, Z, BB, AA, sumxaik, a, T)
AA = a * P / (8.314 ^ 2 * T ^ 2)
BB = b * P / (8.314 * T)
func1 = AA / (2 * 2 ^ 0.5 * BB)
func2 = (2 * sumxaik / a - b_in / b)
func3 = Log((Z + 2.414 * BB) / (Z - 0.414 * BB))
func4 = b_in / b * (Z - 1) - Log(Z - BB)
comp_fg = Exp(-func1 * func2 * func3 + func4) * yin

End Function


Function beta(z1, z2, kk1, kk2)

beta_i = 0.2
beta_ii = 0.21
eps = 1
be = 0
kount = 1
Do While eps > 0.001
be = beta_ii - betafun(z1, z2, kk1, kk2, beta_ii) * (beta_ii - beta_i) / (betafun(z1, z2, kk1, kk2, beta_ii) - betafun(z1, z2, kk1, kk2, beta_i))
eps = Abs(be - beta_ii)
kount = kount + 1
beta_i = beta_ii
beta_ii = be
Loop

beta = be

End Function
Function betafun(z1, z2, kk1, kk2, be)
b1 = z1 * (kk1 - 1) / (1 + be * (kk1 - 1))
b2 = z2 * (kk2 - 1) / (1 + be * (kk2 - 1))
betafun = b1 + b2
End Function


Function liquidz(R, a, b, T, P)
'calculates liquid volume given a, b, T, and P



z_i = 0.01
AA = a * P / (R ^ 2 * T ^ 2)
BB = b * P / (R * T)
eps = 1

Do Until eps < 0.0000001
Z = z_i - z_zero(z_i, AA, BB) / dz(z_i, AA, BB)
eps = Abs(Z - z_i)
z_i = Z
Loop

liquidz = Z


End Function


Function vapourz(R, a, b, T, P)
'calculates vapour volume given a, b, T, and P


z_i = 0.8

AA = a * P / (R ^ 2 * T ^ 2)
BB = b * P / (R * T)
eps = 1

'function estimates a value for compressibility and uses it to find volume
Do Until eps < 0.0000001
Z = z_i - z_zero(z_i, AA, BB) / dz(z_i, AA, BB)
eps = Abs(Z - z_i)
z_i = Z


Loop

vapourz = Z


End Function
