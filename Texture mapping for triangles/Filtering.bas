Attribute VB_Name = "Filtering"

' MODULE NAME: Filtering.BAS
' ==========================
'
' Module for filtering, we can include this module
' in any project, just give a picturebox (or any
' bitmap datas, but i prefer picturebox for ActiveX),
' and the U,V values (float), these values are the
' coordinates for the current texel position in the
' source picturebox, note that we need the fractional
' part of these numbers for filtering.
'
' There are 7 kernels filters:
'
'  - Bilinear 1(fast)
'  - Bilinear 2
'  - Bell
'  - Gaussian
'  - Bicubic B spline
'  - Bicubic BC spline
'  - Bicubic cardinal spline
'
' As this, filtering is very simple ! for exmple,
' you can use this for 2D resampling, raycast
' engines, and any other transforms.

Option Explicit
Function Gaussian(PS As PictureBox, U, V, KernelSize%) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim R%, G%, B%, RR%, GG%, BB%
 Dim M&, N&, R1!, R2!

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 For M = -1 To 2
  R1 = FunctionGaussian(M - FracY, KernelSize)
  For N = -1 To 2
   R2 = FunctionGaussian(FracX - N, KernelSize)

   RR = PS.Point(IntX + N, IntY + M) And 255
   GG = (PS.Point(IntX + N, IntY + M) And 65280) / 256
   BB = (PS.Point(IntX + N, IntY + M) And 16711680) / 65536

   R = R + ((RR * R1) * R2)
   G = G + ((GG * R1) * R2)
   B = B + ((BB * R1) * R2)

  Next N
 Next M

 If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
 If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
 If (B < 0) Then B = 0 Else If (B > 255) Then B = 255

 Gaussian = RGB(R, G, B)

End Function
Function FunctionGaussian(X As Single, KernelSize%) As Single

 Dim O!

 If (Abs(X) < KernelSize) Then
  O = (KernelSize / 3.141593)
  FunctionGaussian = Exp((-X * X) / ((O * O) * 2)) * (0.3989423 / O)
 End If

End Function
Function FunctionBartlettLinear(X As Single, KernelSize%) As Single

 X = Abs(X)

 If (X < KernelSize) Then
  FunctionBartlettLinear = ((1 - Abs(X)) / KernelSize) / KernelSize
 End If

End Function
Function Bilinear2(PS As PictureBox, U, V, KernelSize%) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim R%, G%, B%, RR%, GG%, BB%
 Dim M&, N&, R1!, R2!

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 For M = -1 To 2
  R1 = FunctionBartlettLinear(M - FracY, KernelSize)
  For N = -1 To 2
   R2 = FunctionBartlettLinear(FracX - N, KernelSize)

   RR = PS.Point(IntX + N, IntY + M) And 255
   GG = (PS.Point(IntX + N, IntY + M) And 65280) / 256
   BB = (PS.Point(IntX + N, IntY + M) And 16711680) / 65536

   R = R + ((RR * R1) * R2)
   G = G + ((GG * R1) * R2)
   B = B + ((BB * R1) * R2)

  Next N
 Next M

 If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
 If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
 If (B < 0) Then B = 0 Else If (B > 255) Then B = 255

 Bilinear2 = RGB(R, G, B)

End Function
Function BicubicCardinal(PS As PictureBox, U, V, CubicA!) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim R%, G%, B%, RR%, GG%, BB%
 Dim M&, N&, R1!, R2!

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 For M = -1 To 2
  R1 = FunctionCardinalCubicSpline(M - FracY, CubicA)
  For N = -1 To 2
   R2 = FunctionCardinalCubicSpline(FracX - N, CubicA)

   RR = PS.Point(IntX + N, IntY + M) And 255
   GG = (PS.Point(IntX + N, IntY + M) And 65280) / 256
   BB = (PS.Point(IntX + N, IntY + M) And 16711680) / 65536

   R = R + ((RR * R1) * R2)
   G = G + ((GG * R1) * R2)
   B = B + ((BB * R1) * R2)

  Next N
 Next M

 If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
 If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
 If (B < 0) Then B = 0 Else If (B > 255) Then B = 255

 BicubicCardinal = RGB(R, G, B)

End Function
Function FunctionCardinalCubicSpline(X As Single, CubicA As Single) As Single

 X = Abs(X)

 If X < 1 Then
  FunctionCardinalCubicSpline = (((CubicA + 2) * (X ^ 3)) - ((CubicA + 3) * (X ^ 2))) + 1
 ElseIf X < 2 Then
  FunctionCardinalCubicSpline = (((CubicA * (X ^ 3)) - ((5 * CubicA) * (X ^ 2))) + ((8 * CubicA) * X)) - (4 * CubicA)
 End If

End Function
Function BicubicBCSpline(PS As PictureBox, U, V, CubicB!, CubicC!) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim R%, G%, B%, RR%, GG%, BB%
 Dim M&, N&, R1!, R2!

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 For M = -1 To 2
  R1 = FunctionCubicBCSpline(M - FracY, CubicB, CubicC)
  For N = -1 To 2
   R2 = FunctionCubicBCSpline(FracX - N, CubicB, CubicC)

   RR = PS.Point(IntX + N, IntY + M) And 255
   GG = (PS.Point(IntX + N, IntY + M) And 65280) / 256
   BB = (PS.Point(IntX + N, IntY + M) And 16711680) / 65536

   R = R + ((RR * R1) * R2)
   G = G + ((GG * R1) * R2)
   B = B + ((BB * R1) * R2)

  Next N
 Next M

 If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
 If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
 If (B < 0) Then B = 0 Else If (B > 255) Then B = 255

 BicubicBCSpline = RGB(R, G, B)

End Function
Function FunctionCubicBCSpline(X As Single, CubicB As Single, CubicC As Single) As Single

 X = Abs(X)

 If (X < 1) Then
  FunctionCubicBCSpline = ((12 - (9 * CubicB)) - (6 * CubicC)) * (X ^ 3)
  FunctionCubicBCSpline = FunctionCubicBCSpline + (((-18 + (12 * CubicB)) + (6 * CubicC)) * (X ^ 2))
  FunctionCubicBCSpline = ((FunctionCubicBCSpline + 6) - (2 * CubicB)) * 0.1666666
 ElseIf (X < 2) Then
  FunctionCubicBCSpline = (-CubicB - (6 * CubicC)) * (X ^ 3)
  FunctionCubicBCSpline = (FunctionCubicBCSpline + ((6 * CubicB) + (30 * CubicC)) * (X ^ 2))
  FunctionCubicBCSpline = (FunctionCubicBCSpline + ((-12 * CubicB) - (48 * CubicC)) * X)
  FunctionCubicBCSpline = (FunctionCubicBCSpline + ((8 * CubicB) + (24 * CubicC))) * 0.1666666
 End If

End Function
Function BicubicBSpline(PS As PictureBox, U, V) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim R%, G%, B%, RR%, GG%, BB%
 Dim M&, N&, R1!, R2!

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 For M = -1 To 2
  R1 = FunctionCubicBSpline(M - FracY)
  For N = -1 To 2
   R2 = FunctionCubicBSpline(FracX - N)

   RR = PS.Point(IntX + N, IntY + M) And 255
   GG = (PS.Point(IntX + N, IntY + M) And 65280) / 256
   BB = (PS.Point(IntX + N, IntY + M) And 16711680) / 65536

   R = R + ((RR * R1) * R2)
   G = G + ((GG * R1) * R2)
   B = B + ((BB * R1) * R2)

  Next N
 Next M

 If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
 If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
 If (B < 0) Then B = 0 Else If (B > 255) Then B = 255

 BicubicBSpline = RGB(R, G, B)

End Function
Function FunctionCubicBSpline(X As Single) As Single

 Dim A!, B!, C!, D!, Tmp!

 If (X < 2) Then
  Tmp = (X + 2): If (Tmp > 0) Then A = (Tmp ^ 3)
  Tmp = (X + 1): If (Tmp > 0) Then B = ((Tmp ^ 3) * 4)
  If (X > 0) Then C = ((X ^ 3) * 6)
  Tmp = (X - 1): If (Tmp > 0) Then D = ((Tmp ^ 3) * 4)
  FunctionCubicBSpline = (((A - B) + (C - D)) * 0.1666666)
 End If

End Function
Function FunctionBell(X As Single) As Single

 X = Abs(X)

 If (X < 0.5) Then
  FunctionBell = 0.75 - (X ^ 2)
 ElseIf (X < 1.5) Then
  FunctionBell = ((X - 1.5) ^ 2) * 0.5
 End If

End Function
Function Bilinear(PS As PictureBox, U, V) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim IR1%, IG1%, IB1%, IR2%, IG2%, IB2%
 Dim R%, R1%, R2%, R3%, R4%
 Dim G%, G1%, G2%, G3%, G4%
 Dim B%, B1%, B2%, B3%, B4%

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 R1 = PS.Point(IntX, IntY) And 255
 G1 = (PS.Point(IntX, IntY) And 65280) / 256
 B1 = (PS.Point(IntX, IntY) And 16711680) / 65536

 R2 = PS.Point(IntX + 1, IntY) And 255
 G2 = (PS.Point(IntX + 1, IntY) And 65280) / 256
 B2 = (PS.Point(IntX + 1, IntY) And 16711680) / 65536

 R3 = PS.Point(IntX, IntY + 1) And 255
 G3 = (PS.Point(IntX, IntY + 1) And 65280) / 256
 B3 = (PS.Point(IntX, IntY + 1) And 16711680) / 65536

 R4 = PS.Point(IntX + 1, IntY + 1) And 255
 G4 = (PS.Point(IntX + 1, IntY + 1) And 65280) / 256
 B4 = (PS.Point(IntX + 1, IntY + 1) And 16711680) / 65536

 IR1 = (FracY * R3) + ((1 - FracY) * R1)
 IG1 = (FracY * G3) + ((1 - FracY) * G1)
 IB1 = (FracY * B3) + ((1 - FracY) * B1)

 IR2 = (FracY * R4) + ((1 - FracY) * R2)
 IG2 = (FracY * G4) + ((1 - FracY) * G2)
 IB2 = (FracY * B4) + ((1 - FracY) * B2)

 R = (FracX * IR2) + ((1 - FracX) * IR1)
 G = (FracX * IG2) + ((1 - FracX) * IG1)
 B = (FracX * IB2) + ((1 - FracX) * IB1)

 Bilinear = RGB(R, G, B)

End Function
Function Bell(PS As PictureBox, U, V) As Long

 Dim IntX%, IntY%, FracX!, FracY!
 Dim R%, G%, B%, RR%, GG%, BB%
 Dim M&, N&, R1!, R2!

 IntX = Fix(U): FracX = (U - IntX)
 IntY = Fix(V): FracY = (V - IntY)

 For M = -1 To 2
  R1 = FunctionBell(M - FracY)
  For N = -1 To 2
   R2 = FunctionBell(FracX - N)

   RR = PS.Point(IntX + N, IntY + M) And 255
   GG = (PS.Point(IntX + N, IntY + M) And 65280) / 256
   BB = (PS.Point(IntX + N, IntY + M) And 16711680) / 65536

   R = R + ((RR * R1) * R2)
   G = G + ((GG * R1) * R2)
   B = B + ((BB * R1) * R2)

  Next N
 Next M

 If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
 If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
 If (B < 0) Then B = 0 Else If (B > 255) Then B = 255

 Bell = RGB(R, G, B)

End Function
