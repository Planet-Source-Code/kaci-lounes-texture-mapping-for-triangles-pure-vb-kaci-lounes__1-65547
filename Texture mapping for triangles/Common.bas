Attribute VB_Name = "Common"

' MODULE NAME: Filtering.BAS
' ==========================
'
' Module for common functions (interpolation functions...etc)

Option Explicit
Function BaryInterpolateLinear(X1, Y1, Val1, X2, Y2, Val2, X3, Y3, Val3, PX, PY) As Single

 'Linear triangle interpolation
 '=============================

 Dim D!, U!, V!, W!

 D = 1 / (((X2 - X1) * (Y3 - Y1)) - ((Y2 - Y1) * (X3 - X1)))

 U = (((X2 - PX) * (Y3 - PY)) - ((Y2 - PY) * (X3 - PX))) * D
 V = (((X3 - PX) * (Y1 - PY)) - ((Y3 - PY) * (X1 - PX))) * D
 W = 1 - (U + V)

 BaryInterpolateLinear = (U * Val1) + (V * Val2) + (W * Val3)

End Function
Function BaryInterpolatePerspective(X1, Y1, Z1, Val1, X2, Y2, Z2, Val2, X3, Y3, Z3, Val3, PX, PY) As Single

 'Perspective (hyperbolic) triangle interpolation
 '===============================================

 Dim PX1!, PY1!, PX2!, PY2!, PX3!, PY3!, XP!, YP!, ZP!

 '1- Linearly interpolate the Z coordinate for the input point:
 ZP = 1 / BaryInterpolateLinear(X1, Y1, Z1, X2, Y2, Z2, X3, Y3, Z3, PX, PY)

 '2- Project the input point:
 XP = (PX * ZP): YP = (PY * ZP)

 '3- Project the triangle points:
 PX1 = (X1 / Z1): PY1 = (Y1 / Z1)
 PX2 = (X2 / Z2): PY2 = (Y2 / Z2)
 PX3 = (X3 / Z3): PY3 = (Y3 / Z3)

 'And linearly interpolate the value of the
 ' projected point in the projected triangle !!
 '
 '  This is my algorithm for perspective correction !!

 BaryInterpolatePerspective = BaryInterpolateLinear(PX1, PY1, Val1, PX2, PY2, Val2, PX3, PY3, Val3, XP, YP)

End Function
Function IsInsideTriangle(X1, Y1, X2, Y2, X3, Y3, PX, PY) As Boolean

 ' FUNCTION : IsInsideTriangle
 ' ===========================
 '
 ' RETURNED VALUE: Boolean
 '
 ' Check if a 2D point is inside a 2D triangle.

 Dim CRZ1!, CRZ2!, CRZ3!

 CRZ1 = (((X2 - PX) * (Y3 - PY)) - ((Y2 - PY) * (X3 - PX)))
 CRZ2 = (((X1 - PX) * (Y2 - PY)) - ((Y1 - PY) * (X2 - PX)))
 CRZ3 = (((X3 - PX) * (Y1 - PY)) - ((Y3 - PY) * (X1 - PX)))

 'The point is inside the triangle
 ' if the vars (CRZ1, CRZ2 & CRZ3)
 '  has the same sign:
 If ((CRZ1 > 0) And (CRZ2 > 0) And (CRZ3 > 0)) Or _
    ((CRZ1 < 0) And (CRZ2 < 0) And (CRZ3 < 0)) Then IsInsideTriangle = True

End Function
