VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Texture mapping for triangles, By KACI Lounes"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   632
   ScaleMode       =   0  'User
   ScaleWidth      =   856.595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   12000
      TabIndex        =   26
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Rendering"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9360
      TabIndex        =   25
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13320
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Render triangle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6720
      TabIndex        =   23
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Frame FOptions 
      Caption         =   "Rendering options :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   6495
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   18
         Text            =   "1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   17
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   11
         Text            =   "1,25"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "3,5"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "6,5"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Main.frx":0000
         Left            =   1560
         List            =   "Main.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   4695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Main.frx":00AD
         Left            =   2040
         List            =   "Main.frx":00B7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Krnl.Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4800
         TabIndex        =   22
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CubicC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   21
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CubicB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CubicA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Z3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   14
         Top             =   840
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Z2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Z1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kernels filter :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interpolation type :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.Frame FTex 
      Caption         =   "Input texture : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   7320
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox PTex 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1920
         Left            =   120
         Picture         =   "Main.frx":00F1
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   3
         Top             =   240
         Width           =   1920
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   6360
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00C0C0C0&
            DrawMode        =   6  'Mask Pen Not
            Visible         =   0   'False
            X1              =   0
            X2              =   216
            Y1              =   64
            Y2              =   24
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00C0C0C0&
            DrawMode        =   6  'Mask Pen Not
            Visible         =   0   'False
            X1              =   0
            X2              =   216
            Y1              =   48
            Y2              =   8
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C0C0C0&
            DrawMode        =   6  'Mask Pen Not
            Visible         =   0   'False
            X1              =   0
            X2              =   216
            Y1              =   40
            Y2              =   0
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H000000FF&
            Height          =   135
            Left            =   1440
            Shape           =   3  'Circle
            Top             =   600
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H000000FF&
            Height          =   135
            Left            =   120
            Shape           =   3  'Circle
            Top             =   240
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H000000FF&
            Height          =   135
            Left            =   480
            Shape           =   3  'Circle
            Top             =   1560
            Visible         =   0   'False
            Width           =   135
         End
      End
   End
   Begin VB.Frame FRender 
      Caption         =   "Rendering view :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox PRender 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Left            =   120
         ScaleHeight     =   433
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   1
         Top             =   240
         Width           =   6855
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            Visible         =   0   'False
            X1              =   192
            X2              =   112
            Y1              =   16
            Y2              =   280
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            Visible         =   0   'False
            X1              =   8
            X2              =   112
            Y1              =   56
            Y2              =   304
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            Visible         =   0   'False
            X1              =   8
            X2              =   224
            Y1              =   48
            Y2              =   8
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H000000FF&
            Height          =   135
            Left            =   4560
            Shape           =   3  'Circle
            Top             =   720
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H000000FF&
            Height          =   135
            Left            =   960
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            Height          =   135
            Left            =   0
            Shape           =   3  'Circle
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################
'##                  Author: Mr KACI Lounes                        ##
'##      Texture Mapping for triangles in *Pure* VB Code!          ##
'##    Compile for more speed ! Mail me at KLKEANO@CARAMAIL.COM    ##
'##               Copyright © 2006 - KACI Lounes                   ##
'####################################################################

'SOME INTERNAL VARS
'==================

Dim PDMouseID%, PTexMouseID%
Dim Clicked1 As Boolean, Clicked2 As Boolean
Dim IsRendering As Boolean
Dim RenderOption%, FilterOption%
Dim X1, Y1, X2, Y2, X3, Y3, U1, V1, U2, V2, U3, V3
Sub Render()

 Dim CurX, CurY, CurU, CurV

 For CurY = 0 To PRender.ScaleHeight
  For CurX = 0 To PRender.ScaleWidth

   If IsInsideTriangle(X1, Y1, X2, Y2, X3, Y3, CurX, CurY) = True Then

    If RenderOption = 1 Then
     CurU = BaryInterpolateLinear(X1, Y1, U1, X2, Y2, U2, X3, Y3, U3, CurX, CurY)
     CurV = BaryInterpolateLinear(X1, Y1, V1, X2, Y2, V2, X3, Y3, V3, CurX, CurY)
    ElseIf RenderOption = 2 Then
     CurU = BaryInterpolatePerspective(X1, Y1, Text1.Text, U1, X2, Y2, Text2.Text, U2, X3, Y3, Text3.Text, U3, CurX, CurY)
     CurV = BaryInterpolatePerspective(X1, Y1, Text1.Text, V1, X2, Y2, Text2.Text, V2, X3, Y3, Text3.Text, V3, CurX, CurY)
    End If

    Select Case FilterOption
     Case 1: PRender.PSet (Fix(CurX), Fix(CurY)), PTex.Point(CurU, CurV)
     Case 2: PRender.PSet (Fix(CurX), Fix(CurY)), Bilinear(PTex, CurU, CurV)
     Case 3: PRender.PSet (Fix(CurX), Fix(CurY)), Bilinear2(PTex, CurU, CurV, Text7.Text)
     Case 4: PRender.PSet (Fix(CurX), Fix(CurY)), Bell(PTex, CurU, CurV)
     Case 5: PRender.PSet (Fix(CurX), Fix(CurY)), Gaussian(PTex, CurU, CurV, Text7.Text)
     Case 6: PRender.PSet (Fix(CurX), Fix(CurY)), BicubicBSpline(PTex, CurU, CurV)
     Case 7: PRender.PSet (Fix(CurX), Fix(CurY)), BicubicBCSpline(PTex, CurU, CurV, Text5.Text, Text6.Text)
     Case 8: PRender.PSet (Fix(CurX), Fix(CurY)), BicubicCardinal(PTex, CurU, CurV, Text4.Text)
    End Select

   End If

  Next CurX
  DoEvents: If IsRendering = False Then Exit For
 Next CurY

End Sub
Private Sub Combo1_Click()

 If Combo1.Text = "Affine mapping (linear)" Then

  RenderOption = 1
  Text1.Enabled = False: Text2.Enabled = False: Text3.Enabled = False

 ElseIf Combo1.Text = "Perspective-correct mapping" Then

  RenderOption = 2
  Text1.Enabled = True: Text2.Enabled = True: Text3.Enabled = True

 End If

End Sub

Private Sub Combo2_Click()

 Select Case Combo2.Text

  Case "Nearest Nighbor (no filtering)":
       FilterOption = 1
       Text4.Enabled = False
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = False

  Case "Bilinear 1 (fast)":
       FilterOption = 2
       Text4.Enabled = False
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = False

  Case "Bilinear 2":
       FilterOption = 3
       Text4.Enabled = False
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = True

  Case "Bell":
       FilterOption = 4
       Text4.Enabled = False
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = False

  Case "Gaussian":
       FilterOption = 5
       Text4.Enabled = False
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = True

  Case "BiCubic B Spline":
       FilterOption = 6
       Text4.Enabled = False
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = False

  Case "BiCubic BC Spline":
       FilterOption = 7
       Text4.Enabled = False
       Text5.Enabled = True
       Text6.Enabled = True
       Text7.Enabled = False

  Case "BiCubic Cardinal Spline":
       FilterOption = 8
       Text4.Enabled = True
       Text5.Enabled = False
       Text6.Enabled = False
       Text7.Enabled = False

 End Select

End Sub

Private Sub Command1_Click()

 If Clicked1 = False Then
  MsgBox "Please define the triangle corners in screen space !", vbInformation
  Exit Sub
 End If

 If Clicked2 = False Then
  MsgBox "Please define the triangle corners in texture space !", vbInformation
  Exit Sub
 End If

 If RenderOption = 2 Then
  If IsNumeric(Text1.Text) = False Then MsgBox "Please enter a valid number in Z1 field !", vbInformation: Exit Sub
  If IsNumeric(Text2.Text) = False Then MsgBox "Please enter a valid number in Z2 field !", vbInformation: Exit Sub
  If IsNumeric(Text3.Text) = False Then MsgBox "Please enter a valid number in Z3 field !", vbInformation: Exit Sub
 End If

 Select Case FilterOption

  Case 8:
         If IsNumeric(Text4.Text) = False Then MsgBox "Please enter a valid number in CubicA field !", vbInformation: Exit Sub
         If (Text4.Text < 0) Or (Text4.Text > 1) Then MsgBox "Please enter a number in range 0.....1 in CubicA field !", vbInformation: Exit Sub
  Case 7:
         If IsNumeric(Text5.Text) = False Then MsgBox "Please enter a valid number in CubicB field !", vbInformation: Exit Sub
         If IsNumeric(Text6.Text) = False Then MsgBox "Please enter a valid number in CubicC field !", vbInformation: Exit Sub
         If (Text5.Text < 0) Or (Text5.Text > 1) Then MsgBox "Please enter a number in range 0.....1 in CubicB field !", vbInformation: Exit Sub
         If (Text6.Text < 0) Or (Text6.Text > 1) Then MsgBox "Please enter a number in range 0.....1 in CubicC field !", vbInformation: Exit Sub
  Case 3, 5:
         If IsNumeric(Text7.Text) = False Then MsgBox "Please enter a valid number in Krnl.Size field !", vbInformation: Exit Sub
         If (Text7.Text < 1) Or (Text7.Text > 5) Then MsgBox "Please enter a number in range 1.....5 in Krnl.Size field !", vbInformation: Exit Sub
 End Select

 X1 = Shape1.Left: Y1 = Shape1.Top: U1 = Shape4.Left: V1 = Shape4.Top
 X2 = Shape2.Left: Y2 = Shape2.Top: U2 = Shape5.Left: V2 = Shape5.Top
 X3 = Shape3.Left: Y3 = Shape3.Top: U3 = Shape6.Left: V3 = Shape6.Top

 PRender.Cls

 IsRendering = True
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
 Shape4.Visible = False
 Shape5.Visible = False
 Shape6.Visible = False

 Render

 IsRendering = False
 Command1.Enabled = True
 Command2.Enabled = True
 Command3.Enabled = False
 Text1.Enabled = True
 Text2.Enabled = True
 Text3.Enabled = True
 Shape4.Visible = True
 Shape5.Visible = True
 Shape6.Visible = True

End Sub
Private Sub Command2_Click()

 On Error GoTo Error

 CommonDialog1.ShowOpen
 PTex.Picture = LoadPicture(CommonDialog1.FileName)
 FTex.Caption = "Input texture (" & CommonDialog1.FileName & ") : "

Error:

 If Err.Number <> 0 Then
  MsgBox "Enter a valid picture file !", vbCritical, "Bad file !"
  Exit Sub
 End If

End Sub

Private Sub Command3_Click()

 IsRendering = False

End Sub

Private Sub Command4_Click()

 Dim StrMsg As String

 Unload Me

 StrMsg = "Textue mapping for triangles, Pure VB ! & FREE, but give me a little credit !" & vbNewLine & vbNewLine & _
          "                        KACI Lounes - Mai 2006"

 MsgBox StrMsg, vbInformation, "By !"

 End

End Sub
Private Sub Form_Load()

 Combo1.Text = "Affine mapping (linear)"
 Combo2.Text = "Nearest Nighbor (no filtering)"

End Sub

Private Sub PRender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If IsRendering = False Then

 PRender.Cls

 Select Case PDMouseID
  Case 0: PDMouseID = 1
          Shape1.Left = X: Shape1.Top = Y
          Line1.X1 = Shape1.Left: Line1.Y1 = Shape1.Top
          Line1.X2 = Shape2.Left: Line1.Y2 = Shape2.Top
          Line2.X1 = Shape2.Left: Line2.Y1 = Shape2.Top
          Line2.X2 = Shape3.Left: Line2.Y2 = Shape3.Top
          Line3.X1 = Shape3.Left: Line3.Y1 = Shape3.Top
          Line3.X2 = Shape1.Left: Line3.Y2 = Shape1.Top
  Case 1: PDMouseID = 2
          Shape2.Left = X: Shape2.Top = Y
          Line1.X1 = Shape1.Left: Line1.Y1 = Shape1.Top
          Line1.X2 = Shape2.Left: Line1.Y2 = Shape2.Top
          Line2.X1 = Shape2.Left: Line2.Y1 = Shape2.Top
          Line2.X2 = Shape3.Left: Line2.Y2 = Shape3.Top
          Line3.X1 = Shape3.Left: Line3.Y1 = Shape3.Top
          Line3.X2 = Shape1.Left: Line3.Y2 = Shape1.Top
  Case 2: PDMouseID = 0
          Shape3.Left = X: Shape3.Top = Y
          Line1.X1 = Shape1.Left: Line1.Y1 = Shape1.Top
          Line1.X2 = Shape2.Left: Line1.Y2 = Shape2.Top
          Line2.X1 = Shape2.Left: Line2.Y1 = Shape2.Top
          Line2.X2 = Shape3.Left: Line2.Y2 = Shape3.Top
          Line3.X1 = Shape3.Left: Line3.Y1 = Shape3.Top
          Line3.X2 = Shape1.Left: Line3.Y2 = Shape1.Top
 End Select

 Shape1.Visible = True: Line1.Visible = True
 Shape2.Visible = True: Line2.Visible = True
 Shape3.Visible = True: Line3.Visible = True

 Clicked1 = True

 End If

End Sub
Private Sub PTex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If IsRendering = False Then

 Select Case PTexMouseID
  Case 0: PTexMouseID = 1
          Shape4.Left = X: Shape4.Top = Y
          Line4.X1 = Shape4.Left: Line4.Y1 = Shape4.Top
          Line4.X2 = Shape5.Left: Line4.Y2 = Shape5.Top
          Line5.X1 = Shape5.Left: Line5.Y1 = Shape5.Top
          Line5.X2 = Shape6.Left: Line5.Y2 = Shape6.Top
          Line6.X1 = Shape6.Left: Line6.Y1 = Shape6.Top
          Line6.X2 = Shape4.Left: Line6.Y2 = Shape4.Top
  Case 1: PTexMouseID = 2
          Shape5.Left = X: Shape5.Top = Y
          Line4.X1 = Shape4.Left: Line4.Y1 = Shape4.Top
          Line4.X2 = Shape5.Left: Line4.Y2 = Shape5.Top
          Line5.X1 = Shape5.Left: Line5.Y1 = Shape5.Top
          Line5.X2 = Shape6.Left: Line5.Y2 = Shape6.Top
          Line6.X1 = Shape6.Left: Line6.Y1 = Shape6.Top
          Line6.X2 = Shape4.Left: Line6.Y2 = Shape4.Top
  Case 2: PTexMouseID = 0
          Shape6.Left = X: Shape6.Top = Y
          Line4.X1 = Shape4.Left: Line4.Y1 = Shape4.Top
          Line4.X2 = Shape5.Left: Line4.Y2 = Shape5.Top
          Line5.X1 = Shape5.Left: Line5.Y1 = Shape5.Top
          Line5.X2 = Shape6.Left: Line5.Y2 = Shape6.Top
          Line6.X1 = Shape6.Left: Line6.Y1 = Shape6.Top
          Line6.X2 = Shape4.Left: Line6.Y2 = Shape4.Top
 End Select

 Shape4.Visible = True: Line4.Visible = True
 Shape5.Visible = True: Line5.Visible = True
 Shape6.Visible = True: Line6.Visible = True

 Clicked2 = True

 End If

End Sub

