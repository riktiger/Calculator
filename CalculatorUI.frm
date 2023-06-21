VERSION 5.00
Begin VB.Form CalculatorUI 
   BackColor       =   &H00404040&
   Caption         =   "CALCULATOR"
   ClientHeight    =   3744
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5076
   BeginProperty Font 
      Name            =   "Old English Text MT"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3744
   ScaleWidth      =   5076
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEQUAL 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   23
      Top             =   3120
      Width           =   852
   End
   Begin VB.CommandButton btnADD 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   22
      Top             =   3120
      Width           =   852
   End
   Begin VB.CommandButton btnDOT 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1920
      TabIndex        =   21
      Top             =   3120
      Width           =   492
   End
   Begin VB.CommandButton btnPLUMI 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1080
      TabIndex        =   20
      Top             =   3120
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   492
   End
   Begin VB.CommandButton btn1BY 
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   18
      Top             =   2520
      Width           =   852
   End
   Begin VB.CommandButton btnPERCENT 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   17
      Top             =   1920
      Width           =   852
   End
   Begin VB.CommandButton btnSQRT 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   16
      Top             =   1320
      Width           =   852
   End
   Begin VB.CommandButton btnSUBTRACT 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   15
      Top             =   2520
      Width           =   852
   End
   Begin VB.CommandButton btnMULTIPLY 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   852
   End
   Begin VB.CommandButton btnDIVIDE 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   13
      Top             =   1320
      Width           =   852
   End
   Begin VB.CommandButton btnname 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   2520
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   9
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton btnname 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton btnC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton btnCE 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton btnBACKSPACE 
      Caption         =   "BACKSPACE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin VB.TextBox tbResult 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   120
      TabIndex        =   0
      Top             =   100
      Width           =   4812
   End
End
Attribute VB_Name = "CalculatorUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sLeft, sRight, sOperator As String
Dim iLeft, iRight, iResult As Double
Dim bLeft As Boolean


Private Sub btnBACKSPACE_Click()
If bLeft Then
If Len(sLeft) > 0 Then sLeft = Left(sLeft, Len(sLeft) - 1)
tbResult.Text = sLeft
Else
If Len(sRight) > 0 Then sRight = Left(sRight, Len(sRight) - 1)
tbResult.Text = sRight
End If
End Sub

Private Sub btnname_Click(Index As Integer)
AddNumber (Index)
End Sub

Private Sub btnPERCENT_Click()
AddOperator ("%")
End Sub

Private Sub btnPLUMI_Click()
sLeft = -Val(sLeft)
tbResult.Text = sLeft
End Sub

Private Sub Form_Load()
bLeft = True
End Sub
Private Sub AddNumber(sNumber As String)
If bLeft Then
sLeft = sLeft + sNumber
tbResult.Text = sLeft
Else
sRight = sRight + sNumber
tbResult.Text = sRight
End If
End Sub


Private Sub btndot_click()
AddNumber (".")
End Sub
Private Sub AddOperator(sNewOperator As String)
If bLeft Then
sOperator = sNewOperator
bLeft = False
Else
btnEQUAL_click
sOperator = sNewOperator
sRight = ""
bLeft = False
End If
End Sub
Private Sub btnADD_click()
AddOperator ("+")
End Sub
Private Sub btnSUBTRACT_click()
AddOperator ("-")
End Sub
Private Sub btnMULTIPLY_click()
AddOperator ("*")
End Sub
Private Sub btnDIVIDE_click()
AddOperator ("/")
End Sub

Private Sub btnEQUAL_click()
If sLeft <> "" And sRight = "" And sOperator <> "" Then sRight = sLeft
If sOperator = "+" Then sLeft = Val(sLeft) + Val(sRight)
If sOperator = "-" Then sLeft = Val(sLeft) - Val(sRight)
If sOperator = "*" Then sLeft = Val(sLeft) * Val(sRight)
If sOperator = "/" And Val(sRight) <> 0 Then sLeft = Val(sLeft) / (sRight) Else If sOperator = "/" Then sLeft = "Undefined"
If sOperator = "%" And Val(sRight) <> 0 Then sLeft = (Val(sLeft) / Val(sRight)) * 100 Else If sOperator = "%" Then sLeft = "Undefined"
tbResult.Text = sLeft
End Sub
Private Sub btnSQRT_click()
If sLeft <> "" Then iLeft = sLeft
If iLeft >= 0 Then tbResult.Text = iLeft ^ (1 / 2)
sLeft = tbResult.Text
If iLeft < 0 Then tbResult.Text = "Complex Number"
End Sub
Private Sub btn1BY_click()
If sLeft <> "" And sLeft <> 0 Then
iLeft = sLeft
tbResult.Text = 1 / iLeft
sLeft = tbResult.Text
End If
If sLeft = 0 Then tbResult.Text = "Undefined"
End Sub
Private Sub btnC_click()
sLeft = ""
sRight = ""
sOperator = ""
tbResult = "0"
bLeft = True
End Sub
Private Sub btnCE_click()
If bLeft Then
sLeft = ""
Else
sRight = ""
End If
tbResult.Text = "0"
End Sub



