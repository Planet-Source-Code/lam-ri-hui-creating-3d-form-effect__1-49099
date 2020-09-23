VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "3D Form Demo by Lam Ri Hui"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   3450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ThreeDForm(frmForm As Form)
Const cPi = 3.1415926
Dim intLineWidth As Integer
intLineWidth = 5
Dim intSaveScaleMode As Integer
intSaveScaleMode = frmForm.ScaleMode
frmForm.ScaleMode = 3
Dim intScaleWidth As Integer
Dim intScaleHeight As Integer
intScaleWidth = frmForm.ScaleWidth
intScaleHeight = frmForm.ScaleHeight
frmForm.Cls
frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, _
intScaleHeight), &H808080, BF
frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, _
intScaleHeight), &H808080, BF
Dim intCircleWidth As Integer
intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
frmForm.FillStyle = 0
frmForm.FillColor = QBColor(15)
frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), _
intCircleWidth, QBColor(15), -3.1415926, -3.90953745777778
frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), _
intCircleWidth, QBColor(15), -0.78539815, -1.5707963
frmForm.Line (0, intScaleHeight)-(0, 0), 0
frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
frmForm.ScaleMode = intSaveScaleMode
End Sub



Private Sub Form_Load()
ThreeDForm Me
End Sub

Private Sub Form_Resize()
ThreeDForm Me
End Sub
