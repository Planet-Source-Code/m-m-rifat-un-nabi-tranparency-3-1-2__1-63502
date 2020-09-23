VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\  Programmer - M. M. Rifat-Un-Nabi      \\
'\\  e-mail     - torifat@gmail.com        \\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY As Long = &H1
Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Dim tmpSt As Long
    tmpSt = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    tmpSt = tmpSt Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, tmpSt
    SetLayeredWindowAttributes Me.hwnd, RGB(255, 0, 0), 0, LWA_COLORKEY
    Drop_Shadow frmMain
End Sub
