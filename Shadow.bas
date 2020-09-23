Attribute VB_Name = "Shadow"
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Function Drop_Shadow(TargetForm As Form)
    Dim i, j As Integer
    Dim SForm As Form
    j = 1
    For i = 235 To 255
    j = j + 5
        Set SForm = New SForm
            With SForm
            .Height = TargetForm.Height
            .Width = TargetForm.Width
            .Top = TargetForm.Top + j
            .Left = TargetForm.Left + j
            Set_Trans SForm, 255 - i
            SForm.Show
        End With
    Next
End Function

Public Function Set_Trans(TargetForm As Form, Tranparency As Integer)
Debug.Print Tranparency
    Ret = GetWindowLong(TargetForm.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong TargetForm.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes TargetForm.hwnd, RGB(255, 0, 0), Tranparency, LWA_ALPHA + LWA_COLORKEY
End Function

