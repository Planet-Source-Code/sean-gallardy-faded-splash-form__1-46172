VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   Caption         =   "FrmIntro"
   ClientHeight    =   5400
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AlphaTrans 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   3600
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was adapted from Blue Eyes
'You can find the original at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=41053&lngWId=1
'This code was put into a simple Form with a Timer (yuk vb's timer) but it works with everyone

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000


Dim Current As Integer ' current alpha transparency 0 = transparent 255 = opaque
Dim Max As Integer



Private Sub AlphaTrans_Timer()

Current = Current + 5 '+5 is smooth from the 150 start, you can change
                      'experiment to find the one you like
If Current - 1 >= Max Then
    AlphaTrans.Enabled = False
    Transparent FrmIntro.hWnd, 255
    Exit Sub
End If

Transparent FrmIntro.hWnd, Current

End Sub

Private Sub Form_Load()
AlphaTrans.Interval = 1
AlphaTrans.Enabled = True
Current = 150
Max = 255
Transparent FrmIntro.hWnd, Current
End Sub

Private Function Transparent(ByVal hWnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
      Transparent = 1
    Else
      Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      SetWindowLong hWnd, GWL_EXSTYLE, Msg
      SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
      Transparent = 0
    End If
    If Err Then
      Transparent = 2
    End If
End Function

