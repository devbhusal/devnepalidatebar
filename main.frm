VERSION 5.00
Begin VB.Form devcalender 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "devdate"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   840
      Top             =   1920
   End
   Begin VB.Image close2 
      Height          =   315
      Left            =   2280
      Picture         =   "main.frx":A202
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image close1 
      Height          =   300
      Left            =   2280
      Picture         =   "main.frx":A6E3
      Top             =   120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label showtodaysevent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "World handicapped day"
      BeginProperty Font 
         Name            =   "Action Man Extended"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Image eventbar 
      Height          =   375
      Left            =   0
      Picture         =   "main.frx":AABB
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label nepaliday 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "saturday"
      BeginProperty Font 
         Name            =   "Old Virus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   3
      Top             =   1000
      Width           =   2175
   End
   Begin VB.Label nepalidate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   45
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      MouseIcon       =   "main.frx":BCCA
      MousePointer    =   15  'Size All
      TabIndex        =   2
      ToolTipText     =   "Rightclick to start menu"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label nepalimonth 
      BackStyle       =   0  'Transparent
      Caption         =   "baisakh"
      BeginProperty Font 
         Name            =   "Nanosecond Thick BRK"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MousePointer    =   15  'Size All
      TabIndex        =   1
      Top             =   335
      Width           =   1575
   End
   Begin VB.Label nepaliyear 
      BackStyle       =   0  'Transparent
      Caption         =   "2011"
      BeginProperty Font 
         Name            =   "Alba Super"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      MousePointer    =   15  'Size All
      TabIndex        =   0
      Top             =   550
      Width           =   1215
   End
   Begin VB.Image backimage 
      Height          =   1455
      Left            =   5
      MouseIcon       =   "main.frx":11F54
      MousePointer    =   15  'Size All
      Picture         =   "main.frx":181DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2490
   End
End
Attribute VB_Name = "devcalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hwnd As Long, _
                    ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hwnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) As Long

    Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal crKey As Long, _
                    ByVal bAlpha As Byte, _
                    ByVal dwFlags As Long) As Long
    Private Const GWL_STYLE = (-16)
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2

Dim bMoveFrom As Boolean, LastPoint As POINTAPI
 Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
(ByVal hwnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long
'tesung

Private Sub backimage_DblClick()
calender.Show
End Sub

Private Sub backimage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseDown(Button, Shift, X, Y)

End Sub


Private Sub backimage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseMove(Button, Shift, X, Y)

End Sub


Private Sub backimage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseUp(Button, Shift, X, Y)

End Sub

Private Sub close1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close1.Visible = 0
close2.Visible = 1
End Sub

Private Sub close2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
close2.BorderStyle = 1
End
End Sub

Private Sub Form_Load()
  On Error Resume Next
Unload splash
Dim fcnt As Integer


iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY

Me.BackColor = vbCyan
    On Error Resume Next
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY



' SetWindowRgn hwnd, _
'  CreateRoundRectRgn(0, 11, 160, 88, 40, 60), _
'  True


  On Error Resume Next
Open App.Path & "\tools\location.set" For Input As #1
Input #1, atop
Input #1, aleft
Close #1
  
Me.Top = atop
Me.Left = aleft

  Call adjustlocation
aaaa = 11
On Error Resume Next
Open App.Path & "\tools\skin.set" For Input As #1
Input #1, aaaa
Close #1
  On Error Resume Next

backimage.Picture = LoadPicture(App.Path & "\tools\" & aaaa & ".set")
eventbar.Picture = LoadPicture(App.Path & "\tools\" & aaaa & ".set")
For i = 1 To 11
menu.skin(i).Checked = False
Next i
menu.skin(aaaa).Checked = True





Call showdate


End Sub

'
Private Sub nepalidate_DblClick()
calender.Show

End Sub


Private Sub nepalidate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu menu.menu
Exit Sub
End If



Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub


Private Sub nepalidate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    
    close1.Visible = 1
    close2.Visible = False
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
   Me.Move Me.Left + iDX&, Me.Top + iDY&

End Sub


Private Sub nepalidate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMoveFrom = False
  
  
  Call adjustlocation
End Sub


Private Sub nepaliday_DblClick()
calender.Show

End Sub

Private Sub nepaliday_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseDown(Button, Shift, X, Y)

End Sub


Private Sub nepaliday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseMove(Button, Shift, X, Y)

End Sub


Private Sub nepaliday_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseUp(Button, Shift, X, Y)

End Sub

Private Sub nepalimonth_DblClick()
calender.Show

End Sub

Private Sub nepalimonth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseDown(Button, Shift, X, Y)

End Sub


Private Sub nepalimonth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseMove(Button, Shift, X, Y)

End Sub


Private Sub nepalimonth_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseUp(Button, Shift, X, Y)

End Sub


Private Sub nepaliyear_DblClick()
calender.Show

End Sub

Private Sub nepaliyear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseDown(Button, Shift, X, Y)

End Sub


Private Sub nepaliyear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseMove(Button, Shift, X, Y)

End Sub


Private Sub nepaliyear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseUp(Button, Shift, X, Y)

End Sub


Private Sub showtodaysevent_DblClick()
calender.Show

End Sub


Private Sub showtodaysevent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseDown(Button, Shift, X, Y)

End Sub


Private Sub showtodaysevent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseMove(Button, Shift, X, Y)

End Sub


Private Sub showtodaysevent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call nepalidate_MouseUp(Button, Shift, X, Y)

End Sub


Private Sub Timer1_Timer()
Call showdate
close1.Visible = 0
close2.Visible = 0
End Sub


