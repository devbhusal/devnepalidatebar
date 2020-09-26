VERSION 5.00
Begin VB.Form splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   3720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "with nepali calender"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   2220
      Left            =   6330
      Picture         =   "splash.frx":0000
      Top             =   -50
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   555
      Left            =   2160
      Picture         =   "splash.frx":43E3
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1740
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing program for its 1st use..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 2.0"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "devNEPALIDATEBAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   480
      Picture         =   "splash.frx":51EA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   3960
      Left            =   0
      Picture         =   "splash.frx":62D1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6480
   End
End
Attribute VB_Name = "splash"
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



Private Sub formshow()
status.Caption = "Configuring fonts..."
  
  On Error GoTo last:
  
  Call Shell(App.Path & "\tools\fontinstaller.exe", vbNormalFocus)

  
 
 

For i = 1 To 2005000
DoEvents
Next i
2:
status.Caption = "Configuring Settings..."
On Error Resume Next
Open App.Path & "\tools\location.set" For Output As #1
Write #1, 0
Write #1, Screen.Width

Close #1
On Error Resume Next
Open App.Path & "\tools\cnt.set" For Output As #1
Write #1, 1
Close #1
For i = 1 To 2005000
DoEvents
Next i
status.Caption = "Installation completed. Enjoy using it for free"
For i = 1 To 3005000
DoEvents
Next i
  On Error Resume Next
devcalender.Show
Exit Sub
last:
MsgBox "..\tools\fontinstaller.exe could not be loaded. it may be deleted .Please reinstall devNEPALIDATEBAR" & vbNewLine & "Software may not work properly"
GoTo 2:

End Sub



Private Sub Command1_Click()




End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End
Me.BackColor = vbCyan
    On Error Resume Next
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2



fcnt = 0
For i = 0 To (Screen.FontCount - 1)
 aaa = UCase(Screen.Fonts(i))
 
 If aaa = "ACTION MAN" Then fcnt = fcnt + 1
If aaa = "ACTION MAN SHADED" Then fcnt = fcnt + 1
If aaa = "ALBA SUPER" Then fcnt = fcnt + 1
If aaa = "BOOKMAN OLD STYLE" Then fcnt = fcnt + 1
If aaa = "LCD" Then fcnt = fcnt + 1
If aaa = "NANOSECOND THICK BRK" Then fcnt = fcnt + 1
If aaa = "OLD VIRUS" Then fcnt = fcnt + 1
If aaa = "VERDANA" Then fcnt = fcnt + 1
If aaa = "WISHFULWAVES" Then fcnt = fcnt + 1
If aaa = "ATLAS" Then fcnt = fcnt + 1
If aaa = "SCREWEDSW" Then fcnt = fcnt + 1
Next i
countt = 0
On Error Resume Next
Open App.Path & "\tools\cnt.set" For Input As #1
Input #1, countt
Close #1

If fcnt <> 11 Or countt = 0 Then
  If fcnt <> 11 Then Label2.Caption = "Fonts are missing!!!"
  If countt = 0 Then Label2.Caption = "Initializing program for its 1st use ..."
Timer1.Enabled = True
Else
devcalender.Show
End If




End Sub

Private Sub Timer1_Timer()
Call formshow

Timer1.Enabled = False

End Sub


