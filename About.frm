VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About devNEPALIDATEBAR"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Version Info"
      Height          =   1935
      Left            =   3120
      TabIndex        =   2
      Top             =   3120
      Width           =   4935
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "info@devendrabhusal.com.np"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "For other queries mail me to"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.devendrabhusal.com.np"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "v2.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "devNEPALIDATEBAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Image Image3 
         Height          =   840
         Left            =   120
         Picture         =   "About.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Image Image6 
      Height          =   840
      Left            =   1920
      Picture         =   "About.frx":10E7
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   840
   End
   Begin VB.Image Image5 
      Height          =   720
      Left            =   240
      Picture         =   "About.frx":1D9C
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "devWIPER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "devPLAYER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Other software from devMEDIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":2590
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   2760
      Picture         =   "About.frx":26FE
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dev Bhusal [ Programmer ]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   2355
      Left            =   120
      Picture         =   "About.frx":3505
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2355
   End
   Begin VB.Image Image4 
      Height          =   8100
      Left            =   -2040
      Picture         =   "About.frx":A0D2
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   13320
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontUnderline = 0
Label7.FontUnderline = 0
Label8.FontUnderline = 0
Label10.FontUnderline = 0

End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontUnderline = 0
Label7.FontUnderline = 0
Label8.FontUnderline = 0
Label10.FontUnderline = 0

End Sub

Private Sub Label10_Click()
Call Shell("cmd /c explorer mailto:info@devendrabhusal.com.np", vbHide)

End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.FontUnderline = 1

End Sub


Private Sub Label5_Click()
Call Shell("cmd /c explorer http://www.devendrabhusal.com.np", vbHide)
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontUnderline = 1
End Sub

Private Sub Label7_Click()
Call Shell("cmd /c explorer http://www.devmedianepal.tk/", vbHide)

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontUnderline = 1

End Sub


Private Sub Label8_Click()
Call Shell("cmd /c explorer http://www.devmedianepal.tk/", vbHide)

End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontUnderline = 1

End Sub


