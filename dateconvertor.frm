VERSION 5.00
Begin VB.Form dateconvertor 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date convertor"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "BS to AD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4680
      TabIndex        =   7
      Top             =   360
      Width           =   4575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Copy to clipboard"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox bs3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox bs2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox bs1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Select year ,month and date of BS to convert it to AD. For this package only 2000-2089 year can be converted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "English date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "AD to BS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Copy to clipboard"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox ad3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox ad2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox ad1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select year month and date of AD to convert it to BS. For this package only 1944-2033 year can be converted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nepali date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   4095
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "devMEDIA Nepal"
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
      Left            =   720
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   375
   End
End
Attribute VB_Name = "dateconvertor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

Clipboard.SetText Label1.Caption
MsgBox "Converted date is exported to clipboard. Paste it wherever needed"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Clipboard.SetText Label3.Caption
MsgBox "Converted date is exported to clipboard. Paste it wherever needed"

End Sub


Private Sub Form_Load()
On Error Resume Next
ad1.Clear
For i = 1944 To 2033
ad1.AddItem i
Next i
ad2.Clear
ad2.AddItem "January"
ad2.AddItem "February"
ad2.AddItem "March"
ad2.AddItem "April"
ad2.AddItem "May"
ad2.AddItem "June"
ad2.AddItem "July"
ad2.AddItem "August"
ad2.AddItem "September"
ad2.AddItem "October"
ad2.AddItem "November"
ad2.AddItem "December"

ad3.Clear
For i = 1 To 31
ad3.AddItem i
Next i
ad1.ListIndex = 0
ad2.ListIndex = 0
ad3.ListIndex = 0


bs1.Clear
For i = 2000 To 2089
bs1.AddItem i
Next i
bs2.Clear
bs2.AddItem "Baisakh"
bs2.AddItem "Jestha"
bs2.AddItem "Ashadh"
bs2.AddItem "Shrawan"
bs2.AddItem "Bhadra"
bs2.AddItem "Ashoj"
bs2.AddItem "Kartik"
bs2.AddItem "Mangsir"
bs2.AddItem "Poush"
bs2.AddItem "Magh"
bs2.AddItem "Falgun"
bs2.AddItem "Chaitra"

bs3.Clear
For i = 1 To 32
bs3.AddItem i
Next i
bs1.ListIndex = 0
bs2.ListIndex = 0
bs3.ListIndex = 0



End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Call eng_to_nep(ad1.Text, ad2.ListIndex + 1, ad3.Text)
Label1.Caption = nep_year & " " & get_nepali_month(nep_month) & " " & nep_date & ", " & get_day_of_week(nep_day)

Call nep_to_eng(bs1.Text, bs2.ListIndex + 1, bs3.Text)
Label3.Caption = eng_year & " " & get_english_month(eng_month) & " " & eng_date & ", " & get_day_of_week(eng_day)




End Sub


