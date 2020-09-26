VERSION 5.00
Begin VB.Form calender 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "devNEPALICALENDER"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   12675
   Icon            =   "calender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   6375
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   41
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   145
         Top             =   4800
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   41
            Left            =   0
            TabIndex        =   147
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   41
            Left            =   0
            TabIndex        =   146
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   40
         Left            =   4440
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   142
         Top             =   4800
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   40
            Left            =   0
            TabIndex        =   144
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   40
            Left            =   0
            TabIndex        =   143
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   39
         Left            =   3600
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   139
         Top             =   4800
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   39
            Left            =   0
            TabIndex        =   141
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   39
            Left            =   0
            TabIndex        =   140
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   38
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   136
         Top             =   4800
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   38
            Left            =   0
            TabIndex        =   138
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   38
            Left            =   0
            TabIndex        =   137
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   37
         Left            =   1920
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   133
         Top             =   4800
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   37
            Left            =   0
            TabIndex        =   135
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   37
            Left            =   0
            TabIndex        =   134
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   36
         Left            =   1080
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   130
         Top             =   4800
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   36
            Left            =   0
            TabIndex        =   132
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   36
            Left            =   0
            TabIndex        =   131
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   35
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   127
         Top             =   4800
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   35
            Left            =   0
            TabIndex        =   129
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   35
            Left            =   0
            TabIndex        =   128
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   34
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   124
         Top             =   4080
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   34
            Left            =   0
            TabIndex        =   126
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   34
            Left            =   0
            TabIndex        =   125
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   33
         Left            =   4440
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   121
         Top             =   4080
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   33
            Left            =   0
            TabIndex        =   123
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   33
            Left            =   0
            TabIndex        =   122
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   32
         Left            =   3600
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   118
         Top             =   4080
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   32
            Left            =   0
            TabIndex        =   120
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   32
            Left            =   0
            TabIndex        =   119
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   31
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   115
         Top             =   4080
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   31
            Left            =   0
            TabIndex        =   117
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   31
            Left            =   0
            TabIndex        =   116
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   30
         Left            =   1920
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   112
         Top             =   4080
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   30
            Left            =   0
            TabIndex        =   114
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   30
            Left            =   0
            TabIndex        =   113
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   29
         Left            =   1080
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   109
         Top             =   4080
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   29
            Left            =   0
            TabIndex        =   111
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   29
            Left            =   0
            TabIndex        =   110
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   28
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   106
         Top             =   4080
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   28
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   28
            Left            =   0
            TabIndex        =   107
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   27
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   103
         Top             =   3360
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   27
            Left            =   0
            TabIndex        =   105
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   27
            Left            =   0
            TabIndex        =   104
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   26
         Left            =   4440
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   100
         Top             =   3360
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   26
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   26
            Left            =   0
            TabIndex        =   101
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   25
         Left            =   3600
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   97
         Top             =   3360
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   25
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   25
            Left            =   0
            TabIndex        =   98
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   24
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   94
         Top             =   3360
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   24
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   24
            Left            =   0
            TabIndex        =   95
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   23
         Left            =   1920
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   91
         Top             =   3360
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   23
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   23
            Left            =   0
            TabIndex        =   92
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   22
         Left            =   1080
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   88
         Top             =   3360
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   22
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   22
            Left            =   0
            TabIndex        =   89
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   21
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   85
         Top             =   3360
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   21
            Left            =   0
            TabIndex        =   87
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   21
            Left            =   0
            TabIndex        =   86
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   20
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   82
         Top             =   2640
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   20
            Left            =   0
            TabIndex        =   84
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   20
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   19
         Left            =   4440
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   79
         Top             =   2640
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   19
            Left            =   0
            TabIndex        =   81
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   19
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   18
         Left            =   3600
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   76
         Top             =   2640
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   18
            Left            =   0
            TabIndex        =   78
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   18
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   17
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   73
         Top             =   2640
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   17
            Left            =   0
            TabIndex        =   75
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   17
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   16
         Left            =   1920
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   70
         Top             =   2640
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   16
            Left            =   0
            TabIndex        =   72
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   16
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   15
         Left            =   1080
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   67
         Top             =   2640
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   15
            Left            =   0
            TabIndex        =   69
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   15
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   14
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   64
         Top             =   2640
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   14
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   14
            Left            =   0
            TabIndex        =   65
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   13
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   61
         Top             =   1920
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   13
            Left            =   0
            TabIndex        =   63
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   13
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   12
         Left            =   4440
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   58
         Top             =   1920
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   12
            Left            =   0
            TabIndex        =   60
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   12
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   11
         Left            =   3600
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   55
         Top             =   1920
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   11
            Left            =   0
            TabIndex        =   57
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   11
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   10
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   52
         Top             =   1920
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   10
            Left            =   0
            TabIndex        =   54
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   10
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   9
         Left            =   1920
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   49
         Top             =   1920
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   9
            Left            =   0
            TabIndex        =   51
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   9
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   8
         Left            =   1080
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   46
         Top             =   1920
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   8
            Left            =   0
            TabIndex        =   48
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   8
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   7
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   43
         Top             =   1920
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   7
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   7
            Left            =   0
            TabIndex        =   44
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   6
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   40
         Top             =   1200
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   6
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   6
            Left            =   0
            TabIndex        =   41
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   5
         Left            =   4440
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   37
         Top             =   1200
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   5
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   5
            Left            =   0
            TabIndex        =   38
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   4
         Left            =   3600
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   34
         Top             =   1200
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   4
            Left            =   0
            TabIndex        =   35
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   3
         Left            =   2760
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   31
         Top             =   1200
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   3
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   3
            Left            =   0
            TabIndex        =   32
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   2
         Left            =   1920
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   28
         Top             =   1200
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   2
            Left            =   0
            TabIndex        =   29
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   1
         Left            =   1080
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   25
         Top             =   1200
         Width           =   825
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   855
         End
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   0
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   795
         TabIndex        =   21
         Top             =   1200
         Width           =   825
         Begin VB.Label edayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "sep 22"
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
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
         Begin VB.Label dayyy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Image Image2 
         Height          =   450
         Left            =   5640
         Picture         =   "calender.frx":A202
         ToolTipText     =   "Next month"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   240
         Picture         =   "calender.frx":AD84
         ToolTipText     =   "Previous month"
         Top             =   360
         Width           =   525
      End
      Begin VB.Label english 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "december 64, 2011 to december 75,2011"
         BeginProperty Font 
            Name            =   "Nanosecond Thick BRK"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUN"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MON"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TUE"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WED"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "THU"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   4
         Left            =   3600
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FRI"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   5
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label DAYlabel 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SAT"
         BeginProperty Font 
            Name            =   "ScrewedSW"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   6
         Left            =   5280
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      Caption         =   "Events (in BS)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   9720
      TabIndex        =   11
      Top             =   1080
      Width           =   2895
      Begin VB.ListBox eventslist 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   4830
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Jump to(AD)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Jump"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox ad2 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox ad1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Jump to(BS)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Jump"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1455
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Image mini2 
      Height          =   300
      Left            =   840
      Picture         =   "calender.frx":BA6E
      Top             =   120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image mini1 
      Height          =   285
      Left            =   840
      Picture         =   "calender.frx":BF29
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Image close2 
      Height          =   270
      Left            =   120
      Picture         =   "calender.frx":C318
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image close1 
      Height          =   285
      Left            =   120
      Picture         =   "calender.frx":C830
      Stretch         =   -1  'True
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump"
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
      Left            =   3240
      TabIndex        =   151
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label command4 
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   150
      Top             =   720
      Width           =   615
   End
   Begin VB.Label command2 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump to today"
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
      Left            =   3240
      MousePointer    =   4  'Icon
      TabIndex        =   149
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.devendrabhusal.com.np"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   148
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image4 
      Height          =   6015
      Left            =   60
      Picture         =   "calender.frx":CD2D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B.S"
      BeginProperty Font 
         Name            =   "WishfulWaves"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.Label yearrr 
      BackStyle       =   0  'Transparent
      Caption         =   "2088"
      BeginProperty Font 
         Name            =   "Atlas"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label monthhh 
      BackStyle       =   0  'Transparent
      Caption         =   "Kartik"
      BeginProperty Font 
         Name            =   "Alba Super"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   8040
      Left            =   0
      Picture         =   "calender.frx":1AA84
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12705
   End
End
Attribute VB_Name = "calender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmonth As Integer
Dim cyear As Integer
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





Private Sub close1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close2.Visible = 1
close1.Visible = 0

End Sub


Private Sub close2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload calender
End Sub

Private Sub close2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

close2.Visible = 1
close1.Visible = 0


End Sub


Private Sub Command1_Click()
On Error Resume Next
Call showcalender(bs1.ListIndex + 2000, bs2.ListIndex + 1)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Call eng_to_nep(Right(Date$, 4), Val(Left(Date$, 2)), Mid(Date$, 4, 2))
 currentmonth = nep_month
 currentyear = nep_year
 currentdate = nep_date
 Call showcalender(currentyear, currentmonth)

For i = 0 To 41
If dayyy(i).Caption = currentdate Then
'dayyy(i).Style = 1
dayyy(i).BackColor = &HC0FFC0
End If
Next i
 

End Sub

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.FontUnderline = 1
End Sub


Private Sub Command3_Click()
On Error Resume Next

Call eng_to_nep(ad1.ListIndex + 1944, ad2.ListIndex + 1, 1)


Call showcalender(nep_year, nep_month)

End Sub

Private Sub Command4_Click()
On Error Resume Next


Me.PrintForm


End Sub

Private Sub Command5_Click()

End Sub


Private Sub command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
command4.FontUnderline = 1
End Sub

Private Sub Form_Load()

Call eng_to_nep(Right(Date$, 4), Val(Left(Date$, 2)), Mid(Date$, 4, 2))
 currentmonth = nep_month
 currentyear = nep_year
 currentdate = nep_date
 
 
 Me.Caption = "devNEPALI CALENDER 1.0 ....Current date is " & nep_year & " " & get_nepali_month(nep_month) & " " & nep_date
 iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY

Me.BackColor = vbCyan
    On Error Resume Next
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY

 
 Call showcalender(currentyear, currentmonth)
For i = 0 To 41
If dayyy(i).Caption = currentdate Then
'dayyy(i).Style = 1
dayyy(i).BackColor = &HC0FFFF
End If
calender.edayyy(i).Width = 855
Next i
 
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

bs1.ListIndex = currentyear - 2000
bs2.ListIndex = currentmonth - 1
ad1.ListIndex = 2012 - 1944
ad2.ListIndex = 0
 
 End Sub
 
 
 Public Sub showcalender(nepyearr, nepmonthh)
  Call nep_to_eng(nepyearr, nepmonthh, 1)
 
 cmonth = nepmonthh
 cyear = nepyearr
 firstday = eng_day
 
 For i = 0 To 41
 dayyy(i).Caption = ""
edayyy(i).Caption = ""
dayyy(i).BackColor = &HE0E0E0
dayyy(i).Visible = 0
 edayyy(i).Visible = 0
 Next i
 
 
 eventslist.Clear
 Call checkevents
 Call initilizeClass
 p = firstday - 1
ad = 0
' display nepali dates and eng dates
 X = Left(get_english_month(eng_month), 3)
 For i = 1 To bs(nepyearr - 2000)(nepmonthh)
   dayyy(p).Caption = i
   If (eng_date + ad) >= 28 Then
     Call nep_to_eng(nepyearr, nepmonthh, i)
     ad = 0
     X = Left(get_english_month(eng_month), 3)
   End If
    edayyy(p).Caption = X & " " & eng_date + ad
    dayyy(p).Visible = 1
    edayyy(p).Visible = 1
    ad = ad + 1
    p = p + 1
 Next i
 
' caption
 yearrr.Caption = nepyearr
 monthhh.Caption = get_nepali_month(nepmonthh)


  Call nep_to_eng(nepyearr, nepmonthh, 1)
p1 = get_english_month(eng_month) & " " & eng_date & "(" & eng_year & ") to "
dev1_year = eng_year
dev1_month = eng_month
a = bs(nepyearr - 2000)(nepmonthh)
Call nep_to_eng(nepyearr, nepmonthh, a)
p2 = p1 & get_english_month(eng_month) & " " & eng_date & "(" & eng_year & ")"
dev2_year = eng_year
dev2_month = eng_month
english.Caption = p2
For i = 6 To 41 Step 7
dayyy(i).BackColor = 12632319
Next i
Call checkevents

' Event listing()
DoEvents
eventslist.Clear
 i = 0
Do
If listofevents(i)(1) = "bs" And listofevents(i)(2) = nepmonthh Then
eventslist.AddItem "(" & listofevents(i)(3) & ") " & listofevents(i)(0)
End If

If listofevents(i)(1) = "ad" And (listofevents(i)(2) = dev1_month Or listofevents(i)(2) = dev2_month) Then
 tempstring = Left(get_english_month(listofevents(i)(2)), 3) & " " & listofevents(i)(3)
  For j = 0 To 41
  If edayyy(j).Caption = tempstring Then eventslist.AddItem "(" & dayyy(j).Caption & ") " & listofevents(i)(0)
  Next j




End If



i = i + 1


Loop Until listofevents(i)(0) = "END"

 End Sub

Private Sub Image1_Click()
On Error Resume Next

If cmonth >= 2 And cmonth <= 12 Then

cmonth = cmonth - 1
Call showcalender(cyear, cmonth)
Else

cmonth = 12
cyear = cyear - 1
Call showcalender(cyear, cmonth)
End If






End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0

End Sub


Private Sub Image2_Click()
On Error Resume Next
If cmonth >= 1 And cmonth <= 11 Then

cmonth = cmonth + 1
Call showcalender(cyear, cmonth)
Else

cmonth = 1
cyear = cyear + 1
Call showcalender(cyear, cmonth)
End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1

End Sub


Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 0

End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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


Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame1.Visible = 0
   Frame2.Visible = 0
   close1.Visible = 1
close2.Visible = 0
mini1.Visible = 1
mini2.Visible = 0
command4.FontUnderline = 0
Command2.FontUnderline = 0


Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
   Me.Move Me.Left + iDX&, Me.Top + iDY&

End Sub


Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bMoveFrom = False
End Sub


Private Sub Label2_Click()
Call Shell("cmd /c explorer http://www.devendrabhusal.com.np", vbHide)

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.Visible = 1
Frame2.Visible = 1

End Sub


Private Sub mini1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mini1.Visible = 0
mini2.Visible = 1
End Sub


Private Sub mini2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bMoveFrom = False

Me.WindowState = vbMinimized
End Sub

Private Sub mini2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mini1.Visible = 0
mini2.Visible = 1

End Sub


