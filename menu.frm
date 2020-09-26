VERSION 5.00
Begin VB.Form menu 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menu 
      Caption         =   "a"
      Begin VB.Menu fff 
         Caption         =   "Show Calender"
      End
      Begin VB.Menu fvvf 
         Caption         =   "BS-AD/AD-BS CONVERTOR"
      End
      Begin VB.Menu ert 
         Caption         =   "Skin"
         Begin VB.Menu skin 
            Caption         =   "Grey"
            Index           =   1
         End
         Begin VB.Menu skin 
            Caption         =   "Blue"
            Index           =   2
         End
         Begin VB.Menu skin 
            Caption         =   "Pink"
            Index           =   3
         End
         Begin VB.Menu skin 
            Caption         =   "Yellow"
            Index           =   4
         End
         Begin VB.Menu skin 
            Caption         =   "Green"
            Index           =   5
         End
         Begin VB.Menu skin 
            Caption         =   "Copper"
            Index           =   6
         End
         Begin VB.Menu skin 
            Caption         =   "Orange"
            Index           =   7
         End
         Begin VB.Menu skin 
            Caption         =   "White"
            Index           =   8
         End
         Begin VB.Menu skin 
            Caption         =   "Wheat"
            Index           =   9
         End
         Begin VB.Menu skin 
            Caption         =   "Sunset"
            Index           =   10
         End
         Begin VB.Menu skin 
            Caption         =   "Ocean"
            Index           =   11
         End
      End
      Begin VB.Menu yhytj 
         Caption         =   "About"
      End
      Begin VB.Menu jjk 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fff_Click()
calender.Show
End Sub

Private Sub fvvf_Click()
dateconvertor.Show
End Sub

Private Sub jjk_Click()
End
End Sub

Private Sub skin_Click(Index As Integer)
For i = 1 To 11
skin(i).Checked = False
Next i


skin(Index).Checked = True
  On Error Resume Next
devcalender.backimage.Picture = LoadPicture(App.Path & "\tools\" & Index & ".set")
  devcalender.eventbar.Picture = LoadPicture(App.Path & "\tools\" & Index & ".set")

  On Error Resume Next
Open App.Path & "\tools\skin.set" For Output As #1
Write #1, Index
Close #1


End Sub

Private Sub yhytj_Click()
About.Show
End Sub
