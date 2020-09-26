Attribute VB_Name = "mdlMain"
Option Explicit


Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000

Private m_bInIDE As Boolean

Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()

   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   
   On Error Resume Next
splash.Show

End Sub

Public Sub UnloadApp()
   If Not InIDE() Then
      SetErrorMode SEM_NOGPFAULTERRORBOX
   End If
End Sub

Public Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function
