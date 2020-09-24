Attribute VB_Name = "General"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MOVIN FORM functions!!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const MF_BYCOMMAND = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const SC_CLOSE = &HF060&

