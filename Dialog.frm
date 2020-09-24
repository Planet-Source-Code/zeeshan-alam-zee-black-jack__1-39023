VERSION 5.00
Begin VB.Form Dialog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2310
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   105
      TabIndex        =   3
      Top             =   2205
      Width           =   5790
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   105
      X2              =   5880
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1170
      Left            =   1890
      TabIndex        =   2
      Top             =   735
      Width           =   3690
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1470
      X2              =   5565
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YOU LOST!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1155
      TabIndex        =   1
      Top             =   105
      Width           =   4215
   End
   Begin VB.Image imgLost 
      Height          =   1485
      Left            =   105
      Picture         =   "Dialog.frx":0000
      Stretch         =   -1  'True
      Top             =   525
      Width           =   1485
   End
   Begin VB.Image imgwin 
      Height          =   1485
      Left            =   105
      Picture         =   "Dialog.frx":08CA
      Stretch         =   -1  'True
      Top             =   525
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   3690
      Left            =   0
      Top             =   0
      Width           =   6105
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Caption = lblTitle
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
