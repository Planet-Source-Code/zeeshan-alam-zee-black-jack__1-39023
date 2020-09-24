VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Zee's Black Jack  "
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   6600
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FF80&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton cmdHide 
      BackColor       =   &H0080FF80&
      Caption         =   "H&ide"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0080FF80&
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H0080FF80&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton cmdNewDeal 
      BackColor       =   &H0080C0FF&
      Caption         =   "Deal New Hands"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1785
      Width           =   2010
   End
   Begin VB.CommandButton ccTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "Zee's Black Jack  "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   210
      Width           =   3060
   End
   Begin VB.TextBox txttime 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "00:00:00"
      ToolTipText     =   "Total time for which this game was played!!"
      Top             =   4200
      Width           =   1170
   End
   Begin VB.Timer tmrGiveTip 
      Interval        =   1000
      Left            =   840
      Top             =   5460
   End
   Begin VB.Timer tmrEffect 
      Interval        =   400
      Left            =   315
      Top             =   5460
   End
   Begin VB.TextBox txtCash 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "100"
      ToolTipText     =   "Your balance! You start with 100."
      Top             =   3465
      Width           =   1170
   End
   Begin VB.CommandButton cmdHits 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Hi&ts"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      MaskColor       =   &H8000000F&
      MouseIcon       =   "frmMain.frx":33F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Pull another card!"
      Top             =   2415
      Width           =   855
   End
   Begin VB.CommandButton cmdStand 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "&Stand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1365
      MaskColor       =   &H8000000F&
      MouseIcon       =   "frmMain.frx":3838
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Show the cards"
      Top             =   2415
      Width           =   750
   End
   Begin VB.CommandButton cmdQuit1 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7770
      TabIndex        =   5
      Top             =   7665
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame framInvisible 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6525
      Left            =   10080
      TabIndex        =   0
      Top             =   315
      Width           =   7365
      Begin VB.Label Label5 
         Caption         =   "Click & Drag the fram to see and manage all cards!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1260
         TabIndex        =   1
         Top             =   315
         Width           =   1380
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   51
         Left            =   630
         Picture         =   "frmMain.frx":3C7A
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3990
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   50
         Left            =   420
         Picture         =   "frmMain.frx":77BC
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3675
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   49
         Left            =   210
         Picture         =   "frmMain.frx":B65E
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3465
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   48
         Left            =   0
         Picture         =   "frmMain.frx":ED80
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3150
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   47
         Left            =   5670
         Picture         =   "frmMain.frx":127A2
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   525
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   46
         Left            =   5460
         Picture         =   "frmMain.frx":16964
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   315
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   45
         Left            =   5250
         Picture         =   "frmMain.frx":1A486
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   105
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   44
         Left            =   6825
         Picture         =   "frmMain.frx":1E5C8
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   4725
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   43
         Left            =   6615
         Picture         =   "frmMain.frx":223CA
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   4515
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   42
         Left            =   6405
         Picture         =   "frmMain.frx":2554C
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   4200
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   41
         Left            =   6195
         Picture         =   "frmMain.frx":2896E
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3990
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   40
         Left            =   5985
         Picture         =   "frmMain.frx":2BAD0
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3675
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   39
         Left            =   5775
         Picture         =   "frmMain.frx":2EF72
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3465
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   38
         Left            =   5565
         Picture         =   "frmMain.frx":30214
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   3150
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   37
         Left            =   5355
         Picture         =   "frmMain.frx":318B6
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   2940
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   36
         Left            =   5145
         Picture         =   "frmMain.frx":32818
         Stretch         =   -1  'True
         Tag             =   "10"
         Top             =   2730
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   35
         Left            =   4935
         Picture         =   "frmMain.frx":338DA
         Stretch         =   -1  'True
         Tag             =   "9"
         Top             =   2520
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   34
         Left            =   4725
         Picture         =   "frmMain.frx":348DC
         Stretch         =   -1  'True
         Tag             =   "9"
         Top             =   2310
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   33
         Left            =   4515
         Picture         =   "frmMain.frx":35E3E
         Stretch         =   -1  'True
         Tag             =   "9"
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   32
         Left            =   4305
         Picture         =   "frmMain.frx":37020
         Stretch         =   -1  'True
         Tag             =   "9"
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   31
         Left            =   4095
         Picture         =   "frmMain.frx":37F02
         Stretch         =   -1  'True
         Tag             =   "8"
         Top             =   1575
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   30
         Left            =   3885
         Picture         =   "frmMain.frx":39104
         Stretch         =   -1  'True
         Tag             =   "8"
         Top             =   1365
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   29
         Left            =   3675
         Picture         =   "frmMain.frx":3A146
         Stretch         =   -1  'True
         Tag             =   "8"
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   28
         Left            =   3465
         Picture         =   "frmMain.frx":3B088
         Stretch         =   -1  'True
         Tag             =   "8"
         Top             =   945
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   27
         Left            =   3255
         Picture         =   "frmMain.frx":3C56A
         Stretch         =   -1  'True
         Tag             =   "7"
         Top             =   735
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   26
         Left            =   3045
         Picture         =   "frmMain.frx":3D16C
         Stretch         =   -1  'True
         Tag             =   "7"
         Top             =   525
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   25
         Left            =   2835
         Picture         =   "frmMain.frx":3E1EE
         Stretch         =   -1  'True
         Tag             =   "7"
         Top             =   315
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   24
         Left            =   2625
         Picture         =   "frmMain.frx":3ECD0
         Stretch         =   -1  'True
         Tag             =   "7"
         Top             =   105
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   23
         Left            =   5040
         Picture         =   "frmMain.frx":3FA72
         Stretch         =   -1  'True
         Tag             =   "6"
         Top             =   5460
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   22
         Left            =   4830
         Picture         =   "frmMain.frx":407B4
         Stretch         =   -1  'True
         Tag             =   "6"
         Top             =   5250
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   21
         Left            =   4620
         Picture         =   "frmMain.frx":419D6
         Stretch         =   -1  'True
         Tag             =   "6"
         Top             =   5040
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   20
         Left            =   4410
         Picture         =   "frmMain.frx":42958
         Stretch         =   -1  'True
         Tag             =   "6"
         Top             =   4830
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   19
         Left            =   4200
         Picture         =   "frmMain.frx":437BA
         Stretch         =   -1  'True
         Tag             =   "5"
         Top             =   4620
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   18
         Left            =   3990
         Picture         =   "frmMain.frx":443DC
         Stretch         =   -1  'True
         Tag             =   "5"
         Top             =   4410
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   17
         Left            =   3780
         Picture         =   "frmMain.frx":4511E
         Stretch         =   -1  'True
         Tag             =   "5"
         Top             =   4200
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   16
         Left            =   3570
         Picture         =   "frmMain.frx":45C20
         Stretch         =   -1  'True
         Tag             =   "5"
         Top             =   3990
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   15
         Left            =   3360
         Picture         =   "frmMain.frx":46BA2
         Stretch         =   -1  'True
         Tag             =   "4"
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   14
         Left            =   3150
         Picture         =   "frmMain.frx":476E4
         Stretch         =   -1  'True
         Tag             =   "4"
         Top             =   3570
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   13
         Left            =   2940
         Picture         =   "frmMain.frx":48226
         Stretch         =   -1  'True
         Tag             =   "4"
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   12
         Left            =   2730
         Picture         =   "frmMain.frx":48CA8
         Stretch         =   -1  'True
         Tag             =   "4"
         Top             =   3045
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   11
         Left            =   2520
         Picture         =   "frmMain.frx":4962A
         Stretch         =   -1  'True
         Tag             =   "3"
         Top             =   2835
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   10
         Left            =   2310
         Picture         =   "frmMain.frx":4A24C
         Stretch         =   -1  'True
         Tag             =   "3"
         Top             =   2625
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   9
         Left            =   2100
         Picture         =   "frmMain.frx":4AD8E
         Stretch         =   -1  'True
         Tag             =   "3"
         Top             =   2415
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   8
         Left            =   1890
         Picture         =   "frmMain.frx":4B7F0
         Stretch         =   -1  'True
         Tag             =   "3"
         Top             =   2205
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   7
         Left            =   1680
         Picture         =   "frmMain.frx":4C5B2
         Stretch         =   -1  'True
         Tag             =   "2"
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   6
         Left            =   1470
         Picture         =   "frmMain.frx":4D034
         Stretch         =   -1  'True
         Tag             =   "2"
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   5
         Left            =   1260
         Picture         =   "frmMain.frx":4DA16
         Stretch         =   -1  'True
         Tag             =   "2"
         Top             =   1470
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   4
         Left            =   1050
         Picture         =   "frmMain.frx":4E318
         Stretch         =   -1  'True
         Tag             =   "2"
         Top             =   1260
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   3
         Left            =   840
         Picture         =   "frmMain.frx":4EEFA
         Stretch         =   -1  'True
         Tag             =   "11"
         Top             =   1050
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   2
         Left            =   630
         Picture         =   "frmMain.frx":4F83C
         Stretch         =   -1  'True
         Tag             =   "11"
         Top             =   840
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   1
         Left            =   420
         Picture         =   "frmMain.frx":5003E
         Stretch         =   -1  'True
         Tag             =   "11"
         Top             =   525
         Width           =   1170
      End
      Begin VB.Image imgAllCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Index           =   0
         Left            =   210
         Picture         =   "frmMain.frx":50820
         Stretch         =   -1  'True
         Tag             =   "11"
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Label lblNOTE 
      Alignment       =   2  'Center
      Caption         =   "Many of the form cont- rols are hidden, beyong this form..>>..  expand the form to see all of them! Do not delete them!"
      Height          =   1380
      Left            =   5460
      TabIndex        =   17
      Top             =   4515
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblCash 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   420
      TabIndex        =   15
      ToolTipText     =   "Total time for which this game was played!!"
      Top             =   4200
      Width           =   645
   End
   Begin VB.Image imgAnim 
      Height          =   330
      Left            =   6825
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmMain.frx":511E2
      Stretch         =   -1  'True
      Top             =   210
      Width           =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      Index           =   1
      X1              =   210
      X2              =   7245
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   210
      X2              =   7245
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   2310
      Width           =   2010
   End
   Begin VB.Label lblTip 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Developed by M Zeeshan Alam, zeeshandj@yahoo.com... click to visit: ispro.netfirms.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   14
      ToolTipText     =   "Tip: Not always work correct!"
      Top             =   5985
      Width           =   7155
   End
   Begin VB.Label lblDealer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   2310
      TabIndex        =   13
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2415
      TabIndex        =   12
      Top             =   3150
      Width           =   1275
   End
   Begin VB.Label lblCash 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   11
      Top             =   3465
      Width           =   645
   End
   Begin VB.Label lblPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   2415
      TabIndex        =   10
      ToolTipText     =   "Current Score"
      Top             =   1785
      Width           =   330
   End
   Begin VB.Label lblPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   1
      Left            =   2415
      TabIndex        =   9
      ToolTipText     =   "Current Score"
      Top             =   4410
      Width           =   330
   End
   Begin VB.Image imgPlayer 
      Height          =   1695
      Index           =   0
      Left            =   2940
      Stretch         =   -1  'True
      Top             =   3570
      Width           =   1275
   End
   Begin VB.Image imgdealer 
      Height          =   1695
      Index           =   0
      Left            =   2940
      Stretch         =   -1  'True
      Top             =   1365
      Width           =   1275
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   9
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   5775
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      Height          =   1695
      Index           =   1
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   210
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   8
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   1470
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "10 player's image for holding cards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   8715
      TabIndex        =   4
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "10 dealer's image for holding cards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   8610
      TabIndex        =   3
      Top             =   0
      Width           =   1485
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1545
      Index           =   7
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   6
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   5
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   5250
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   4
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   4515
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   3
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Image imgPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Index           =   2
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   3465
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   9
      Left            =   8610
      Stretch         =   -1  'True
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   8
      Left            =   8610
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   4
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   1
      Left            =   7665
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   1065
   End
   Begin VB.Label Label3 
      Caption         =   "We are taking the maximum possibilities of hits that's why we are taking ten for each playuer otherwise 4-5 is enuf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   7560
      TabIndex        =   2
      Top             =   0
      Width           =   3270
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   7
      Left            =   9555
      Stretch         =   -1  'True
      Top             =   2730
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   6
      Left            =   9555
      Stretch         =   -1  'True
      Top             =   2310
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   5
      Left            =   9555
      Stretch         =   -1  'True
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   3
      Left            =   9555
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Image imgdealer 
      Height          =   1275
      Index           =   2
      Left            =   9765
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   0
      Left            =   420
      Top             =   3465
      Width           =   1800
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   435
      Index           =   0
      Left            =   315
      Top             =   3360
      Width           =   2010
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   1
      Left            =   420
      Top             =   4200
      Width           =   1800
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   435
      Index           =   1
      Left            =   315
      Top             =   4095
      Width           =   2010
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lets start by declarin a few global variables
'
'

Dim CardPoint(52) As Integer                'how much point each card owns
Dim Cash As Integer                               'money, how much player has in his pockets
Dim WhoseChance As Integer                  '0=dealer, 1=player

Dim DealerCurrentCards(10) As Integer            'per game per card points
Dim DealerCurrentCardsIndex(10) As Integer  'in hits these card should not be repeated coz they are already there

Dim PlayerCurrentCards(10) As Integer       'same two var. for player
Dim PlayerCurrentCardsIndex(10) As Integer

Dim playerCardCount As Integer          'how many cards the user has!
Dim dealerCardCount As Integer          'how many cards the dealer has!

Private Sub AssignCardPoints()

Dim i As Integer        'i = counter
For i = 0 To 51
    'we have scored the card points on each corrosponding image tag!
    CardPoint(i) = Val(imgAllCard(i).Tag)       'tough JOB! :-(
Next i
End Sub

Private Sub ccTitle_Click()
frmAbout.Show vbModal       'About The Game~!
End Sub

Private Sub cmdHelp_Click()
'MsgBox "Give your help here!!"
On Error Resume Next        'file not found!!
Shell "explorer.exe " & App.Path + "\readme.htm", vbNormalFocus
End Sub

Private Sub cmdHide_Click()
WindowState = vbMinimized
End Sub

Private Sub cmdHits_Click()
'add one more card to it
'*****************************************
'very much similar to Select2Card
'*****************************************
Dim i1 As Integer, i2 As Integer    '2 temp. variables.... :-)
Dim goAhead As Boolean ' heoy... used while checkin two same cards aren't drawn!
Dim iPoint As Integer
On Error Resume Next                    'no error handlin, its upto you!!
        goAhead = True
            Do
                i1 = Int(Rnd * 51)
                For i = 0 To 10     'check for all 10 player cards
                    If i1 = PlayerCurrentCardsIndex(i) Then _
                        goAhead = False Else goAhead = True
                Next i
            Loop While goAhead = False
            
            'hooray finally we have one card in our pocket
            
            imgPlayer(playerCardCount).Picture = imgAllCard(i1).Picture     'assign picture
             
             Sleep 100
            imgPlayer(playerCardCount).ZOrder 0      'on top of the frame! otherwise it is invisible! bad idea!!
            'for all cards except this first one, they must be displayed like real cards!
            
            If playerCardCount <> 0 Then        'here every player card count  will be greater than zero   ;-)
                imgPlayer(playerCardCount).Visible = False
                imgPlayer(playerCardCount).Left = (imgPlayer(playerCardCount - 1).Left + 200)  'Special FX
                imgPlayer(playerCardCount).Top = (imgPlayer(playerCardCount - 1).Top + 30)
               Debug.Print imgPlayer(playerCardCount).Top, imgPlayer(playerCardCount - 1).Top + 120
               imgPlayer(playerCardCount).Width = imgPlayer(playerCardCount - 1).Width
               imgPlayer(playerCardCount).Height = imgPlayer(playerCardCount - 1).Height
            
            End If
            'Call Animate(0, (playerCardCount))
            
            imgPlayer(playerCardCount).Visible = True
            
            PlayerCurrentCards(playerCardCount) = CardPoint(i1)     'remember we assigned this value in AssignCardPoints
            lblPoints(1) = lblPoints(1) + PlayerCurrentCards(playerCardCount)
            
            playerCardCount = playerCardCount + 1       'another card is done, prepare for the next one
            
            CheckAce (0)        'for player
            CheckGameLost

End Sub

Private Sub CheckGameLost()
'***************************************
'check whether player lost or won the game
'***************************************
'my kb is tough to work with! poor me :-((
'***************************************

If Val(lblPoints(1)) > 21 Then
        Sleep 1000
        ExMsgBox "Your current score is more than 21.", False
        Money = Money - 5
        Form_Load           'restart
ElseIf Val(lblPoints(1)) = 21 Then
        Sleep 1000
        ExMsgBox "You make an exact score of 21.", True
        Money = Money + 10
        Form_Load           'restart
End If

End Sub

Private Sub cmdNew_Click()
    'game_over
    StartTime = Time
    Money = 100
    Form_Load
    Form_Activate

End Sub

Private Sub cmdNewDeal_Click()
cmdNewDeal.Enabled = False
WhoseChance = 1

ReassignGameGlobals
playerCardCount = 0
dealerCardCount = 0
lblPoints(0) = 0
lblPoints(1) = 0
AssignCardPoints
Select2Card (0)         '0 for player
Select2Card (1)          '1 for dealer
Money = Money - 5      'each game takes 5
txtCash = Money
cmdHits.Enabled = True
cmdStand.Enabled = True
End Sub

Private Sub ReassignGameGlobals()
For i = 0 To 9
    PlayerCurrentCards(i) = -1                  '<> 0 coz zero is there already in index values!!
    PlayerCurrentCardsIndex(i) = -1
    DealerCurrentCards(i) = -1
    DealerCurrentCardsIndex(i) = -1
    imgdealer(i).Visible = False
    imgPlayer(i).Visible = False
Next i
playerCardCount = 0
dealerCardCount = 0
lblPoints(0) = 0
lblPoints(1) = 0
End Sub


Private Sub Select2Card(index As Integer)
'****************************************************
'This sub droved all the nuts of my mind out *** poor me ***
'****************************************************

Dim i1 As Integer, i2 As Integer    '2 temp. variables.... :-)
Dim goAhead As Boolean ' heoy... used while checkin two same cards aren't drawn!
Dim iPoint As Integer
Randomize
If index = 0 Then               '0=for player, (hero for me)
Do
        goAhead = True
            Do
                i1 = Int(Rnd * 51)
                For i = 0 To 10     'check for all 10 player cards
                    If i1 = PlayerCurrentCardsIndex(i) Then _
                        goAhead = False Else goAhead = True
                Next i
                
            Debug.Print "We are having :=", i1, playerCardCount
            Loop While goAhead = False
            'hooray finally we have one card in our pocket
            imgPlayer(playerCardCount).ZOrder 0      'on top of the frame! otherwise it is invisible! bad idea!!
            imgPlayer(playerCardCount).Visible = True
            imgPlayer(playerCardCount).Picture = imgAllCard(i1).Picture     'assign picture
            Debug.Print imgPlayer(playerCardCount).Left, imgPlayer(playerCardCount).Top, imgPlayer(playerCardCount).Visible
            'for all cards except this first one, they must be displayed like real cards!
            If playerCardCount <> 0 Then
                imgPlayer(playerCardCount).Left = (imgPlayer(playerCardCount - 1).Left + 300)  'Special FX
                imgPlayer(playerCardCount).Top = (imgPlayer(playerCardCount - 1).Top + 30)
               Debug.Print imgPlayer(playerCardCount).Top, imgPlayer(playerCardCount - 1).Top + 120
               imgPlayer(playerCardCount).Width = imgPlayer(playerCardCount - 1).Width
               imgPlayer(playerCardCount).Height = imgPlayer(playerCardCount - 1).Height
            End If
                        
            PlayerCurrentCards(playerCardCount) = CardPoint(i1)     'remember we assigned this value in AssignCardPoints
            lblPoints(1) = Val(lblPoints(1)) + PlayerCurrentCards(playerCardCount)
            
            playerCardCount = playerCardCount + 1       'another card is done, prepare for the next one
            
            CheckAce (0)        'check wht value of ace we need
            'if two card add to more than 21 we lost
            CheckGameLost

Loop While playerCardCount < 2       'this sub is only for the first two card so do both of them

ElseIf index = 1 Then '1 = no.1 for dealer i.e. computer
Do
        goAhead = True
            Do
                i1 = Int(Rnd * 51)
                For i = 0 To 10     'check for all 10 dealer cards
                    If i1 = DealerCurrentCardsIndex(i) Then _
                        goAhead = False Else goAhead = True
                Next i
                
            Debug.Print "We are having :=", i1, dealerCardCount
            Loop While goAhead = False
            'hooray finally we have one card in our pocket
            'imgdealer(dealerCardCount).ZOrder 0      'on top of the frame! otherwise it is invisible! bad idea!!
            imgdealer(dealerCardCount).Visible = True
            imgdealer(dealerCardCount).Picture = imgAllCard(i1).Picture     'assign picture
            Debug.Print imgdealer(dealerCardCount).Left, imgdealer(dealerCardCount).Top, imgdealer(dealerCardCount).Visible
            'for all cards except this first one, they must be displayed like real cards!
            If dealerCardCount <> 0 Then
                imgdealer(dealerCardCount).Visible = False
                imgdealer(dealerCardCount).Left = (imgdealer(dealerCardCount - 1).Left + 300)  'Special FX
                imgdealer(dealerCardCount).Top = (imgdealer(dealerCardCount - 1).Top + 30)
                imgdealer(dealerCardCount).Width = imgdealer(dealerCardCount - 1).Width
                imgdealer(dealerCardCount).Height = imgdealer(dealerCardCount - 1).Height
               Debug.Print imgdealer(dealerCardCount).Top, imgdealer(dealerCardCount - 1).Top + 120
            End If
                        
            DealerCurrentCards(dealerCardCount) = CardPoint(i1)     'remember we assigned this value in AssignCardPoints
            
            dealerCardCount = dealerCardCount + 1       'another card is done, prepare for the next one
                        
Loop While dealerCardCount < 2       'this sub is only for the first two card so do both of them

End If

End Sub


Private Sub cmdQuit_Click()
Dim ohNo As Integer
ohNo = MsgBox("Do you really wanna quit?", vbYesNo + vbDefaultButton2)
If ohNo = vbYes Then
                End
Else
Cancel = True
End If
End Sub



Private Sub cmdStand_Click()
'well..... here goes all the conditioning and computer A.I. thinkin'
'we could make computer to always win the game, but i m giving some mercy to
'the player, so leave it!!
'anyway... lets get down to business
'****************************************
'CHECKING FOR CASH MUST BE DONE OVER HERE!!
'*****************************************

'''''lock hit button
'
'show tht hidden card
Dim buff As String
On Error Resume Next
For i = 10 To 51
    'buff = buff & vbCrLf & imgAllCard(i).Tag
Next
'MsgBox buff

imgdealer(0).Visible = True
imgdealer(1).Visible = True
imgdealer(1).ZOrder

Randomize
cmdNewDeal.Enabled = True

For i = 0 To dealerCardCount - 1
        lblPoints(0) = Val(lblPoints(0)) + DealerCurrentCards(i)
Next i

CheckAce (1)

'Check Game Lost

If Val(lblPoints(0)) > 21 Then
        Sleep 1000
        ExMsgBox "Dealer's point is greater than 21, it is:" & lblPoints(0), True
        Money = Money + 10
        Form_Load           'restart
        Exit Sub
ElseIf Val(lblPoints(0)) = 21 Then
        Sleep 1000
        ExMsgBox "Dealer's point is equal to 21", False
        Money = Money - 5
        Form_Load           'restart
        Exit Sub
ElseIf Val(lblPoints(0).Caption) > Val(lblPoints(1).Caption) Then
        Sleep 1000
        ExMsgBox "Dealer's score " & lblPoints(0) & " is more than your points", False
        Money = Money - 5
        Form_Load           'restart
        Exit Sub
ElseIf Val(lblPoints(0).Caption) = Val(lblPoints(1).Caption) Then
        Sleep 1000
        ExMsgBox "It's a tie!", True
        Money = Money + 5
        Form_Load
        Exit Sub
End If

'though seems like computer didn't won, lets give it one more chance!!
i = 0
Randomize       'i wish there was randome-zee !
Do
rand = Int(Rnd * 10) + 1
If rand < 3 Then        'whether computer choose hits is chanced 2 in 10....!
    Hit4Dealer              'pull another card
    If Val(lblPoints(0)) > 21 Then Exit Sub
    'did we lose the game?
    If Val(lblPoints(0)) = 0 Then Exit Sub
End If
i = i + 1
Loop Until i > 5        'WE GIVE COMPUTER 5 CHANCES to pull another card!!

'final checking who won this black jack

If Val(lblPoints(0)) > Val(lblPoints(1)) Then
        Sleep 1000
        ExMsgBox "Dealer's score is greater than your score", False
        Money = Money - 5
        Form_Load           'restart
        Exit Sub
ElseIf Val(lblPoints(0).Caption) = Val(lblPoints(1).Caption) Then
        Sleep 1000
        ExMsgBox "It's a tie!", True
        Money = Money + 5
        Form_Load
        Exit Sub
Else
        Sleep 1000
        ExMsgBox "Your score is larger than dealer's score! Dealer score is:" & lblPoints(0), True
        Money = Money + 10
        Form_Load           'restart
        Exit Sub
End If


cmdNewDeal.Enabled = True
End Sub

Private Sub Form_Activate()
'humm...........!!
'well...

For i = 1 To imgPlayer.Count - 1
        imgPlayer(i).Width = imgPlayer(0).Width
        imgPlayer(i).Height = imgPlayer(0).Height
        imgdealer(i).Width = imgdealer(0).Width
        imgdealer(i).Height = imgdealer(0).Height
Next i
End Sub

Private Sub Form_Load()
cmdNewDeal.Enabled = True
cmdHits.Enabled = False
cmdStand.Enabled = False
txtCash.Text = Money
End Sub

Private Sub CheckAce(index As Integer)       'how the ekka (ace) must be used, as 1 or 12
If index = 0 Then   '0=player
    For i = 0 To playerCardCount - 1
        Debug.Print playerCardCount, PlayerCurrentCards(playerCardCount), i
        If PlayerCurrentCards(i) = 11 Then
                    If Val(lblPoints(1)) > 22 Then           'depends if 11 or 1 is needed
                            lblPoints(1) = lblPoints(1) - 10
                            PlayerCurrentCards(i) = 1
                    End If
        End If
    Next i
ElseIf index = 1 Then       'same for computer dealer
    For i = 0 To dealerCardCount - 1
        If DealerCurrentCards(i) = 11 Then
                If lblPoints(0) > 22 Then
                        lblPoints(0) = lblPoints(0) - 10
                End If
        End If
        
    Next i
End If
End Sub
'******************************************************************
'add one more card to it
'*****************************************
'very much similar to Select2Card
'*****************************************
Private Sub Hit4Dealer()
Dim i1 As Integer, i2 As Integer    '2 temp. variables.... :-)
Dim goAhead As Boolean ' heoy... used while checkin two same cards aren't drawn!
Dim iPoint As Integer
On Error Resume Next                    'no error handlin, its upto you!!
        goAhead = True
            Do
                i1 = Int(Rnd * 51)
                For i = 0 To 10     'check for all 10 dealer cards
                    If i1 = DealerCurrentCardsIndex(i) Then _
                        goAhead = False Else goAhead = True
                Next i
            Loop While goAhead = False
            
            'hooray finally we have one card in our pocket
            
            imgdealer(dealerCardCount).ZOrder 0      'on top of the frame! otherwise it is invisible! bad idea!!
            imgdealer(dealerCardCount).Visible = True
            imgdealer(dealerCardCount).Picture = imgAllCard(i1).Picture     'assign picture
            
            Debug.Print imgdealer(dealerCardCount).Left, imgdealer(dealerCardCount).Top, imgdealer(dealerCardCount).Visible
            
            'for all cards except this first one, they must be displayed like real cards!
            
            If dealerCardCount <> 0 Then        'here every dealer card count  will be greater than zero   ;-)
            
                imgdealer(dealerCardCount).Left = (imgdealer(dealerCardCount - 1).Left + 300)  'Special FX
                imgdealer(dealerCardCount).Top = (imgdealer(dealerCardCount - 1).Top + 30)
               Debug.Print imgdealer(dealerCardCount).Top, imgdealer(dealerCardCount - 1).Top + 120
               imgdealer(dealerCardCount).Width = imgdealer(dealerCardCount - 1).Width
               imgdealer(dealerCardCount).Height = imgdealer(dealerCardCount - 1).Height
            
            End If
                        
            DealerCurrentCards(dealerCardCount) = CardPoint(i1)     'remember we assigned this value in AssignCardPoints
            lblPoints(0) = lblPoints(0) + DealerCurrentCards(dealerCardCount)
            
            dealerCardCount = dealerCardCount + 1       'another card is done, prepare for the next one
            
            CheckAce (1)        'for dealer
            'CheckGameLost
            
            If Val(lblPoints(0)) > 21 Then
                    Sleep 1000
                    ExMsgBox "Dealer score greater than 21", True
                    Money = Money + 10
                    Form_Load
                    'restart
                    Exit Sub
            End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'move the form ;-)
    If Button = vbLeftButton Then
        Call ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
End Sub

Private Sub Form_Paint()
txtCash = Money
End Sub

Private Sub lblTip_Click()
On Error Resume Next
Shell "start http://ispro.netfirms.com", vbNormalFocus
End Sub

Private Sub tmrEffect_Timer()
Dim tip As String
Dim l, r As String
'------------------------------------
tip = ccTitle.Caption
l = Left(tip, 1)
r = Right(tip, Len(tip) - 1)
ccTitle.Caption = r + l

txttime.Text = Format(Time - StartTime, "hh:mm:ss")
End Sub

Private Sub tmrGiveTip_Timer()
'see wht's goin on in this game and give an A.I. help
If cmdNewDeal.Enabled = True And Money >= 10 Then
        lblTip = "Click on Deal New Hand to play game...  "
ElseIf Val(lblPoints(1)) < 15 Then
        lblTip = "Try your luck.. mau be you can score better, click HIT  "
ElseIf Val(lblPoints(1)) > 15 Then
        lblTip = "Click STAND clicking HIT may be risky   "
    
Else
        lblTip = "Email the developer at zeeshandj@yahoo.com  "
        'ccTitle.ForeColor = QBColor(Int(Rnd * 15) + 1)      'this effect comes very lately...! why? just i like so...!!! >:-)

End If

End Sub

Private Sub txtCash_Change()
If Val(txtCash.Text) < 0 Then
    Sleep 1000
    MsgBox "All of your balance is lost! You can play no more! Click ok restart a new game"
    Form_Load
    Money = 100
End If
End Sub

Private Sub ExMsgBox(Des As String, Win As Boolean)
If Win = True Then 'Hooray, we win!!
        Dialog.lblDescription = Des + vbCrLf + " Your points is " & lblPoints(1)
        Dialog.lblTitle = "Congatulations, you win!"
        Dialog.lblTip = "Hooray for you, seems like you are a good player," + vbCrLf + _
                                "Keep it up. Lets defeat the dealer once more!" + vbCrLf + _
                                "Register for the complete version of Zee Black Jack!"
        Dialog.imgwin.Visible = True
        Dialog.imgLost.Visible = False
ElseIf Win = False Then
        Dialog.lblDescription = Des + vbCrLf + " Your points is " & lblPoints(1)
        Dialog.lblTitle = "You Lost!"
        Dialog.lblTip = "Are you new to Zee Black Jack?" + vbCrLf + _
                                "Read the manual, or click Help to see online version of the manual" + vbCrLf + _
                                "You didn't play well, Try again!!"
        Dialog.imgwin.Visible = False
        Dialog.imgLost.Visible = True
End If
Dialog.Show vbModal
End Sub
