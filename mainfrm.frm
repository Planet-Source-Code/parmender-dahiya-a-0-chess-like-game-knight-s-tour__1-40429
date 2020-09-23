VERSION 5.00
Begin VB.Form mainfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KNIGHT's TOUR"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox hintpict 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Left            =   6900
      Picture         =   "mainfrm.frx":0E42
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   84
      Top             =   5475
      Visible         =   0   'False
      Width           =   520
   End
   Begin VB.CommandButton cmdhow 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6900
      TabIndex        =   0
      ToolTipText     =   "How to Play KNIGHT'S TOUR"
      Top             =   3225
      Width           =   990
   End
   Begin VB.CommandButton cmdhint 
      Caption         =   "Show Possible Moves"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5550
      TabIndex        =   2
      ToolTipText     =   "Show all possible moves from current position"
      Top             =   2700
      Width           =   2340
   End
   Begin VB.CommandButton cmdpause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5550
      TabIndex        =   1
      ToolTipText     =   "Pause/Resume Game"
      Top             =   3225
      Width           =   990
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6225
      Top             =   150
   End
   Begin VB.PictureBox emptypict 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Left            =   7500
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   520
   End
   Begin VB.Frame Frame1 
      Caption         =   "Legengs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   5550
      TabIndex        =   75
      Top             =   3750
      Width           =   2340
      Begin VB.PictureBox fillpict 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   150
         Picture         =   "mainfrm.frx":1154
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   77
         TabStop         =   0   'False
         ToolTipText     =   "This picture shows covered position"
         Top             =   825
         Width           =   520
      End
      Begin VB.PictureBox mainpict 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   150
         Picture         =   "mainfrm.frx":145E
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "This picture shows current position"
         Top             =   225
         Width           =   520
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Covered Position"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   750
         TabIndex        =   79
         ToolTipText     =   "This picture shows covered moves"
         Top             =   975
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Current Position"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   750
         TabIndex        =   78
         Top             =   450
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6900
      TabIndex        =   4
      ToolTipText     =   "About Author"
      Top             =   1350
      Width           =   990
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5475
      TabIndex        =   3
      ToolTipText     =   "One Move back"
      Top             =   1350
      Width           =   990
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6900
      TabIndex        =   6
      ToolTipText     =   "Quit Game"
      Top             =   825
      Width           =   990
   End
   Begin VB.CommandButton cmdagain 
      Caption         =   "Retry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5475
      TabIndex        =   5
      ToolTipText     =   "Restart Game"
      Top             =   825
      Width           =   990
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   63
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   62
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   61
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   60
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   59
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   58
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   57
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   56
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4575
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   55
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   54
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   53
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   52
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   51
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   50
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   49
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   48
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4050
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   47
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   46
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   45
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   44
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   43
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   42
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   41
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   40
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3525
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   39
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   38
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   37
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   36
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   35
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   34
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   33
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   32
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3000
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   31
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   30
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   29
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   28
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   27
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   26
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   25
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   24
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2475
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   23
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   22
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   21
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   20
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   19
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   18
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   17
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   16
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1950
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   15
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   14
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   13
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   12
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   11
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   10
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   9
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   8
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1425
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   7
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   6
      Left            =   3900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   5
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   4
      Left            =   2850
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   3
      Left            =   2325
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   2
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   1
      Left            =   1275
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.PictureBox board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   520
      Index           =   0
      Left            =   750
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   900
      Width           =   520
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5550
      TabIndex        =   83
      Top             =   2250
      Width           =   675
   End
   Begin VB.Label lblscore 
      AutoSize        =   -1  'True
      Caption         =   "64 squares left"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   6450
      TabIndex        =   82
      ToolTipText     =   "Current score"
      Top             =   2290
      Width           =   1470
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "KNIGHT'S TOUR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   570
      Left            =   2625
      TabIndex        =   80
      Top             =   75
      Width           =   3495
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      Caption         =   "0 Seconds"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   6825
      TabIndex        =   74
      ToolTipText     =   "Time elapsed since start of game"
      Top             =   1915
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5550
      TabIndex        =   73
      Top             =   1875
      Width           =   720
   End
   Begin VB.Label lbldir 
      AutoSize        =   -1  'True
      Caption         =   "Please select a square on the board to start."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   675
      TabIndex        =   72
      ToolTipText     =   "Direction for player"
      Top             =   5775
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Direction:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   675
      TabIndex        =   71
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   675
      X2              =   5025
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   675
      X2              =   5025
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5025
      X2              =   5025
      Y1              =   825
      Y2              =   5175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   675
      X2              =   675
      Y1              =   825
      Y2              =   5175
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub board_Click(Index As Integer)
current = Index

If cmdpause.Caption = "Resume" Then
    MsgBox "Game has been paused. Click 'Resume' to play.", vbOKOnly + vbExclamation, "KNIGHT'S TOUR"
    Exit Sub
End If

If start = False Then   ' For the first time to select the initial position
    If MsgBox("Are you sure to select this starting position ?" & vbCrLf & _
               "You can not unselect it even by 'Back' button.", vbYesNo + vbQuestion, "KNIGHT'S TOUR") = vbNo Then Exit Sub
    board(Index).Picture = mainpict.Picture   ' show picture
    start = True
    lbldir.Caption = "Please select a square to move to."
    score = score - 1
    lblscore.Caption = score & " squares left"
    back(64 - score) = current   ' store current move
    showmoves (current)
    cmdpause.Enabled = True
    cmdhint.Enabled = True
    previous = current
    Timer1.Enabled = True
    Exit Sub
End If

If current = previous Then Exit Sub  ' if clicked on the same square then do nothing

Call index_to_pos(current)

If status(row, col) = False Then
    MsgBox "This square has already been covered.", vbOKOnly + vbExclamation, "KNIGHT'S TOUR"
    Exit Sub
End If

' check if move is valid or not
validmove = False
For i = 1 To 8
    Call index_to_pos(current)
    If hintrow(i) = row And hintcol(i) = col Then
        validmove = True
        lbldir.Caption = "Please select a square to move to."
        hintpict.Visible = False
        updboard
    If score = 0 Then
        Timer1.Enabled = False
        If MsgBox("Great!!! You are a GENIUS." & vbCrLf & _
                  "You have covered all the squares." & vbCrLf & _
                  "To PLAY AGAIN click 'Yes', to QUIT click 'No'", vbYesNo + vbExclamation, "KNIGHT'S TOUR") = vbYes Then
            Call again
            Exit Sub
        Else
            End
        End If
    End If
        If showmoves(current) = False Then
            Timer1.Enabled = False
            If MsgBox("GAME OVER." & vbCrLf & _
                   "There are still " & score & " squares left." & vbCrLf & _
                   "To TRY AGAIN clicl 'Yes', to QUIT click 'No'", vbYesNo + vbExclamation, "KNIGHT'S TOUR") = vbYes Then
                Call again
            Else
                End
            End If
        End If
        Exit For
    End If
Next i

If validmove = True Then
    previous = current
Else
    MsgBox "Invalid Move.", vbOKOnly + vbExclamation, "KNIGHT'S TOUR"
End If
End Sub

Private Sub cmdabout_Click()
MsgBox "KNIGHT'S TOUR - Ver 1.0" & vbCrLf & _
       "Made by PARMENDER DAHIYA " & vbCrLf & _
       "TATA CONSULTANCY SERVICES, " & vbCrLf & _
       "New Delhi, INDIA." & vbCrLf & _
       "For solution mail to : ps_dahiya@yahoo.com ", vbOKOnly + vbExclamation, "KNIGHT'S TOUR - About"

End Sub

Private Sub cmdagain_Click()
If start = False Then
    MsgBox "You are already at the start of game.", vbOKOnly + vbExclamation, "KNIGHT'S TOUR"
    Exit Sub
Else
    
    If MsgBox("This game will end." & vbCrLf & _
              "Are you sure to restart the game ?", vbYesNo + vbExclamation, "KNIGHT'S TOUR") = vbYes Then
    Call again
    End If
End If

End Sub

Private Sub cmdback_Click()
score = score + 1
lblscore = score & " squares left"
Call index_to_pos(back(64 - score))

status(row, col) = True

For i = 1 To 8
    For j = 1 To 8
        If status(i, j) = False Then
            mainfrm.board(pos_to_index(i, j)).Picture = mainfrm.fillpict.Picture
        Else
            mainfrm.board(pos_to_index(i, j)).Picture = mainfrm.emptypict.Picture
        End If
    Next j
Next i

mainfrm.board(back(64 - score)).Picture = mainfrm.mainpict.Picture
current = back(64 - score)
previous = back(64 - score - 1)
showmoves (back(64 - score))

If score = 63 Then
    cmdback.Enabled = False
    previous = current
    Exit Sub
End If

End Sub

Private Sub cmdhint_Click()
For i = 1 To 8
    If valid(i) = True Then board(pos_to_index(hintrow(i), hintcol(i))).Picture = hintpict.Picture
Next i
lbldir.Caption = "Please select a square to move from possible moves shown by"
hintpict.Visible = True
End Sub

Private Sub cmdhow_Click()
MsgBox "1. First select the starting position on the board." & vbCrLf & _
       "2. Then as per rules of moving KNIGHT of Chess, move" & vbCrLf & _
       "     the KNIGHT from all the possible moves shown." & vbCrLf & _
       "3. After your selection that square will be marked covered." & vbCrLf & _
       "4. You can not go on a covered square." & vbCrLf & _
       "5. In the same way keep selecting the squares from" & vbCrLf & _
       "     the possible moves shown and try to cover all" & vbCrLf & _
       "     the squares." & vbCrLf & _
       " " & vbCrLf & _
       "GOOD LUCK and KEEP PLAYING.", vbOKOnly + vbExclamation, "KNIGHT'S TOUR - How to Play"
End Sub

Private Sub cmdpause_Click()
If cmdpause.Caption = "Pause" Then
    cmdpause.Caption = "Resume"
    Timer1.Enabled = False
    cmdback.Enabled = False
    Me.Caption = "KNIGHT'S TOUR - PAUSED"
Else
    cmdpause.Caption = "Pause"
    Timer1.Enabled = True
    cmdback.Enabled = True
    Me.Caption = "KNIGHT'S TOUR"
End If
End Sub

Private Sub cmdquit_Click()
If MsgBox("Are you sure to quit KNIGHT'S TOUR", vbYesNo + vbQuestion, "KNIGHT'S TOUR") = vbYes Then End
End Sub

Private Sub Form_Load()
start = False
score = 64
tottime = 0
current = 100
previous = 100
For i = 1 To 8
    For j = 1 To 8
        status(i, j) = True
    Next j
Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure to quit KNIGHT'S TOUR", vbYesNo + vbQuestion, "KNIGHT'S TOUR") = vbYes Then
    End
Else
    Cancel = True
End If
End Sub

Private Sub Timer1_Timer()
tottime = tottime + 1
lbltime.Caption = tottime & " Seconds"
If tottime - (tottime \ 300) * 300 = 0 Then
    MsgBox "This is to remind you that Max. Time allowed is 1 hour" & vbCrLf & _
           " And " & score & " squares are still remaining.", vbOKOnly + vbExclamation, "KNIGHT'S TOUR - Timer Alert"
End If

If tottime >= 3600 Then
    If MsgBox("TIME OVER." & vbCrLf & _
           "You are playing this game for past one hour." & vbCrLf & _
           "There are still " & score & " squares remaining." & vbCrLf & _
           "To TRY AGAIN click 'Yes', to QUIT click 'No'", vbYesNo + vbExclamation, "KNIGHT'S TOUR - Timer Alert") = vbYes Then
        Call again
    Else
        End
    End If
End If

End Sub
