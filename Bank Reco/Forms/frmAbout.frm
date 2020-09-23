VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy! Infosoft"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4110
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1440
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3720
      Top             =   720
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   3855
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   3615
         TabIndex        =   11
         Top             =   240
         Width           =   3615
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Graphics Designer : Raviraj Jadeja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   840
            TabIndex        =   15
            Top             =   1920
            Width           =   2475
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Database Designer : Raviraj Jadeja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   840
            TabIndex        =   14
            Top             =   1320
            Width           =   2550
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "System Analyst : Raviraj Jadeja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   840
            TabIndex        =   13
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Programmer : Raviraj Jadeja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   840
            TabIndex        =   12
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   135
         TabIndex        =   16
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label LBLRG 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "www.easyinfosoft.110mb.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1080
         MouseIcon       =   "frmAbout.frx":06EA
         TabIndex        =   9
         Top             =   1035
         Width           =   2205
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Website : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1030
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Developed By : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   5
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "E Mail : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   6
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Contact : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   5
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Easy! Infosoft"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "93775 99995 / 97250 19858"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "easyinfosoft@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   1080
         MouseIcon       =   "frmAbout.frx":09F4
         TabIndex        =   2
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Developers Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or 13 Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
End Sub

Private Sub Timer1_Timer()
    If Label3.Top <= -800 Then Label3.Top = 1560
    If Label4.Top <= -800 Then Label4.Top = 1560
    If Label5.Top <= -800 Then Label5.Top = 1560
    If Label6.Top <= -800 Then Label6.Top = 1560
                
    Label3.Top = Label3.Top - 15
    Label4.Top = Label4.Top - 15
    Label5.Top = Label5.Top - 15
    Label6.Top = Label6.Top - 15
End Sub
