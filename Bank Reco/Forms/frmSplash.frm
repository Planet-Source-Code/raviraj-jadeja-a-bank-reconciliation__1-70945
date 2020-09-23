VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   2745
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer TimerPb 
      Interval        =   50
      Left            =   3600
      Top             =   2160
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   3240
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3975
         TabIndex        =   10
         Top             =   120
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http:\\www.easyinfosoft.110mb.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   2685
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visit :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Easy! Infosoft"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000040C0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1455
      Left            =   5400
      Top             =   -120
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000040C0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   375
      Left            =   480
      Top             =   960
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000040C0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000040C0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1455
      Left            =   0
      Top             =   -120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Bank Reconciliation System"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Easy! Bank Reco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   660
      TabIndex        =   6
      Top             =   435
      Width           =   4575
   End
   Begin VB.Label Label13 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1425
      Width           =   5895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Me.MousePointer = vbDefault
End Sub

Private Sub TimerPb_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 1
    Select Case ProgressBar1.Value

    Case 20
    Label5.Caption = "Please Wait."
    
    Case 40
    Label5.Caption = "Please Wait.."
    
    Case 60
    Label5.Caption = "Please Wait..."
    
    Case 80
    Label5.Caption = "Please Wait...."
    
    End Select
    
    If ProgressBar1.Value = ProgressBar1.Max Then
    Unload Me
    MDIForm1.Show
    End If
End Sub
