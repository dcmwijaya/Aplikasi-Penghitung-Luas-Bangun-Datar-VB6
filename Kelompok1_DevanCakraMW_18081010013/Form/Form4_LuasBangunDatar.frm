VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Luas Setengah Lingkaran"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   Picture         =   "Form4_LuasBangunDatar.frx":0000
   ScaleHeight     =   9915
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "[ Program Luas Setengah Lingkaran ]"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   4680
      TabIndex        =   0
      Top             =   2280
      Width           =   10215
      Begin VB.TextBox Hasil 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   5160
         TabIndex        =   7
         Top             =   2640
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Hitung"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox r 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   480
         Picture         =   "Form4_LuasBangunDatar.frx":8216B
         ScaleHeight     =   3975
         ScaleWidth      =   4335
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Perhitungan"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   3
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jari-jari (r)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Menu HU 
      Caption         =   "Halaman Utama"
      Index           =   0
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
      Index           =   0
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim io As New RumusHitung

Private Sub Command1_Click()
    Hasil = io.LSL(Val(r))
End Sub

Private Sub Command2_Click()
    r = ""
    Hasil = ""
    r.SetFocus
End Sub

Private Sub HU_Click(Index As Integer)
    r = ""
    Hasil = ""
    Form1.Show
End Sub

Private Sub Exit_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar", vbQuestion + vbOKCancel, "Informasi")
    If Q = vbOK Then
    End
    End If
End Sub
