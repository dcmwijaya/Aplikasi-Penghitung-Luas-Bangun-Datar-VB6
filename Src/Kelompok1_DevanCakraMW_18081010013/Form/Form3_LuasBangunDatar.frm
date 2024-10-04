VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Luas Belah Ketupat"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   Picture         =   "Form3_LuasBangunDatar.frx":0000
   ScaleHeight     =   9915
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "[ Program Luas Belah Ketupat ]"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   2160
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
         Height          =   1335
         Left            =   5160
         TabIndex        =   9
         Top             =   3240
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
         TabIndex        =   8
         Top             =   2040
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox d2 
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
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox d1 
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
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   480
         Picture         =   "Form3_LuasBangunDatar.frx":8216B
         ScaleHeight     =   3975
         ScaleWidth      =   4335
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagonal (d1)"
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
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagonal (d2)"
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
         Left            =   5160
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
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
         Top             =   3000
         Width           =   3615
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim io As New RumusHitung
    
Private Sub Command1_Click()
    Hasil = io.LBK(Val(d1), Val(d2))
End Sub

Private Sub Command2_Click()
    d1 = ""
    d2 = ""
    Hasil = ""
    d1.SetFocus
End Sub

Private Sub HU_Click(Index As Integer)
    d1 = ""
    d2 = ""
    Hasil = ""
    Form1.Show
End Sub

Private Sub Exit_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar", vbQuestion + vbOKCancel, "Informasi")
    If Q = vbOK Then
    End
    End If
End Sub

