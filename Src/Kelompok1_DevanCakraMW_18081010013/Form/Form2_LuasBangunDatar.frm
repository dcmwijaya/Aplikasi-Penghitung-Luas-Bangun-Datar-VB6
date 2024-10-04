VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Luas Jajaran Genjang"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "Form2_LuasBangunDatar.frx":0000
   ScaleHeight     =   9915
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "[ Program Luas Jajaran Genjang ]"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   2400
      Width           =   10215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   480
         Picture         =   "Form2_LuasBangunDatar.frx":8216B
         ScaleHeight     =   3975
         ScaleWidth      =   4335
         TabIndex        =   9
         Top             =   600
         Width           =   4335
      End
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
         Height          =   1215
         Left            =   5160
         TabIndex        =   8
         Top             =   3360
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox t 
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
         Left            =   6120
         TabIndex        =   4
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox a 
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
         Left            =   6120
         TabIndex        =   1
         Top             =   600
         Width           =   3615
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
         TabIndex        =   7
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tinggi (t)"
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
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Alas (a)"
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
         Left            =   5280
         TabIndex        =   2
         Top             =   720
         Width           =   1215
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim io As New RumusHitung

Private Sub Command1_Click()
    Hasil = io.LJG(Val(a), Val(t))
End Sub

Private Sub Command2_Click()
    a = ""
    t = ""
    Hasil = ""
    a.SetFocus
End Sub

Private Sub HU_Click(Index As Integer)
    a = ""
    t = ""
    Hasil = ""
    Form1.Show
End Sub

Private Sub Exit_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar", vbQuestion + vbOKCancel, "Informasi")
    If Q = vbOK Then
    End
    End If
End Sub


