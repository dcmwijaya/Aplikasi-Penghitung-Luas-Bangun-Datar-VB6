VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Halaman Utama Aplikasi Penghitung Luas Bangun Datar"
   ClientHeight    =   8760
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   Picture         =   "Form1_LuasBangunDatar.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   16725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Kelompok 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KELOMPOK 1 MATA KULIAH PEMROGRAMAN API-A"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   2
      Top             =   8520
      Width           =   5895
   End
   Begin VB.Label Copyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright @2020 | Luas Bangun Datar All Right Reserved"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   7920
      Width           =   9135
   End
   Begin VB.Label Creator 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEVAN CAKRA MUDRA WIJAYA (18081010013)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5880
      TabIndex        =   0
      Top             =   9000
      Width           =   7695
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Index           =   0
      Begin VB.Menu LJG 
         Caption         =   "Luas Jajaran Genjang"
         Index           =   0
      End
      Begin VB.Menu LBK 
         Caption         =   "Luas Belah Ketupat"
         Index           =   0
      End
      Begin VB.Menu LSL 
         Caption         =   "Luas Setengah Lingkaran"
         Index           =   0
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
      Index           =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar", vbQuestion + vbOKCancel, "Informasi")
    If Q = vbOK Then
    End
    End If
End Sub

Private Sub LJG_Click(Index As Integer)
    Form2.Show
End Sub

Private Sub LBK_Click(Index As Integer)
    Form3.Show
End Sub

Private Sub LSL_Click(Index As Integer)
    Form4.Show
End Sub
