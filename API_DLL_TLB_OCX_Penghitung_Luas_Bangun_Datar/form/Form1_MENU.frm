VERSION 5.00
Begin VB.Form Form1_MENU 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "Form1_MENU.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEVAN CAKRA MUDRA WIJAYA"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   8760
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MATA KULIAH PEMROGRAMAN API - A"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   8280
      Width           =   10575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT @ 2020  |  LUAS BANGUN DATAR ALL RIGHT RESERVED"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   7680
      Width           =   12015
   End
   Begin VB.Menu menu 
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
      Begin VB.Menu ocxLJG 
         Caption         =   "OCX Luas Jajaran Genjang"
         Index           =   0
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
      Index           =   0
   End
End
Attribute VB_Name = "Form1_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LJG_Click(Index As Integer)
    Form2_LJG.Show
End Sub

Private Sub LBK_Click(Index As Integer)
    Form3_LBK.Show
End Sub

Private Sub LSL_Click(Index As Integer)
    Form4_LSL.Show
End Sub

Private Sub exit_Click(Index As Integer)
    Q = MsgBox("Anda yakin akan keluar", vbQuestion + vbOKCancel, "Informasi")
    If Q = vbOK Then
        End
    End If
End Sub

Private Sub ocxLJG_Click(Index As Integer)
    Form5_OCX.Show
End Sub
