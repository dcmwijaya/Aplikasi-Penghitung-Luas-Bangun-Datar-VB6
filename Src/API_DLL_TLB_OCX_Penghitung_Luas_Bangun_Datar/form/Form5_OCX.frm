VERSION 5.00
Object = "{495F1F62-ECA2-4571-ABFD-1759FFC8EFF7}#1.0#0"; "Project1_Control.ocx"
Begin VB.Form Form5_OCX 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox HitungOCX1 
      Height          =   10215
      Left            =   0
      ScaleHeight     =   10155
      ScaleWidth      =   18915
      TabIndex        =   0
      Top             =   0
      Width           =   18975
      Begin Project1_Control.OCX_Class OCX_Class1 
         Height          =   10215
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   18975
         _ExtentX        =   33470
         _ExtentY        =   18018
      End
   End
End
Attribute VB_Name = "Form5_OCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
