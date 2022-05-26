VERSION 5.00
Begin VB.Form Tela_Etiqueta 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   54.504
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   144.727
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FR_Etiqueta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1757
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5046
      Begin VB.PictureBox PIC_CodigoBarra 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1080
         ScaleHeight     =   615
         ScaleWidth      =   2895
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label LB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Etiqueta de Rastreabilidade de Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label LB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "CONESTEEL VÁLVULAS E CONEXÕES INDUSTRIAIS LTDA."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   4500
      End
   End
End
Attribute VB_Name = "Tela_Etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
