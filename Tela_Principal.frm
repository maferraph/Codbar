VERSION 5.00
Begin VB.Form Tela_Principal 
   AutoRedraw      =   -1  'True
   Caption         =   "Codigo de Barra"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   4905
      TabIndex        =   5
      Top             =   1320
      Width           =   4935
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Grande"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Medio"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Pequeno"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Sair"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Informe o texto para o codigo de barras"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo de barras copiado para o clipboard"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3000
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()  ' codigo do botão Sair
   End
End Sub

Private Sub cmdPrint_Click()  'codigo do botão imprimir

'Printer.NewPage
Printer.PaintPicture Picture1, 5000, 5000
Printer.EndDoc

End Sub


Private Sub Command1_Click()
    Call DrawBarcode(Text1, Tela_Etiqueta.PIC_CodigoBarra)
    Tela_Etiqueta.PrintForm
End Sub

Private Sub Form_Activate()  ' código do evento activate do formulário
  optSize(1) = 1
End Sub


Private Sub optSize_Click(Index As Integer)  ' codigo dos botões de opções
Picture1.ScaleMode = 3

Select Case Index
  Case 0
      Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
      Picture1.FontSize = 8
  Case 1
      Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
      Picture1.FontSize = 10
  Case 2
      Picture1.Height = Picture1.Height * (3 * 40 / Picture1.ScaleHeight)
      Picture1.FontSize = 14
End Select


Call Text1_Change

End Sub

Private Sub Text1_Change()  'codigo do evento change da caixa de texto

Call DrawBarcode(Text1, Picture1)

  MinWidth = 2 * Text1.Left + Text1.Width
  pw = 2 * Picture1.Left + Picture1.Width
  fw = MinWidth
  If pw > fw Then fw = pw
  Form1.Width = fw

End Sub


