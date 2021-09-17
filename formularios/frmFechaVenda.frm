VERSION 5.00
Begin VB.Form frmFechaVenda 
   Caption         =   "Fechamento de Venda"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3840
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.Frame fraPagar 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2100
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3405
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtDinheiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtTicket 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSubDinheiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSubCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSubCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSubTicket 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label labTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   0
         TabIndex        =   17
         Top             =   30
         Width           =   405
      End
      Begin VB.Label labDinheiro 
         AutoSize        =   -1  'True
         Caption         =   "Dinheiro:"
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   390
         Width           =   630
      End
      Begin VB.Label labCheque 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   750
         Width           =   600
      End
      Begin VB.Label labCartao 
         AutoSize        =   -1  'True
         Caption         =   "Cartão:"
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   1110
         Width           =   510
      End
      Begin VB.Label labTicket 
         AutoSize        =   -1  'True
         Caption         =   "Ticket:"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label labDesconto 
         AutoSize        =   -1  'True
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label labSubTotal 
         AutoSize        =   -1  'True
         Caption         =   "SubTotal:"
         Height          =   195
         Left            =   2040
         TabIndex        =   11
         Top             =   120
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmFechaVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCartao_GotFocus()
    txtCartao.SelStart = 0
    txtCartao.SelLength = 10
    txtCartao.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtCartao_KeyPress(KeyAscii As Integer)
    'KeyAscii = Pagar(KeyAscii, txtCartao, txtSubCartao)
End Sub

Private Sub txtCartao_LostFocus()
    txtCartao.Text = modMoeda.FmtMoeda(txtCartao.Text)
    txtCartao.BackColor = &H80000005 'Branco
End Sub

Private Sub txtCheque_GotFocus()
    txtCheque.SelStart = 0
    txtCheque.SelLength = 10
    txtCheque.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
    'KeyAscii = Pagar(KeyAscii, txtCheque, txtSubCheque)
End Sub

Private Sub txtCheque_LostFocus()
    txtCheque.Text = modMoeda.FmtMoeda(txtCheque.Text)
    txtCheque.BackColor = &H80000005 'Branco
End Sub

Private Sub txtDesconto_GotFocus()
    txtDesconto.SelStart = 0
    txtDesconto.SelLength = 10
    txtDesconto.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
    txtDinheiro.Text = Empty
    Call txtDinheiro_KeyPress(KeyAscii)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtDesconto_LostFocus()
    txtDesconto.Text = modMoeda.FmtMoeda(txtDesconto.Text)
    txtDesconto.BackColor = &H80000005 'Branco
End Sub

Private Sub txtDinheiro_GotFocus()
    txtDinheiro.SelStart = 0
    txtDinheiro.SelLength = 10
    txtDinheiro.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtDinheiro_KeyPress(KeyAscii As Integer)
    'KeyAscii = Pagar(KeyAscii, txtDinheiro, txtSubDinheiro)
End Sub

Private Sub txtDinheiro_LostFocus()
    txtDinheiro.Text = modMoeda.FmtMoeda(txtDinheiro.Text)
    txtDinheiro.BackColor = &H80000005 'Branco
End Sub

Private Sub txtTicket_GotFocus()
    txtTicket.SelStart = 0
    txtTicket.SelLength = 10
    txtTicket.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtTicket_KeyPress(KeyAscii As Integer)
    'KeyAscii = Pagar(KeyAscii, txtTicket, txtSubTicket)
End Sub

Private Sub txtTicket_LostFocus()
    txtTicket.Text = modMoeda.FmtMoeda(txtTicket.Text)
    txtTicket.BackColor = &H80000005 'Branco
End Sub

Private Sub txtTotal_GotFocus()
    txtTotal.SelStart = 0
    txtTotal.SelLength = 10
    txtTotal.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtTotal_LostFocus()
    txtTotal.Text = modMoeda.FmtMoeda(txtTotal.Text)
    txtTotal.BackColor = &H80000005 'Branco
End Sub

