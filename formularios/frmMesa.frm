VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Mesas"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9870
   Icon            =   "frmMesa.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCores 
      Interval        =   20000
      Left            =   6600
      Top             =   5760
   End
   Begin MSComctlLib.ImageList imlMesaD 
      Left            =   4920
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":0CCA
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":19A4
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":267E
            Key             =   "transferir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":3358
            Key             =   "estornar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":4032
            Key             =   "fechar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":4D0C
            Key             =   "pagar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":55E6
            Key             =   "sair"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrMesa 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6120
      Top             =   5760
   End
   Begin MSComctlLib.ImageList imlMesas 
      Left            =   4320
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":62C0
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":6F9A
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":7C74
            Key             =   "transferir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":894E
            Key             =   "estornar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":9628
            Key             =   "fechar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":A302
            Key             =   "pagar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":ABDC
            Key             =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlListView 
      Left            =   5520
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":B8B6
            Key             =   "amarela"
            Object.Tag             =   "amarela"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":BC50
            Key             =   "verde"
            Object.Tag             =   "verde"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":BFEA
            Key             =   "vermelha"
            Object.Tag             =   "vermelha"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesa.frx":C384
            Key             =   "azul"
            Object.Tag             =   "azul"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraProduto 
      Caption         =   "Produtos"
      Height          =   1935
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   9615
      Begin VB.TextBox txtCodPro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtDescPro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtQtdePro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtVlrPro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid mfgMesas 
         Height          =   1575
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin VB.Label labPreco 
         AutoSize        =   -1  'True
         Caption         =   "Preço"
         Height          =   195
         Left            =   2160
         TabIndex        =   35
         Top             =   360
         Width           =   420
      End
      Begin VB.Label labQtde 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   1200
         TabIndex        =   34
         Top             =   360
         Width           =   345
      End
      Begin VB.Label labDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   720
      End
      Begin VB.Label labCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraMesa 
      Caption         =   "Mesas"
      Height          =   3255
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   9615
      Begin VB.Frame fraFecharMesa 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2100
         Left            =   7200
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   3525
         Begin VB.TextBox txtCodFuncFecharMesa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   18
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtNomeFuncFecharMesa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   19
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtTotMesa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTaxaServ 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   21
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtTotGeral 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   22
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label labFuncionario 
            AutoSize        =   -1  'True
            Caption         =   "Funcionário:"
            Height          =   195
            Left            =   0
            TabIndex        =   41
            Top             =   45
            Width           =   870
         End
         Begin VB.Label labTotMesa 
            AutoSize        =   -1  'True
            Caption         =   "Total Mesa:"
            Height          =   195
            Left            =   0
            TabIndex        =   40
            Top             =   765
            Width           =   840
         End
         Begin VB.Label labTaxaServico 
            AutoSize        =   -1  'True
            Caption         =   "Taxa Serv."
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   1125
            Width           =   780
         End
         Begin VB.Label labTotGeral 
            AutoSize        =   -1  'True
            Caption         =   "Total Geral:"
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   1485
            Width           =   825
         End
      End
      Begin VB.Frame fraPagarMesa 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2100
         Left            =   3720
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox txtSubTicket 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   52
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtSubCartao 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   51
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtSubCheque 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   50
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtSubDinheiro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   840
            TabIndex        =   17
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtTicket 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   840
            TabIndex        =   16
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtCartao 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   840
            TabIndex        =   15
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtCheque 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   840
            TabIndex        =   14
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDinheiro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   840
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   840
            TabIndex        =   12
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label labSubTotal 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal:"
            Height          =   195
            Left            =   2040
            TabIndex        =   53
            Top             =   120
            Width           =   690
         End
         Begin VB.Label labDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
            Height          =   195
            Left            =   0
            TabIndex        =   48
            Top             =   1830
            Width           =   735
         End
         Begin VB.Label labTicket 
            AutoSize        =   -1  'True
            Caption         =   "Ticket:"
            Height          =   195
            Left            =   0
            TabIndex        =   47
            Top             =   1470
            Width           =   495
         End
         Begin VB.Label labCartao 
            AutoSize        =   -1  'True
            Caption         =   "Cartão:"
            Height          =   195
            Left            =   0
            TabIndex        =   46
            Top             =   1110
            Width           =   510
         End
         Begin VB.Label labCheque 
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
            Height          =   195
            Left            =   0
            TabIndex        =   45
            Top             =   750
            Width           =   600
         End
         Begin VB.Label labDinheiro 
            AutoSize        =   -1  'True
            Caption         =   "Dinheiro:"
            Height          =   195
            Left            =   0
            TabIndex        =   44
            Top             =   390
            Width           =   630
         End
         Begin VB.Label labTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   0
            TabIndex        =   43
            Top             =   30
            Width           =   405
         End
      End
      Begin VB.TextBox txtCodFunc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtMesa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtParcial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtNomeFunc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mebTurnoHora 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebTurnoData 
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lvwMesas 
         Height          =   2895
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imlListView"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.Image imgSalvar 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2160
         Picture         =   "frmMesa.frx":C71E
         ToolTipText     =   "Salvar"
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgImprimir 
         Enabled         =   0   'False
         Height          =   240
         Left            =   1800
         Picture         =   "frmMesa.frx":CAA8
         ToolTipText     =   "Imprimir"
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label labStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1080
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Image imgCancelar 
         Enabled         =   0   'False
         Height          =   240
         Left            =   1440
         Picture         =   "frmMesa.frx":CE32
         ToolTipText     =   "Cancelar Operação"
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label labMesa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mesa:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   420
         Width           =   435
      End
      Begin VB.Label labTurno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   780
         Width           =   465
      End
      Begin VB.Label labParcial 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcial:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2220
         Width           =   525
      End
      Begin VB.Label labGarcom 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Garçom:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1140
         Width           =   600
      End
   End
   Begin MSComctlLib.StatusBar stbMesas 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   6135
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12515
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1852
            MinWidth        =   1852
            TextSave        =   "07/02/2007"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "20:48"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMesas 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   630
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   9870
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrMesas"
      MinHeight1      =   38
      Width1          =   554
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMesas 
         Height          =   570
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlMesas"
         DisabledImageList=   "imlMesaD"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
               Object.ToolTipText     =   "Consultar"
               ImageKey        =   "consultar"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cancelar"
               Object.ToolTipText     =   "Cancelar Mesa"
               ImageKey        =   "cancelar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "transferir"
               Object.ToolTipText     =   "Transferir"
               ImageKey        =   "transferir"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "estornar"
               Object.ToolTipText     =   "Estornar"
               ImageKey        =   "estornar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "fechar"
               Object.ToolTipText     =   "Fechar"
               ImageKey        =   "fechar"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "pagar"
               Object.ToolTipText     =   "Pagar"
               ImageKey        =   "pagar"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sair"
               Object.ToolTipText     =   "Sair"
               ImageKey        =   "sair"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mniMenu 
         Caption         =   "&Menu"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mniPedido 
         Caption         =   "&Pedido"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mniSepUm 
         Caption         =   "-"
      End
      Begin VB.Menu mniSair 
         Caption         =   "Sai&r"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mniConsultar 
         Caption         =   "&Consultar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mniCancelar 
         Caption         =   "C&ancelar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mniImprimir 
         Caption         =   "Im&primir"
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mniSepDois 
         Caption         =   "-"
      End
      Begin VB.Menu mniTransfer 
         Caption         =   "&Transferir"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mniEstornar 
         Caption         =   "&Estornar"
         Shortcut        =   {F8}
         Visible         =   0   'False
      End
      Begin VB.Menu mniFechar 
         Caption         =   "&Fechar"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mniPagar 
         Caption         =   "&Pagar"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mniConteudo 
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Dim LocalizarOQue As String
Public Localizar As Boolean 'Define o retorno da tela localizar
Dim itemNumb As Integer ' Número de itens
Dim Sair As Boolean
Dim TempoDeMesaFechada As Integer
Dim TransfOrigem As String
Dim TransfDestino As String
Dim TotMesa As Single
Dim TaxaServ As Single

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If txtMesa.Enabled = False Then Call imgCancelar_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rowPos As Integer   ' Posição vertical do item
    Dim colPos As Integer   ' Posição horizontal do item
    Dim colNumb As Integer  ' Quantidade de itens por linha
    Dim I As Integer
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    itemNumb = Conn.GetValue("SELECT valor FROM config WHERE campo = 'QtdeMesas'")
    If Conn.NumErro <> 0 Then GoTo SubFail
    TempoDeMesaFechada = Conn.GetValue("SELECT valor FROM config WHERE campo = 'MesaFechada'")
    If Conn.NumErro <> 0 Then GoTo SubFail
    colPos = 0
    rowPos = 0
    colNumb = 0
    TransfOrigem = "00"
    For I = 0 To itemNumb
        colNumb = colNumb + 1
        lvwMesas.ListItems.Add I + 1, "_" & Format(I, "00"), Format(I, "00"), imlListView.ListImages.Item("azul").Index
        lvwMesas.ListItems.Item(I + 1).Left = colPos
        lvwMesas.ListItems.Item(I + 1).Top = rowPos
        Select Case MesaStatus(Format(I, "00"))
            Case -1
                GoTo SubFail
            Case 0
            Case 1
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("verde").Index
            Case 2
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("amarela").Index
            Case 3
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("vermelha").Index
        End Select
        colPos = colPos + 360
        If colNumb = 15 Then
            colPos = 0
            rowPos = rowPos + 600
            colNumb = 0
        End If
    Next
    With mfgMesas
        .Cols = 4
        .Rows = 1
        .Clear
        .ColAlignment(3) = 1
        .ColWidth(0) = 975
        .ColWidth(1) = 3210
        .ColWidth(2) = 855
        .ColWidth(3) = 1095
        .TextArray(0) = "Código"
        .TextArray(1) = "Descrição"
        .TextArray(2) = "Qtde"
        .TextArray(3) = "Preço"
    End With
    txtMesa.Text = "00"
    txtMesa.SelStart = 0
    txtMesa.SelLength = 2
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Disconnect
End Sub

Private Sub imgCancelar_Click()
    If Me.Tag = "Consultar" Then GoTo NaoPergunta
    If Me.Tag = "Cancelar" Then GoTo NaoPergunta
    If Me.Tag = "Transferir" Then GoTo NaoPergunta
    If Me.Tag = "Fechar" Then GoTo NaoPergunta
    If Me.Tag = "Pagar" Then GoTo NaoPergunta
    If txtMesa.DataChanged = True Then GoTo Question
    If mebTurnoData.DataChanged = True Then GoTo Question
    If mebTurnoHora.DataChanged = True Then GoTo Question
    If txtCodFunc.DataChanged = True Then GoTo Question
    If txtNomeFunc.DataChanged = True Then GoTo Question
    If txtParcial.DataChanged = True Then GoTo Question
    If txtCodPro.DataChanged = True Then GoTo Question
    If txtQtdePro.DataChanged = True Then GoTo Question
    If txtVlrPro.DataChanged = True Then GoTo Question
    If txtDescPro.DataChanged = True Then GoTo Question
    If mfgMesas.Tag = "Alterado" Then GoTo Question
    GoTo NaoPergunta
Question:
    If MsgBox("Descartar Alterações?", vbQuestion + vbYesNo _
        + vbDefaultButton2, "Confirmação.") = vbNo Then Exit Sub
NaoPergunta:
    mebTurnoData.Text = Empty
    mebTurnoHora.Text = Empty
    txtMesa.Text = Empty
    txtCodFunc.Text = Empty
    txtNomeFunc.Text = Empty
    txtParcial.Text = Empty
    txtCodPro.Text = Empty
    txtVlrPro.Text = Empty
    txtQtdePro.Text = Empty
    txtDescPro.Text = Empty
    mfgMesas.Enabled = False
    mfgMesas.Rows = 1
    Call HabCampos(True)
    lvwMesas.Enabled = False
    imgSalvar.Enabled = False
    imgSalvar.Visible = False
    imgImprimir.Enabled = False
    imgImprimir.Visible = False
    mniImprimir.Visible = False
    mniConsultar.Visible = True
    lvwMesas.ListItems.Item("_" & TransfOrigem).Ghosted = False
    tbrMesas.Buttons.Item("consultar").Visible = True
    txtMesa.DataChanged = False
    mebTurnoData.DataChanged = False
    mebTurnoHora.DataChanged = False
    txtCodFunc.DataChanged = False
    txtNomeFunc.DataChanged = False
    txtParcial.DataChanged = False
    txtCodPro.DataChanged = False
    txtQtdePro.DataChanged = False
    txtVlrPro.DataChanged = False
    txtDescPro.DataChanged = False
    mfgMesas.Tag = Empty
    Me.Caption = "Controle de Mesas"
    fraPagarMesa.Visible = False
    mebTurnoData.Left = 840
    mebTurnoHora.Left = 2040
    txtMesa.Left = 840
    imgCancelar.Left = 1440
    labTurno.Caption = "Turno:"
    labMesa.Caption = "Mesa:"
    txtMesa.SetFocus
End Sub

Private Sub imgSalvar_Click()
    If txtMesa.DataChanged = True Then GoTo Salva
    If mebTurnoData.DataChanged = True Then GoTo Salva
    If mebTurnoHora.DataChanged = True Then GoTo Salva
    If txtCodFunc.DataChanged = True Then GoTo Salva
    If txtNomeFunc.DataChanged = True Then GoTo Salva
    If txtParcial.DataChanged = True Then GoTo Salva
    If txtCodPro.DataChanged = True Then GoTo Salva
    If txtQtdePro.DataChanged = True Then GoTo Salva
    If txtVlrPro.DataChanged = True Then GoTo Salva
    If txtDescPro.DataChanged = True Then GoTo Salva
    If mfgMesas.Tag = "Alterado" Then GoTo Salva
    Exit Sub
Salva:
    Dim Mesa As Long
    Dim I As Integer
    If mfgMesas.Rows = 1 Then
        MsgBox "Por favor informe pelo menos" & Chr(13) & _
            "um item para gravar a mesa", vbExclamation, "Mesa Vazia"
        Exit Sub
    End If
    If labStatus.Caption = "**Abrindo**" Then
        Call Conn.Query("INSERT INTO mesa ( " _
            & "id, mesa, dataAbe, horaAbe, fid ) " _
            & "VALUES ( null, '" _
            & txtMesa.Text & "', '" _
            & mebTurnoData.Text & "', '" _
            & mebTurnoHora.Text & "', " _
            & txtCodFunc.Text & " )")
        Mesa = Conn.GetValue("SELECT MAX(id) FROM mesa")
        For I = 1 To mfgMesas.Rows - 1
            Call Conn.Query("INSERT INTO mesaItem ( " _
                & "id, mid, pid, qtde, preco ) " _
                & "VALUES ( null, " _
                & Mesa & ", " _
                & mfgMesas.TextMatrix(I, 0) & ", " _
                & mfgMesas.TextMatrix(I, 2) & ", " _
                & Replace(CDbl(mfgMesas.TextMatrix(I, 3)), ",", ".") _
                & " )")
        Next
    Else
        Call Conn.Query("UPDATE mesa SET " _
            & "fid = " & txtCodFunc.Text _
            & " WHERE mesa = '" & txtMesa.Text & "'")
        Mesa = Conn.GetValue("SELECT id FROM mesa WHERE mesa = '" _
            & txtMesa.Text & "'")
        Call Conn.Query("DELETE FROM mesaItem WHERE mid = " & Mesa)
        For I = 1 To mfgMesas.Rows - 1
            Call Conn.Query("INSERT INTO mesaItem ( " _
                & "id, mid, pid, qtde, preco ) " _
                & "VALUES ( null, " _
                & Mesa & ", " _
                & mfgMesas.TextMatrix(I, 0) & ", " _
                & mfgMesas.TextMatrix(I, 2) & ", " _
                & Replace(CDbl(mfgMesas.TextMatrix(I, 3)), ",", ".") _
                & " )")
        Next
    End If
    lvwMesas.ListItems.Item("_" & txtMesa.Text).Icon = imlListView.ListImages.Item("verde").Index
    mebTurnoData.Text = Empty
    mebTurnoHora.Text = Empty
    txtMesa.Text = Empty
    txtCodFunc.Text = Empty
    txtNomeFunc.Text = Empty
    txtParcial.Text = Empty
    txtCodPro.Text = Empty
    txtVlrPro.Text = Empty
    txtQtdePro.Text = Empty
    txtDescPro.Text = Empty
    mfgMesas.Enabled = False
    mfgMesas.Rows = 1
    stbMesas.Panels.Item(1).Text = ""
    Call HabCampos(True)
    imgSalvar.Enabled = False
    imgSalvar.Visible = False
    txtMesa.DataChanged = False
    mebTurnoData.DataChanged = False
    mebTurnoHora.DataChanged = False
    txtCodFunc.DataChanged = False
    txtNomeFunc.DataChanged = False
    txtParcial.DataChanged = False
    txtCodPro.DataChanged = False
    txtQtdePro.DataChanged = False
    txtVlrPro.DataChanged = False
    txtDescPro.DataChanged = False
    mfgMesas.Tag = Empty
    txtMesa.SetFocus
End Sub

Private Sub lvwMesas_Click()
    Call lvwMesas_KeyUp(vbKeyLeft, 0)
End Sub

Private Sub lvwMesas_DblClick()
    Call lvwMesas_KeyPress(vbKeyReturn)
End Sub

Private Sub lvwMesas_GotFocus()
    lvwMesas.Tag = "Selected"
End Sub

Private Sub lvwMesas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Select Case Me.Tag
            Case "Consultar"
                Select Case MesaStatus(lvwMesas.SelectedItem)
                    Case 0 'Mesa Livre
                        imgCancelar_Click
                        txtMesa.Text = lvwMesas.SelectedItem.Text
                        Call txtMesa_KeyPress(vbKeyReturn)
                    Case 1 'Mesa Aberta
                        mniFechar_Click
                        imgCancelar_Click
                    Case 2, 3 'Mesa Fechada há muito tempo; 'Mesa Fechada
                        'As duas execuções são necessárias.
                        'Preparação para pagamento
                        mniPagar_Click
                        'Pagamento de fato
                        mniPagar_Click
                End Select
            Case "Cancelar"
                If Not MesaStatus(lvwMesas.SelectedItem) = 0 Then
                    mniCancelar_Click
                End If
            Case "Transferir"
                mniTransfer_Click
            Case "Fechar"
                If MesaStatus(lvwMesas.SelectedItem) = 1 Then
                    mniFechar_Click
                End If
            Case "Pagar"
                If MesaStatus(lvwMesas.SelectedItem) = 2 Or _
                    MesaStatus(lvwMesas.SelectedItem) = 3 Then
                    mniPagar_Click
                End If
        End Select
    End If
End Sub

Private Sub lvwMesas_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Or _
        KeyCode = vbKeyRight Or _
        KeyCode = vbKeyDown Or _
        KeyCode = vbKeyUp Then
        Dim Livre As Boolean
        Select Case MesaStatus(lvwMesas.SelectedItem)
            Case 0 'Mesa Livre
                Livre = True
            Case 1 'Mesa Aberta
                labStatus.Caption = "**Mesa Aberta**"
                Livre = False
            Case 2 'Mesa Fechada há muito tempo
                labStatus.Caption = "**Mesa Fechada**"
                Livre = False
            Case 3 'Mesa Fechada
                labStatus.Caption = "**Mesa Fechada**"
                Livre = False
        End Select
        If Livre Then
            labStatus.Caption = "**Mesa Livre**"
            txtMesa.Text = Empty
            mebTurnoData.Text = Empty
            mebTurnoHora.Text = Empty
            txtCodFunc.Text = Empty
            mfgMesas.Rows = 1
        Else
            Call CarregaValores("lvwMesas_KeyUp")
        End If
        Call CalculaParcial
        If Me.Tag = "Transferir" And _
            Mid(stbMesas.Panels.Item(1).Text, 1, 6) = "Origem" Then
            stbMesas.Panels.Item(1).Text = _
                Mid(stbMesas.Panels.Item(1).Text, 1, 10) _
                & " Destino: " & lvwMesas.SelectedItem
        End If
    End If
End Sub

Private Sub lvwMesas_LostFocus()
    lvwMesas.Tag = ""
End Sub

Private Sub lvwMesas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.Tag = "Cancelar" Then Exit Sub
    If Me.Tag = "Transferir" Then Exit Sub
    If Me.Tag = "Fechar" Then Exit Sub
    If Me.Tag = "Pagar" Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuEditar
    End If
End Sub

Private Sub mebTurnoData_GotFocus()
    mebTurnoData.Tag = mebTurnoData.Text
    mebTurnoData.SelStart = 0
    mebTurnoData.SelLength = 10
    mebTurnoData.BackColor = &H80000018 'Amarelo
End Sub

Private Sub mebTurnoData_LostFocus()
    mebTurnoData.BackColor = &H80000005 'Branco
    If mebTurnoData.Tag = mebTurnoData.Text Then Exit Sub
    mebTurnoData.Text = modData.FmtData(mebTurnoData.Text)
End Sub

Private Sub mebTurnoHora_GotFocus()
    mebTurnoHora.BackColor = &H80000018 'Amarelo
    mebTurnoHora.Tag = mebTurnoHora.Text
    mebTurnoHora.SelStart = 0
    mebTurnoHora.SelLength = 5
End Sub

Private Sub mebTurnoHora_LostFocus()
    mebTurnoHora.BackColor = &H80000005 'Branco
    If mebTurnoHora.Tag = mebTurnoHora.Text Then Exit Sub
    mebTurnoHora.Text = modData.FmtHora(mebTurnoHora.Text)
End Sub

Private Sub mfgMesas_GotFocus()
    stbMesas.Panels.Item(1).Text = "Pressione Del para Subtrair 1, Shift+Del para Excluir o item, Esc para Cancelar"
End Sub

Private Sub mfgMesas_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And Shift = 1 Then
        GoTo Exclui
    ElseIf KeyCode = vbKeyDelete And Shift = 0 Then
        If mfgMesas.TextMatrix(mfgMesas.Row, 2) = 1 Then
            GoTo Exclui
        Else
            mfgMesas.TextMatrix(mfgMesas.Row, 2) = _
                CInt(mfgMesas.TextMatrix(mfgMesas.Row, 2)) - 1
        End If
        GoTo Finaliza
    End If
    Exit Sub
Exclui:
    If mfgMesas.Rows = 2 Then
        mfgMesas.Rows = 1
        mfgMesas.Enabled = False
        txtCodPro.SetFocus
    Else
        mfgMesas.RemoveItem mfgMesas.Row
    End If
Finaliza:
    mfgMesas.Tag = "Alterado"
    Call CalculaParcial
End Sub

Private Sub mfgMesas_LostFocus()
    stbMesas.Panels.Item(1).Text = "Pressione Esc para Cancelar"
End Sub

Private Sub mniCancelar_Click()
    If lvwMesas.Tag = Empty Then
        Call HabCampos(False)
        lvwMesas.Enabled = True
        labStatus.Caption = "**Cancelando**"
        labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
        stbMesas.Panels.Item(1).Text = "Pressione Enter para cancelar, Esc para Cancelar"
        Call lvwMesas_KeyUp(vbKeyLeft, 0)
        Me.Tag = "Cancelar"
        Me.Caption = Me.Caption & " - Cancelando"
        lvwMesas.SetFocus
    Else
        frmPrin.Auth = False
        frmPrin.TituloSenha = "Senha de Usuário"
        frmSenha.Show vbModal, Me
        frmPrin.TituloSenha = "Senha de Acesso"
        If Not frmPrin.Auth Then
            Exit Sub
        End If
        frmPrin.Auth = False
        frmPrin.TituloSenha = "Senha de Supervisor"
        frmSenha.Show vbModal, Me
        frmPrin.TituloSenha = "Senha de Acesso"
        If frmPrin.Auth And frmPrin.NivelUsuarioLogado <= 2 Then
            Dim MesaId As Long
            MesaId = Conn.GetValue("SELECT id FROM mesa WHERE mesa = '" _
                & lvwMesas.SelectedItem & "'")
            Call Conn.Query("DELETE FROM mesaItem WHERE mid = " & MesaId)
            Call Conn.Query("DELETE FROM mesa WHERE mesa = '" _
                & lvwMesas.SelectedItem & "'")
            lvwMesas.SelectedItem.Icon = _
                imlListView.ListImages.Item("azul").Index
            MsgBox "Cancelado com Êxito", vbInformation, "Concluído!"
        Else
            MsgBox "Necessário Autorização de Supervisor", _
                vbExclamation, "Autorização"
            lvwMesas.Tag = "Selected"
        End If
    End If
End Sub

Private Sub mniConsultar_Click()
    imgCancelar.Enabled = True
    imgCancelar.Visible = True
    imgImprimir.Enabled = True
    imgImprimir.Visible = True
    txtMesa.Enabled = False
    mniConsultar.Enabled = False
    mniConsultar.Visible = False
    mniConteudo.Enabled = False
    mniMenu.Enabled = False
    mniPedido.Enabled = False
    mniSair.Enabled = False
    mniImprimir.Visible = True
    tbrMesas.Buttons.Item("sair").Enabled = False
    tbrMesas.Buttons.Item("consultar").Enabled = False
    tbrMesas.Buttons.Item("consultar").Visible = False
    labStatus.Visible = True
    labStatus.Caption = "**Consultando**"
    labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
    lvwMesas.Enabled = True
    Call lvwMesas_KeyUp(vbKeyLeft, 0)
    stbMesas.Panels.Item(1).Text = ""
    Me.Tag = "Consultar"
    Me.Caption = Me.Caption & " - Consultando"
    lvwMesas.SetFocus
End Sub

Private Sub mniEstornar_Click()
    'TODO
End Sub

Private Sub mniFechar_Click()
    If lvwMesas.Tag = Empty Then
        Call HabCampos(False)
        lvwMesas.Enabled = True
        labStatus.Caption = "**Fechando**"
        labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
        stbMesas.Panels.Item(1).Text = "Pressione Enter para Fechar, Esc para Cancelar"
        Call lvwMesas_KeyUp(vbKeyLeft, 0)
        Me.Tag = "Fechar"
        Me.Caption = Me.Caption & " - Fechando"
        lvwMesas.SetFocus
    Else
        fraFecharMesa.Visible = True
        fraFecharMesa.Left = 140
        fraFecharMesa.Top = 1080
        mebTurnoData.Left = mebTurnoData.Left + 260
        mebTurnoHora.Left = mebTurnoHora.Left + 260
        txtMesa.Left = txtMesa.Left + 260
        imgCancelar.Left = imgCancelar.Left + 260
        txtCodFuncFecharMesa.Text = txtCodFunc.Text
        txtNomeFuncFecharMesa.Text = txtNomeFunc.Text
        txtTotMesa.Text = Format(TotMesa, "R$ ###,##0.00")
        txtTaxaServ.Text = Format(TaxaServ, "R$ ###,##0.00")
        txtTotGeral.Text = txtParcial.Text
        frmPrin.Auth = False
        frmPrin.TituloSenha = "Senha de Usuário"
        frmSenha.Show vbModal, Me
        frmPrin.TituloSenha = "Senha de Acesso"
        If frmPrin.Auth Then
            Call Conn.Query("UPDATE mesa SET " _
                & "dataFec = '" & Replace(Date, "/", "") _
                & "', horaFec = '" & Replace(Time, ":", "") _
                & "' WHERE mesa = '" & lvwMesas.SelectedItem & "'")
            lvwMesas.SelectedItem.Icon = _
                imlListView.ListImages.Item("vermelha").Index
            MsgBox "Fechado com Êxito", vbInformation, "Concluído!"
            lvwMesas.Tag = "Selected"
        End If
        fraFecharMesa.Visible = False
        mebTurnoData.Left = 840
        mebTurnoHora.Left = 2040
        txtMesa.Left = 840
        imgCancelar.Left = 1440
    End If
End Sub

Private Sub mniMenu_Click()
    frmPrin.WindowState = vbNormal
    Unload Me
End Sub

Private Sub mniPagar_Click()
    If lvwMesas.Tag = Empty Then
        Call HabCampos(False)
        lvwMesas.Enabled = True
        labStatus.Caption = "**Pagando**"
        labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
        stbMesas.Panels.Item(1).Text = "Pressione Enter para Pagar, Esc para Cancelar"
        Call lvwMesas_KeyUp(vbKeyLeft, 0)
        Me.Tag = "Pagar"
        Me.Caption = Me.Caption & " - Pagando"
        lvwMesas.SetFocus
    Else
        imgImprimir.Visible = False
        imgSalvar.Visible = True
        fraPagarMesa.Visible = True
        fraPagarMesa.Left = 120
        fraPagarMesa.Top = 1080
        txtMesa.Left = 840 + 120
        imgCancelar.Left = 1440 + 120
        imgSalvar.Left = imgCancelar.Left + 360
        mebTurnoData.Left = 840 + 120
        mebTurnoHora.Left = 2040 + 120
        txtTotal.Text = txtParcial.Text
        lvwMesas.Enabled = False
        labTurno.Caption = "Emissão:"
        stbMesas.Panels.Item(1).Text = "Restante: " & txtTotal.Text
        txtSubDinheiro.Text = Empty
        txtSubCheque.Text = Empty
        txtSubCartao.Text = Empty
        txtSubTicket.Text = Empty
        txtDesconto.Text = Empty
        txtDinheiro.SetFocus
'        frmPrin.Auth = False
'        frmPrin.TituloSenha = "Senha de Usuário"
'        frmSenha.Show vbModal, Me
'        frmPrin.TituloSenha = "Senha de Acesso"
'        If frmPrin.Auth Then
'            Call Conn.Query("UPDATE mesa SET " _
'                & "dataFec = '" & Replace(Date, "/", "") _
'                & "', horaFec = '" & Replace(Time, ":", "") _
'                & "' WHERE mesa = '" & lvwMesas.SelectedItem & "'")
'            lvwMesas.SelectedItem.Icon = _
'                imlListView.ListImages.Item("vermelha").Index
'            MsgBox "Fechado com Êxito", vbInformation, "Concluído!"
'            lvwMesas.Tag = "Selected"
'        End If
'        fraPagarMesa.Visible = False
'        mebTurnoData.Left = 840
'        mebTurnoHora.Left = 2040
'        txtMesa.Left = 840
'        imgCancelar.Left = 1440
    End If
End Sub

Private Sub mniPedido_Click()
    'TODO
End Sub

Private Sub mniImprimir_Click()
    MsgBox "print"
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniTransfer_Click()
    If lvwMesas.Tag = Empty Then
        Call HabCampos(False)
        txtCodFunc.Enabled = False
        txtCodPro.Enabled = False
        lvwMesas.Enabled = True
        labStatus.Caption = "**Transferindo**"
        labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
        stbMesas.Panels.Item(1).Text = "Selecione a Mesa de Origem e Pressione Enter, Esc para Cancelar"
        Call lvwMesas_KeyUp(vbKeyLeft, 0)
        Me.Tag = "Transferir"
        Me.Caption = Me.Caption & " - Transferindo"
        lvwMesas.SetFocus
    Else
        If Mid(stbMesas.Panels.Item(1).Text, 1, 9) = "Selecione" Then
            TransfOrigem = lvwMesas.SelectedItem.Text
            lvwMesas.SelectedItem.Ghosted = True
            stbMesas.Panels.Item(1).Text = "Origem: " _
                & TransfOrigem & " Destino:"
        Else
            frmPrin.Auth = False
            frmPrin.TituloSenha = "Senha de Usuário"
            frmSenha.Show vbModal, Me
            frmPrin.TituloSenha = "Senha de Acesso"
            If Not frmPrin.Auth Then
                Exit Sub
            End If
            frmPrin.Auth = False
            frmPrin.TituloSenha = "Senha de Supervisor"
            frmSenha.Show vbModal, Me
            frmPrin.TituloSenha = "Senha de Acesso"
            If frmPrin.Auth And frmPrin.NivelUsuarioLogado <= 2 Then
                Dim MesaId As Long
                TransfDestino = lvwMesas.SelectedItem.Text
                If TransfDestino = TransfOrigem Then
                    lvwMesas.ListItems.Item("_" & TransfOrigem).Ghosted = False
                    Exit Sub
                End If
                Call Conn.Query("UPDATE mesa SET mesa = '" _
                    & TransfDestino & "' WHERE mesa = '" _
                    & TransfOrigem & "'")
                lvwMesas.ListItems.Item("_" & TransfDestino).Icon = _
                    lvwMesas.ListItems.Item("_" & TransfOrigem).Icon
                lvwMesas.ListItems.Item("_" & TransfOrigem).Icon = _
                    imlListView.ListImages.Item("azul").Index
                lvwMesas.ListItems.Item("_" & TransfOrigem).Ghosted = False
                stbMesas.Panels.Item(1).Text = _
                    "Selecione a Mesa de Origem e Pressione Enter" _
                    & ", Esc para Cancelar"
                MsgBox "Transferido com Êxito", vbInformation, "Concluído!"
            Else
                MsgBox "Necessário Autorização de Supervisor", _
                    vbExclamation, "Autorização"
                lvwMesas.Tag = "Selected"
            End If
        End If
    End If
End Sub

Private Sub tbrMesas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "consultar"
            mniConsultar_Click
        Case "cancelar"
            mniCancelar_Click
        Case "transferir"
            mniTransfer_Click
        Case "estornar"
            mniEstornar_Click
        Case "fechar"
            mniFechar_Click
        Case "pagar"
            mniPagar_Click
        Case "sair"
            mniSair_Click
    End Select
End Sub

Private Sub tmrCores_Timer()
    Dim I As Integer
    For I = 0 To itemNumb - 1
        Select Case MesaStatus(Format(I, "00"))
            Case -1
                GoTo SubFail
            Case 0
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("azul").Index
            Case 1
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("verde").Index
            Case 2
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("amarela").Index
            Case 3
                lvwMesas.ListItems.Item(I + 1).Icon = _
                    imlListView.ListImages.Item("vermelha").Index
        End Select
    Next
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
End Sub

Private Sub tmrMesa_Timer()
    tmrMesa.Enabled = False
    stbMesas.Panels.Item(1).Text = Empty
End Sub

Private Sub txtCartao_GotFocus()
    txtCartao.SelStart = 0
    txtCartao.SelLength = 10
    txtCartao.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtCartao_KeyPress(KeyAscii As Integer)
    KeyAscii = Pagar(KeyAscii, txtCartao, txtSubCartao)
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
    KeyAscii = Pagar(KeyAscii, txtCheque, txtSubCheque)
End Sub

Private Sub txtCheque_LostFocus()
    txtCheque.Text = modMoeda.FmtMoeda(txtCheque.Text)
    txtCheque.BackColor = &H80000005 'Branco
End Sub

Private Sub txtCodFunc_Change()
    If txtCodFunc.Text = Empty Then GoTo Limpa
    Call Conn.Query("SELECT nome FROM funcionario" _
        & " WHERE id = " & txtCodFunc.Text)
    If Conn.Rs.RecordCount = 0 Then
        txtNomeFunc.Text = "Registro Inexistente."
    Else
        txtNomeFunc.Text = Conn.Rs.Fields.Item("nome").Value
    End If
    Exit Sub
Limpa:
    txtNomeFunc.Text = Empty
End Sub

Private Sub txtCodFunc_GotFocus()
    txtCodFunc.SelStart = 0
    txtCodFunc.SelLength = 11
    txtCodFunc.BackColor = &H80000018 'Amarelo
    stbMesas.Panels.Item(1).Text = "Pressione F5 para Localizar, Esc para Cancelar"
    LocalizarOQue = "Funcionario"
End Sub

Private Sub txtCodFunc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodFunc_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite FROM funcionario")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodFunc.Text = Conn.Rs.Fields.Item("id").Value
    End If
End Sub

Private Sub txtCodFunc_LostFocus()
    txtCodFunc.BackColor = &H80000005 'Branco
    stbMesas.Panels.Item(1).Text = "Pressione Esc para Cancelar"
    LocalizarOQue = "Mesa"
End Sub

Public Sub Load()
    Select Case LocalizarOQue
        Case "Funcionario"
            Call LoadFuncionario
        Case "Produto"
            Call LoadProduto
    End Select
End Sub

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case LocalizarOQue
        Case "Funcionario"
            Call CampoFuncionario(txt)
        Case "Produto"
            Call CampoProduto(txt)
    End Select
End Sub

Public Sub LoadFuncionario()
    With frmLocalizar.cbbProcura
        .AddItem ("Código")
        .AddItem ("Nome")
        .AddItem ("CPF")
        .Text = "Nome"
    End With
    With frmLocalizar.mfgLocalizar
        .Cols = 3
        .Rows = 1
        .Clear
        .ColWidth(0) = 1000
        .ColWidth(1) = 4000
        .ColWidth(2) = 2100
        .TextArray(0) = "Código"
        .TextArray(1) = "Nome"
        .TextArray(2) = "CPF"
        Conn.Rs.MoveFirst
        Do While Not Conn.Rs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Conn.Rs.Fields.Item("id").Value
            .TextMatrix(.Rows - 1, 1) = Conn.Rs.Fields.Item("nome").Value
            .TextMatrix(.Rows - 1, 2) = Conn.Rs.Fields.Item("cpf").Value
            Conn.Rs.MoveNext
        Loop
    End With
    frmLocalizar.Caption = "Localização de Funcionários"
    frmLocalizar.tabela = "funcionario"
End Sub

Public Sub CampoFuncionario(txt As String)
    Select Case txt
        Case "Código"
            frmLocalizar.Campo = "id"
        Case "Nome"
            frmLocalizar.Campo = "nome"
        Case "CPF"
            frmLocalizar.Campo = "cpf"
    End Select
End Sub

Public Sub LoadProduto()
    With frmLocalizar.cbbProcura
        .AddItem ("Código")
        .AddItem ("Descrição")
        .Text = "Descrição"
    End With
    With frmLocalizar.mfgLocalizar
        .Cols = 4
        .Rows = 1
        .Clear
        .ColWidth(0) = 800
        .ColWidth(1) = 3200
        .ColWidth(2) = 2100
        .ColWidth(3) = 1447
        .TextArray(0) = "Código"
        .TextArray(1) = "Descrição"
        .TextArray(2) = "Preço"
        .TextArray(3) = "Tipo"
        Conn.Rs.MoveFirst
        Do While Not Conn.Rs.EOF
            Dim x As Integer
            Dim y As String
            x = Conn.Rs.Fields.Item("tipo").Value
            y = IIf(x = 1, "Comida", "Bebida")
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Conn.Rs.Fields.Item("id").Value
            .TextMatrix(.Rows - 1, 1) = Conn.Rs.Fields.Item("descricao").Value
            .TextMatrix(.Rows - 1, 2) = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value)
            .TextMatrix(.Rows - 1, 3) = y
            Conn.Rs.MoveNext
        Loop
        frmLocalizar.Caption = "Localização de Produtos"
        frmLocalizar.tabela = "produto"
    End With
End Sub

Public Sub CampoProduto(txt As String)
    Select Case txt
        Case "Código"
            frmLocalizar.Campo = "id"
        Case "Descrição"
            frmLocalizar.Campo = "descricao"
    End Select
End Sub

Private Sub txtCodPro_Change()
    If txtCodPro.Text = Empty Then GoTo Limpa
    Call Conn.Query("SELECT produto.descricao, produto.preco" _
        & ", promocao.preco as promo FROM produto" _
        & " LEFT JOIN promocao ON produto.id = promocao.pid" _
        & " AND promocao.dia = " & Weekday(Date, vbSunday) _
        & " WHERE produto.id = " & txtCodPro.Text)
    If Conn.Rs.RecordCount = 0 Then
        txtVlrPro.Text = Empty
        txtQtdePro.Text = Empty
        txtDescPro.Text = "Registro Inexistente."
        txtQtdePro.Enabled = False
        txtVlrPro.Enabled = False
    Else
        If IsNull(Conn.Rs.Fields.Item("promo").Value) Then
            Conn.Rs.Fields.Item("promo").Value = _
                Conn.Rs.Fields.Item("preco").Value
        End If
        txtDescPro.Text = Conn.Rs.Fields.Item("descricao").Value
        txtVlrPro.Text = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("promo").Value)
        txtQtdePro.Text = 1
        txtQtdePro.Enabled = True
        txtVlrPro.Enabled = True
    End If
    Exit Sub
Limpa:
    txtDescPro.Text = Empty
    txtVlrPro.Text = Empty
    txtVlrPro.Enabled = False
    txtQtdePro.Text = Empty
    txtQtdePro.Enabled = False
End Sub

Private Sub txtCodPro_GotFocus()
    txtCodPro.SelStart = 0
    txtCodPro.SelLength = 11
    txtCodPro.BackColor = &H80000018 'Amarelo
    stbMesas.Panels.Item(1).Text = "Pressione F5 para Localizar, Enter para inserir, Esc para Cancelar"
    LocalizarOQue = "Produto"
End Sub

Private Sub txtCodPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim I As Integer
        If txtCodPro.Text = Empty Then
            If txtMesa.Text = "00" Then
                lvwMesas.Tag = "Pagar"
                Me.Tag = "Pagar"
                lvwMesas.ListItems.Item(1).Selected = True
                Call mniPagar_Click
                stbMesas.Panels.Item(1).Text = "Restante: " _
                    & Format(CSng(txtTotal.Text) - _
                    CSng(IIf(txtSubDinheiro.Text = Empty, 0, txtSubDinheiro.Text)) - _
                    CSng(IIf(txtSubCheque.Text = Empty, 0, txtSubCheque.Text)) - _
                    CSng(IIf(txtSubCartao.Text = Empty, 0, txtSubCartao.Text)) - _
                    CSng(IIf(txtSubTicket.Text = Empty, 0, txtSubTicket.Text)), _
                    "R$ ###,##0.00")
            Else
                Call imgSalvar_Click
            End If
            Exit Sub
        End If
        If txtDescPro.Text = "Registro Inexistente." Then
            MsgBox "Registro Inexistente.", vbInformation
            Exit Sub
        End If
        With mfgMesas
            For I = 0 To .Rows - 1
                If .TextMatrix(I, 0) = txtCodPro.Text _
                    And .TextMatrix(I, 3) = txtVlrPro.Text Then
                    .TextMatrix(I, 2) = CInt(.TextMatrix(I, 2)) + CInt(txtQtdePro.Text)
                    txtCodPro.Text = Empty
                    txtDescPro.Text = Empty
                    txtVlrPro.Text = Empty
                    txtQtdePro.Text = Empty
                    Call CalculaParcial
                    Exit Sub
                End If
            Next
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = txtCodPro.Text
            .TextMatrix(.Rows - 1, 1) = txtDescPro.Text
            .TextMatrix(.Rows - 1, 2) = txtQtdePro.Text
            .TextMatrix(.Rows - 1, 3) = modMoeda.FmtMoeda(txtVlrPro.Text)
            mfgMesas.Enabled = True
            txtCodPro.Text = Empty
            txtDescPro.Text = Empty
            txtVlrPro.Text = Empty
            txtQtdePro.Text = Empty
        End With
        txtCodPro.SelStart = 0
        txtCodPro.SelLength = 11
        Call CalculaParcial
        Exit Sub
    End If
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodPro.Text = Conn.Rs.Fields.Item("id").Value
    End If
End Sub

Private Sub txtCodPro_LostFocus()
    txtCodPro.BackColor = &H80000005 'Branco
    stbMesas.Panels.Item(1).Text = "Pressione Esc para Cancelar"
    LocalizarOQue = "Mesa"
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
    KeyAscii = Pagar(KeyAscii, txtDinheiro, txtSubDinheiro)
End Sub

Private Sub txtDinheiro_LostFocus()
    txtDinheiro.Text = modMoeda.FmtMoeda(txtDinheiro.Text)
    txtDinheiro.BackColor = &H80000005 'Branco
End Sub

Private Sub txtMesa_GotFocus()
    txtMesa.BackColor = &H80000018 'Amarelo
    stbMesas.Panels.Item(1).Text = "Digite o número da mesa e pressione ENTER para abrí-la"
End Sub

Private Sub txtMesa_KeyPress(KeyAscii As Integer)
    If txtMesa.Text = Empty Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        If itemNumb < txtMesa.Text Then
            tmrMesa.Enabled = True
            stbMesas.Panels.Item(1).Text = "Mesa Inexistente"
            Exit Sub
        End If
        Select Case MesaStatus(Format(txtMesa.Text, "00"))
            Case 0 'Mesa Livre
                mebTurnoData.Text = Date
                mebTurnoHora.Text = Time
                imgSalvar.Enabled = True
                imgSalvar.Visible = True
                imgSalvar.Left = 1800
                Call HabCampos(False)
                txtCodFunc.Enabled = True
                txtCodPro.Enabled = True
                txtMesa.Text = Format(txtMesa.Text, "00")
                labStatus.Caption = "**Abrindo**"
                labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
                If txtMesa.Text = "00" Then
                    labMesa.Caption = "Balcão:"
                    txtCodFunc.Enabled = False
                    txtCodFunc.Text = Conn.GetValue("SELECT valor FROM config WHERE campo = 'GarcomBalcao'")
                    txtCodPro.SetFocus
                Else
                    txtCodFunc.SetFocus
                End If
                Exit Sub
            Case 1 'Mesa Aberta
                labStatus.Caption = "**Mesa Aberta**"
                Call CarregaValores("txtMesa_KeyPress")
                Call CalculaParcial
            Case 2, 3 'Mesa Fechada há muito tempo, Mesa Fechada
                labStatus.Visible = True
                labStatus.Caption = "**Mesa Fechada**"
                labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
                Call CarregaValores("txtMesa_KeyPress")
                Call CalculaParcial
                Call HabCampos(False)
                txtCodFunc.Enabled = False
                txtCodPro.Enabled = False
                Me.Tag = "Consultar"
                Exit Sub
        End Select
        Call HabCampos(False)
        txtCodFunc.Enabled = True
        txtCodPro.Enabled = True
        txtMesa.Text = Format(txtMesa.Text, "00")
        imgSalvar.Enabled = True
        imgSalvar.Visible = True
        imgSalvar.Left = 1800
        labStatus.Caption = "**Alterando**"
        labStatus.Left = (lvwMesas.Left - fraMesa.Left - labStatus.Width) / 2
        txtCodPro.SetFocus
        txtMesa.DataChanged = False
        mebTurnoData.DataChanged = False
        mebTurnoHora.DataChanged = False
        txtCodFunc.DataChanged = False
        txtNomeFunc.DataChanged = False
        txtParcial.DataChanged = False
        txtCodPro.DataChanged = False
        txtQtdePro.DataChanged = False
        txtVlrPro.DataChanged = False
        txtDescPro.DataChanged = False
        mfgMesas.Tag = Empty
    End If
End Sub

Private Sub txtMesa_LostFocus()
    txtMesa.BackColor = &H80000005 'Branco
End Sub

Private Sub txtParcial_GotFocus()
    txtParcial.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtParcial_LostFocus()
    txtParcial.BackColor = &H80000005 'Branco
End Sub

Private Sub HabCampos(situ As Boolean)
    txtCodFunc.Enabled = situ
    txtCodPro.Enabled = situ
    imgCancelar.Enabled = Not situ
    imgCancelar.Visible = Not situ
    labStatus.Visible = Not situ
    txtMesa.Enabled = situ
    mniCancelar.Enabled = situ
    mniConsultar.Enabled = situ
    mniConteudo.Enabled = situ
    mniEstornar.Enabled = situ
    mniFechar.Enabled = situ
    mniMenu.Enabled = situ
    mniPagar.Enabled = situ
    mniPedido.Enabled = situ
    mniSair.Enabled = situ
    mniTransfer.Enabled = situ
    tbrMesas.Buttons.Item("transferir").Enabled = situ
    tbrMesas.Buttons.Item("estornar").Enabled = situ
    tbrMesas.Buttons.Item("pagar").Enabled = situ
    tbrMesas.Buttons.Item("sair").Enabled = situ
    tbrMesas.Buttons.Item("consultar").Enabled = situ
    tbrMesas.Buttons.Item("fechar").Enabled = situ
    tbrMesas.Buttons.Item("cancelar").Enabled = situ
End Sub

Private Sub txtQtdePro_GotFocus()
    txtQtdePro.SelStart = 0
    txtQtdePro.SelLength = 10
    txtQtdePro.BackColor = &H80000018 'Amarelo
    stbMesas.Panels.Item(1).Text = "Pressione Enter para inserir, Esc para Cancelar"
End Sub

Private Sub txtQtdePro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call txtCodPro_KeyPress(KeyAscii)
        txtQtdePro.SelStart = 0
        txtQtdePro.SelLength = 10
        txtCodPro.SetFocus
    End If
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtQtdePro_LostFocus()
    txtQtdePro.BackColor = &H80000005 'Branco
    stbMesas.Panels.Item(1).Text = "Pressione Esc para Cancelar"
End Sub

Private Sub txtTicket_GotFocus()
    txtTicket.SelStart = 0
    txtTicket.SelLength = 10
    txtTicket.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtTicket_KeyPress(KeyAscii As Integer)
    KeyAscii = Pagar(KeyAscii, txtTicket, txtSubTicket)
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

Private Sub txtVlrPro_GotFocus()
    txtVlrPro.SelStart = 0
    txtVlrPro.SelLength = 10
    txtVlrPro.BackColor = &H80000018 'Amarelo
    stbMesas.Panels.Item(1).Text = "Pressione Enter para inserir, Esc para Cancelar"
End Sub

Private Sub txtVlrPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtVlrPro.Text = modMoeda.FmtMoeda(txtVlrPro.Text)
        Call txtCodPro_KeyPress(KeyAscii)
        txtVlrPro.SelStart = 0
        txtVlrPro.SelLength = 10
        txtCodPro.SetFocus
    End If
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtVlrPro_LostFocus()
    txtVlrPro.Text = modMoeda.FmtMoeda(txtVlrPro.Text)
    txtVlrPro.BackColor = &H80000005 'Branco
    stbMesas.Panels.Item(1).Text = "Pressione Esc para Cancelar"
End Sub

Private Sub CalculaParcial()
    Call Conn.Query("SELECT valor FROM config WHERE campo = 'Taxa';")
    Dim I As Integer
    Dim Valor As Integer
    TotMesa = Empty
    TaxaServ = Empty
    For I = 1 To mfgMesas.Rows - 1
        TotMesa = TotMesa + CSng(mfgMesas.TextMatrix(I, 3)) * mfgMesas.TextMatrix(I, 2)
    Next
    Valor = Val(Conn.Rs.Fields.Item("valor").Value)
    txtParcial.Text = Format(TotMesa * ((Valor / 100) + 1), "R$ ###,##0.00")
    TotMesa = Format(TotMesa, "R$ ###,##0.00")
    TaxaServ = (TotMesa * ((Valor / 100) + 1)) - TotMesa
    If txtParcial.Text = "R$ 0,00" Then txtParcial.Text = Empty
End Sub

Private Sub CarregaValores(QualSub As String)
    Dim MesaId As Long
    MesaId = Conn.Rs.Fields.Item("id").Value
    txtMesa.Text = Conn.Rs.Fields.Item("mesa").Value
    mebTurnoData.Text = Conn.Rs.Fields.Item("dataAbe").Value
    mebTurnoHora.Text = Conn.Rs.Fields.Item("horaAbe").Value
    txtCodFunc.Text = Conn.Rs.Fields.Item("fid").Value
    Call Conn.Query("SELECT mesaItem.pid, mesaItem.qtde" _
        & ", mesaItem.preco, produto.descricao " _
        & "FROM mesaItem " _
        & "LEFT JOIN produto ON produto.id = mesaItem.pid " _
        & "WHERE mid = " & MesaId)
    mfgMesas.Rows = 1
    Do While Not Conn.Rs.EOF
        With mfgMesas
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Conn.Rs.Fields.Item("pid").Value
            .TextMatrix(.Rows - 1, 1) = Conn.Rs.Fields.Item("descricao").Value
            .TextMatrix(.Rows - 1, 2) = Conn.Rs.Fields.Item("qtde").Value
            .TextMatrix(.Rows - 1, 3) = Format(Conn.Rs.Fields.Item("preco").Value, "R$ ###,##0.00")
        End With
        Conn.Rs.MoveNext
    Loop
End Sub

Private Property Get MesaStatus(Mesa As String) As Integer
    Dim newHour As Integer
    Dim newMinute As Integer
    Dim newSecond As Integer
    Dim MesaTime As String
    Dim DataFec As Date
    Dim HoraFec As Date
    Dim StrDataFec As String
    Dim StrHoraFec As String
    Call Conn.Query("SELECT id, mesa, dataAbe, horaAbe, dataFec, horaFec, fid " _
        & "FROM mesa WHERE mesa = '" & Mesa & "'")
    If Conn.NumErro <> 0 Then GoTo MesaStatusFail
    If Conn.Rs.RecordCount = 0 Then
        MesaStatus = 0 'Mesa Livre
    Else
        If Conn.Rs.Fields.Item("dataFec").Value = Empty Then
            MesaStatus = 1 'Mesa Aberta
        Else
            StrDataFec = IIf(IsNull(Conn.Rs.Fields.Item("dataFec").Value), "", Conn.Rs.Fields.Item("dataFec").Value)
            StrHoraFec = IIf(IsNull(Conn.Rs.Fields.Item("horaFec").Value), "", Conn.Rs.Fields.Item("horaFec").Value)
            DataFec = CDate(Mid(StrDataFec, 1, 2) & "/" _
                & Mid(StrDataFec, 3, 2) & "/" & Mid(StrDataFec, 5, 4))
            HoraFec = CDate(Mid(StrHoraFec, 1, 2) & ":" _
                & Mid(StrHoraFec, 3, 2))
            newHour = Hour(HoraFec)
            newMinute = Minute(HoraFec) + TempoDeMesaFechada
            newSecond = Second(HoraFec)
            MesaTime = TimeSerial(newHour, newMinute, newSecond)
            If DataFec < Date Or MesaTime < Time Then
                MesaStatus = 2 'Mesa Fechada há muito tempo
            ElseIf MesaTime >= Time Then
                MesaStatus = 3 'Mesa Fechada
            End If
        End If
    End If
    Exit Property
MesaStatusFail:
    MesaStatus = -1
End Property

Private Property Get Pagar(KeyAscii As Integer, ByRef SrcField As Object, ByRef DestField As Object)
    Pagar = KeyAscii
    If KeyAscii = vbKeyReturn Then
        If SrcField.Text = Empty Then
            If CSng(Mid(stbMesas.Panels.Item(1).Text, 10)) _
                - CSng(IIf(txtDesconto.Text = Empty, _
                0, txtDesconto.Text)) = 0 Then
                Dim MesaId As Integer
                Dim Taxa As Integer
                Dim I As Integer
                Dim VendaId As Integer
                MesaId = Conn.GetValue("SELECT id FROM mesa WHERE mesa = '" _
                    & lvwMesas.SelectedItem.Text & "'")
                Taxa = Conn.GetValue("SELECT valor FROM config WHERE " _
                    & "campo = 'Taxa'")
                Call Conn.Query("DELETE FROM mesaItem WHERE mid = " & MesaId)
                Call Conn.Query("DELETE FROM mesa WHERE id = " & MesaId)
                If txtSubDinheiro.Text = Empty Then txtSubDinheiro.Text = "0"
                If txtSubCheque.Text = Empty Then txtSubCheque.Text = "0"
                If txtSubCartao.Text = Empty Then txtSubCartao.Text = "0"
                If txtSubTicket.Text = Empty Then txtSubTicket.Text = "0"
                If txtDesconto.Text = Empty Then txtDesconto.Text = "0"
                Call Conn.Query("INSERT INTO venda ( id, fid, dataVen" _
                    & ", horaVen, taxa, mesa, dinheiro, cheque, cartao" _
                    & ", ticket, desconto ) VALUES ( null" _
                    & ", " & txtCodFunc.Text _
                    & ", '" & Replace(Date, "/", "") & "'" _
                    & ", '" & Replace(Time, ":", "") & "'" _
                    & ", " & Taxa _
                    & ", '" & lvwMesas.SelectedItem.Text & "'" _
                    & ", " & Replace(CSng(txtSubDinheiro.Text), ",", ".") _
                    & ", " & Replace(CSng(txtSubCheque.Text), ",", ".") _
                    & ", " & Replace(CSng(txtSubCartao.Text), ",", ".") _
                    & ", " & Replace(CSng(txtSubTicket.Text), ",", ".") _
                    & ", " & Replace(CSng(txtDesconto.Text), ",", ".") & ")")
                VendaId = Conn.GetValue("SELECT MAX(id) FROM venda")
                For I = 1 To mfgMesas.Rows - 1
                    Call Conn.Query("INSERT INTO vendaItem ( " _
                        & "id, pid, vid, qtde, preco ) " _
                        & "VALUES ( null, " _
                        & mfgMesas.TextMatrix(I, 0) & ", " _
                        & VendaId & ", " _
                        & mfgMesas.TextMatrix(I, 2) & ", " _
                        & Replace(CDbl(mfgMesas.TextMatrix(I, 3)), ",", ".") _
                        & " )")
                Next
                lvwMesas.SelectedItem.Icon = _
                    imlListView.ListImages.Item("azul").Index
                imgCancelar_Click
            End If
        ElseIf SrcField.Text = "-" Then
            DestField.Text = Empty
            SrcField.Text = Empty
            stbMesas.Panels.Item(1).Text = "Restante: " _
                & Format(CSng(txtTotal.Text) - _
                CSng(IIf(txtSubDinheiro.Text = Empty, 0, txtSubDinheiro.Text)) - _
                CSng(IIf(txtSubCheque.Text = Empty, 0, txtSubCheque.Text)) - _
                CSng(IIf(txtSubCartao.Text = Empty, 0, txtSubCartao.Text)) - _
                CSng(IIf(txtSubTicket.Text = Empty, 0, txtSubTicket.Text)), _
                "R$ ###,##0.00")
        Else
            DestField.Text = _
                Format(CSng(SrcField.Text) + _
                CSng(IIf(DestField.Text = Empty, 0, _
                DestField.Text)), "R$ ###,##0.00")
            SrcField.Text = Empty
            stbMesas.Panels.Item(1).Text = "Restante: " _
                & Format(CSng(txtTotal.Text) - _
                CSng(IIf(txtSubDinheiro.Text = Empty, 0, txtSubDinheiro.Text)) - _
                CSng(IIf(txtSubCheque.Text = Empty, 0, txtSubCheque.Text)) - _
                CSng(IIf(txtSubCartao.Text = Empty, 0, txtSubCartao.Text)) - _
                CSng(IIf(txtSubTicket.Text = Empty, 0, txtSubTicket.Text)), _
                "R$ ###,##0.00")
        End If
        Exit Property
    ElseIf KeyAscii = 45 Then ' -
        If Not SrcField.Text = Empty Then
            Pagar = Empty
        End If
        Exit Property
    End If
    Pagar = modGetKeyAscii.Moeda(KeyAscii)
End Property
