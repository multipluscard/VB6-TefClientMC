VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FormPrincipal 
   Caption         =   "Exemplo TefClientMC - VB6"
   ClientHeight    =   9195
   ClientLeft      =   1860
   ClientTop       =   1500
   ClientWidth     =   7215
   LinkTopic       =   "FormPrincipal"
   ScaleHeight     =   9195
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAtributos 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txbTelefone 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txbData 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txbCnpjParceiro 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txbControle 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txbCupom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txbTexto 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox txbCnpjCliente 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txbParcelas 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txbValor 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbTelefone 
         Caption         =   "TELEFONE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbData 
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbCnpjParceiro 
         Caption         =   "CNPJ PARCEIRO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbControle 
         Caption         =   "CONTROLE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbCupom 
         Caption         =   "CUPOM"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbCnpj 
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lbTexto 
         Caption         =   "TEXTO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbParcela 
         Caption         =   "PARCELA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lbValor 
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTabTipos 
      Height          =   5655
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CARTÃO"
      TabPicture(0)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LineCartao(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTabCartao(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "QRMULTIPLUS"
      TabPicture(1)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTabPix(1)"
      Tab(1).Control(1)=   "LineQRmultiplus(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "LINKPAGO"
      TabPicture(2)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTabLinkPago(0)"
      Tab(2).Control(1)=   "LineLinkPago(0)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "CLIENT"
      TabPicture(3)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTabClient(1)"
      Tab(3).Control(1)=   "LineLinkPago(1)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "OUTROS"
      TabPicture(4)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTabParceleMais(0)"
      Tab(4).Control(1)=   "LineLinkPago(2)"
      Tab(4).ControlCount=   2
      Begin TabDlg.SSTab SSTabCartao 
         Height          =   4815
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "CREDITO"
         TabPicture(0)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":008C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "brnCancelarCredito(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnCreditoAVista(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "btnCreditoParcelamentoADM(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnCreditoParcelamentoLoja(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "DEBITO"
         TabPicture(1)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":00A8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "btnVendeDebitoAVista(0)"
         Tab(1).Control(1)=   "btnVendeDebito(1)"
         Tab(1).Control(2)=   "brnCancelarDebito(1)"
         Tab(1).ControlCount=   3
         Begin VB.CommandButton btnVendeDebitoAVista 
            Caption         =   "VENDE DEBITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   -74640
            TabIndex        =   26
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebito 
            Caption         =   "VENDE DEBITO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   -74640
            TabIndex        =   25
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnCreditoParcelamentoLoja 
            Caption         =   "CREDITO PARCELAMENTO LOJA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   360
            TabIndex        =   23
            Top             =   2400
            Width           =   2895
         End
         Begin VB.CommandButton btnCreditoParcelamentoADM 
            Caption         =   "CREDITO PARCELAMENTO ADM"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnCreditoAVista 
            Caption         =   "CREDITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   2895
         End
         Begin VB.PictureBox brnCancelarDebito 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   -74640
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   27
            Top             =   3840
            Width           =   2175
         End
         Begin VB.PictureBox brnCancelarCredito 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   360
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   24
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6570
            Y1              =   -120
            Y2              =   -117
         End
      End
      Begin TabDlg.SSTab SSTabPix 
         Height          =   4815
         Index           =   1
         Left            =   -74880
         TabIndex        =   28
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PIX"
         TabPicture(0)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":00C4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "brnCancelarPIX(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnEnviaPix(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.CommandButton btnEnviaPix 
            Caption         =   "ENVIA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   360
            TabIndex        =   31
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebito 
            Caption         =   "VENDE DEBITO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   -74640
            TabIndex        =   30
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebitoAVista 
            Caption         =   "VENDE DEBITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   -74640
            TabIndex        =   29
            Top             =   1560
            Width           =   2895
         End
         Begin VB.PictureBox brnCancelarPIX 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   360
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   33
            Top             =   3840
            Width           =   2175
         End
         Begin VB.PictureBox brnCancelarDebito 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   -74640
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   32
            Top             =   3840
            Width           =   2175
         End
      End
      Begin TabDlg.SSTab SSTabLinkPago 
         Height          =   4815
         Index           =   0
         Left            =   -74880
         TabIndex        =   34
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LINKPAGO"
         TabPicture(0)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":00E0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "btnEnviaLinkPago(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnListaLinkPago(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "btnManutLinkPago(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "TxbQtdeItens"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "TxbItens"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         Begin VB.TextBox TxbItens 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   59
            Top             =   3960
            Width           =   4815
         End
         Begin VB.TextBox TxbQtdeItens 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   58
            Top             =   3120
            Width           =   1695
         End
         Begin VB.CommandButton btnManutLinkPago 
            Caption         =   "MANUTENÇÃO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   3360
            TabIndex        =   55
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnListaLinkPago 
            Caption         =   "LISTAR LINKS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   360
            TabIndex        =   39
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebitoAVista 
            Caption         =   "VENDE DEBITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   -74640
            TabIndex        =   37
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebito 
            Caption         =   "VENDE DEBITO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   -74640
            TabIndex        =   36
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnEnviaLinkPago 
            Caption         =   "ENVIA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   360
            TabIndex        =   35
            Top             =   720
            Width           =   2895
         End
         Begin VB.PictureBox brnCancelarDebito 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   -74640
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   38
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "ITENS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "QTDE ITENS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   2760
            Width           =   1575
         End
      End
      Begin TabDlg.SSTab SSTabClient 
         Height          =   4815
         Index           =   1
         Left            =   -74880
         TabIndex        =   40
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "CLIENT"
         TabPicture(0)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":00FC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "btnATV(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnAdm(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "btnReimpressao(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnReimpressaoDireta(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "btnSolicitarCpf(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "btnRelatorio(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         Begin VB.CommandButton btnRelatorio 
            Caption         =   "RELATORIO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   3480
            TabIndex        =   49
            Top             =   2400
            Width           =   2895
         End
         Begin VB.CommandButton btnSolicitarCpf 
            Caption         =   "SOLICITAR CPF"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Top             =   2400
            Width           =   2895
         End
         Begin VB.CommandButton btnReimpressaoDireta 
            Caption         =   "REIMPRESSAO DIRETA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   3480
            TabIndex        =   47
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnReimpressao 
            Caption         =   "REIMPRESSAO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   3480
            TabIndex        =   46
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnAdm 
            Caption         =   "ADM"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebito 
            Caption         =   "VENDE DEBITO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   -74640
            TabIndex        =   43
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebitoAVista 
            Caption         =   "VENDE DEBITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   -74640
            TabIndex        =   42
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnATV 
            Caption         =   "ATV"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   1560
            Width           =   2895
         End
         Begin VB.PictureBox brnCancelarDebito 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   -74640
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   45
            Top             =   3840
            Width           =   2175
         End
      End
      Begin TabDlg.SSTab SSTabParceleMais 
         Height          =   4815
         Index           =   0
         Left            =   -74880
         TabIndex        =   50
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PARCELE MAIS"
         TabPicture(0)   =   "ExemploDLLTefClientMC_ComTelaVB6.frx":0118
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "btnParceleMais(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.CommandButton btnVendeDebitoAVista 
            Caption         =   "VENDE DEBITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   -74640
            TabIndex        =   53
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton btnVendeDebito 
            Caption         =   "VENDE DEBITO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   -74640
            TabIndex        =   52
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton btnParceleMais 
            Caption         =   "PARCELE MAIS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   360
            TabIndex        =   51
            Top             =   720
            Width           =   2895
         End
         Begin VB.PictureBox brnCancelarDebito 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   -74640
            ScaleHeight     =   435
            ScaleWidth      =   2115
            TabIndex        =   54
            Top             =   3840
            Width           =   2175
         End
      End
      Begin VB.Line LineLinkPago 
         BorderColor     =   &H00008080&
         BorderWidth     =   3
         Index           =   2
         X1              =   -74760
         X2              =   -68400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line LineLinkPago 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   -74760
         X2              =   -68400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line LineLinkPago 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         Index           =   0
         X1              =   -74760
         X2              =   -68400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line LineQRmultiplus 
         BorderColor     =   &H00FFC0FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   -74760
         X2              =   -68400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line LineCartao 
         BorderColor     =   &H00808000&
         BorderWidth     =   3
         Index           =   0
         X1              =   240
         X2              =   6600
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TRANSACAO DE CREDITO, NAO SABENDO SE É A VISTA OU PARCELADO
Private Declare Function VendeCredito Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iNumeroCupom As Long, ByVal iLeitor As Integer) As String

'TRANSACAO DE CREDITO A VISTA
Private Declare Function VendeCreditoVista Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iNumeroCupom As Integer, ByVal iLeitor As Integer) As String

'TRANSACAO DE CREDITO PARCELADO LOJA
Private Declare Function VendeCreditoParcLoja Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iParcelas As Integer, ByVal dValor As Double, ByVal iNumeroCupom As Integer, ByVal iLeitor As Integer) As String

'TRANSACAO DE CREDITO PARCELADO ADM
Private Declare Function VendeCreditoParcAdm Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iParcelas As Integer, ByVal dValor As Double, ByVal iNumeroCupom As Integer, ByVal iLeitor As Integer) As String

'TRANSACAO DEBITO
Private Declare Function VendeDebito Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iNumeroCupom As Long, ByVal iLeitor As Integer) As String

'TRANSACAO DEBITO A VISTA
Private Declare Function VendeDebitoAVista Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iNumeroCupom As Long, ByVal iLeitor As Integer) As String

'CONFIRMACAO DE TRANSACOES
Private Declare Function Confirmar Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iCupom As Integer) As String

'CANCELAR TRANSACOES
Private Declare Function Cancelar Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iCupom As Integer, ByVal sControle As String, ByVal iLeitor As Integer) As String

'FUNÇÕES ADM
Private Declare Function Adm Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'ATV
Private Declare Function Atv Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'RELATÓRIO
Private Declare Function Relatorio Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'DESFAZIMENTO
Private Declare Function Desfazimento Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iCupom As Integer) As String

'SOLICITAR CPF
Private Declare Function SolicitarCPF Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iCupom As Integer) As String

'PIX
Private Declare Function VendeCarteiraDigitalPix Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'LINK PAGO
Private Declare Function LinkPagamento Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iParcelas As Integer, ByVal dValor As Double, ByVal iCupom As Integer, ByVal iQtdeItens As Integer, ByVal sItens As String, ByVal sTelefone As String, ByVal sTexto As String, ByVal iLeitor As Integer) As String

'LISTAR LINK PAGO
Private Declare Function ListarLinkPagamentoPago Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'REIMPRESSAO
Private Declare Function Reimpressao Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'REIMPRESSAO DIRETO
Private Declare Function ReimpressaoDireto Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal sControle As String, ByVal sData As String, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'PARCELE MAIS
Private Declare Function ParceleMais Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal dValor As Double, ByVal iCupom As Integer, ByVal iLeitor As Integer) As String

'MANUTENCAO DE LINKS
Private Declare Function ManutencaoLinkPagamento Lib "TefClientmc.dll" (ByVal sCNPJCliente As String, ByVal sCNPJParceiro As String, ByVal iLeitor As Integer) As String

Public sCNPJClient As String
Public sCNPJParceiro As String
Public sData As String

Private Sub Transacionar(sTipo As String)
    
    Dim dValor As Double
    Dim iCupom As Long
    Dim iParcelas As Long
    Dim sRetorno As String
    Dim sTelefone As String
    Dim sTexto As String
    Dim sControle As String
    Dim vLinhas As Variant
    Dim vTipos As Variant
    
    Dim QtdeItens As Integer
    Dim sItens As String
    
    Dim sRetornoTransacao As String
    Dim sMensagemTransacao As String
    Dim sComprovanteTransacao As String
    
    sCNPJClient = txbCnpjCliente.Text
    sCNPJParceiro = txbCnpjParceiro.Text
    sTelefone = txbTelefone.Text
    sTexto = txbTexto.Text
    'sData = txbData.Text
    sControle = txbControle.Text
    
    If sCNPJClient = "" Or sCNPJParceiro = "" Then
      Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
         If Response = vbYes Then
             MyString = "Yes"
         Exit Sub
         End If
    End If
    
    If txbValor.Text = "" Then
         txbValor.Text = 0
         Exit Sub
    End If
    
    If txbCupom.Text = "" Then
         txbCupom.Text = 0
         Exit Sub
    End If
    
    If txbParcelas.Text = "" Then
         txbParcelas.Text = 0
         Exit Sub
    End If
    
    If TxbQtdeItens.Text = "" Then
         TxbQtdeItens.Text = 0
         Exit Sub
    End If
    
         
    dValor = CDbl(txbValor.Text)
    iCupom = CInt(txbCupom.Text)
    iParcelas = CInt(txbParcelas.Text)
    sData = CDate(txbData.Text)
    QtdeItens = CInt(TxbQtdeItens.Text)
    
    sItens = TxbItens.Text
    
    Select Case sTipo
        Case "CREDITO_A_VISTA"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
               Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                  If Response = vbYes Then
                      MyString = "Yes"
                      Exit Sub
                  End If
                  Exit Sub
            End If
            sRetorno = VendeCreditoVista(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "CREDITO_PARCELAMENTO_ADM"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or iParcelas = 0 Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = VendeCreditoParcAdm(sCNPJClient, sCNPJParceiro, iParcelas, dValor, iCupom, 0)
        Case "CREDITO_PARCELAMENTO_LOJA"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or iParcelas = 0 Or dValor = 0 Or iCupom = 0 Then
               Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                  If Response = vbYes Then
                      MyString = "Yes"
                      Exit Sub
                  End If
                  Exit Sub
            End If
            sRetorno = VendeCreditoParcLoja(sCNPJClient, sCNPJParceiro, iParcelas, dValor, iCupom, 0)
        Case "DEBITO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = VendeDebito(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "DEBITO_A_VISTA"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = VendeDebitoAVista(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "CANCELAR"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or sControle = "" Or dValor = 0 Or iCupom = 0 Then
               Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                  If Response = vbYes Then
                      MyString = "Yes"
                      Exit Sub
                  End If
                  Exit Sub
            End If
            sRetorno = Cancelar(sCNPJClient, sCNPJParceiro, dValor, iCupom, sControle, 0)
        Case "VENDE_CARTEIRA_DIGITAL_PIX"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = VendeCarteiraDigitalPix(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "LINK_PAGO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or sTelefone = "" Or sTexto = "" Or iParcelas = 0 Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = LinkPagamento(sCNPJClient, sCNPJParceiro, iParcelas, dValor, iCupom, QtdeItens, sItens, sTelefone, sTexto, 0)
        Case "LISTAR_LINK_PAGO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = ListarLinkPagamentoPago(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "ADM"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = Adm(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "ATV"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = Atv(sCNPJClient, sCNPJParceiro, iCupom, 0)
        Case "REIMPRESSAO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = Reimpressao(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
        Case "REIMPRESSAO_DIRETO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or sControle = "" Or sData = "" Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = ReimpressaoDireto(sCNPJClient, sCNPJParceiro, sControle, sData, iCupom, 0)
        Case "RELATORIO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = Relatorio(sCNPJClient, sCNPJParceiro, iCupom, 0)
        Case "SOLICITAR_CPF"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = SolicitarCPF(sCNPJClient, sCNPJParceiro, iCupom)
        Case "PARCELE_MAIS"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = ParceleMais(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
            
        Case "MANUTENCAO_LINK_PAGO"
            If sCNPJClient = "" Or sCNPJParceiro = "" Or dValor = 0 Or iCupom = 0 Then
              Response = MsgBox("Por favor verifique os campos solicitados.", vbOKOnly + vbCritical + vbDefaultButton2, "ERRO", "", 0)
                 If Response = vbYes Then
                     MyString = "Yes"
                     Exit Sub
                 End If
                 Exit Sub
            End If
            sRetorno = ParceleMais(sCNPJClient, sCNPJParceiro, dValor, iCupom, 0)
            
            
            
    End Select
    
    
    If sRetorno = "" Then
        Exit Sub
    End If
    
    vLinhas = Split(sRetorno, vbNewLine)
    For I = LBound(vLinhas) To UBound(vLinhas)
        vTipos = Split(vLinhas(I), ";")
        If vTipos(0) = "S" Then
          sRetornoTransacao = vTipos(1)
          sMensagemTransacao = vTipos(2)
          
          If sRetornoTransacao = "0" Then
              MsgBox (sMensagemTransacao)
              Exit Sub
          End If
        Else
            If vTipos(0) = "I" Then
                sComprovanteTransacao = sComprovanteTransacao + vTipos(1) + vbNewLine
            End If
        End If
    Next
    
    MsgBox (sMensagemTransacao + vbNewLine + sComprovanteTransacao)
    sRetorno = Confirmar(sCNPJClient, sCNPJParceiro, 12345)
End Sub


Private Sub brnCancelarCredito_Click(Index As Integer)
     Transacionar ("CANCELAR")
End Sub

Private Sub brnCancelarDebito_Click(Index As Integer)
      Transacionar ("CANCELAR")
End Sub

Private Sub brnCancelarPIX_Click(Index As Integer)
     Transacionar ("CANCELAR")
End Sub

Private Sub btnAdm_Click(Index As Integer)
   Transacionar ("ADM")
End Sub

Private Sub btnATV_Click(Index As Integer)
      Transacionar ("ATV")
End Sub

Private Sub btnEnviaLinkPago_Click(Index As Integer)
   Transacionar ("LINK_PAGO")
End Sub

Private Sub btnEnviaPix_Click(Index As Integer)
   Transacionar ("VENDE_CARTEIRA_DIGITAL_PIX")
End Sub

Private Sub btnListaLinkPago_Click(Index As Integer)
   Transacionar ("LISTAR_LINK_PAGO")
End Sub

Private Sub btnManutLinkPago_Click(Index As Integer)
    Transacionar ("MANUTENCAO_LINK_PAGO")
End Sub

Private Sub btnParceleMais_Click(Index As Integer)
    Transacionar ("PARCELE_MAIS")
End Sub

Private Sub btnReimpressao_Click(Index As Integer)
   Transacionar ("REIMPRESSAO")
End Sub

Private Sub btnReimpressaoDireta_Click(Index As Integer)
    Transacionar ("REIMPRESSAO_DIRETO")
End Sub

Private Sub btnRelatorio_Click(Index As Integer)
   Transacionar ("RELATORIO")
End Sub

Private Sub btnSolicitarCpf_Click(Index As Integer)
   Transacionar ("SOLICITAR_CPF")
End Sub

Private Sub btnVendeDebito_Click(Index As Integer)
   Transacionar ("DEBITO")
End Sub

Private Sub btnVendeDebitoAVista_Click(Index As Integer)
   Transacionar ("DEBITO_A_VISTA")
End Sub

Private Sub buttonLog_Click()
      FormRefErro.Show
End Sub

Private Sub Form_Load()
    ChDir "C:\DLL" 'Mudar o diretório atual para reconhecer a dll
    txbData.Text = Date
End Sub

Private Sub btnCreditoParcelamentoADM_Click(Index As Integer)
   Transacionar ("CREDITO_PARCELAMENTO_ADM")
End Sub

Private Sub btnCreditoParcelamentoLoja_Click(Index As Integer)
   Transacionar ("CREDITO_PARCELAMENTO_LOJA")
End Sub

Private Sub btnCreditoAVista_Click(Index As Integer)
  Transacionar ("CREDITO_A_VISTA")
End Sub

Private Sub txbCnpjCliente_GotFocus()
   On Error Resume Next
   txbCnpjCliente.SelStart = 0
   txbCnpjCliente.SelLength = Len(txbCnpjCliente.Text)
End Sub

Private Sub txbCnpjCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbCnpjParceiro_GotFocus()
   On Error Resume Next
   txbCnpjParceiro.SelStart = 0
   txbCnpjParceiro.SelLength = Len(txbCnpjParceiro.Text)
End Sub

Private Sub txbCnpjParceiro_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbControle_GotFocus()
   On Error Resume Next
   txbControle.SelStart = 0
   txbControle.SelLength = Len(txbControle.Text)
End Sub

Private Sub txbControle_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbData_GotFocus()
   On Error Resume Next
   txbData.SelStart = 0
   txbData.SelLength = Len(txbData.Text)
End Sub

Private Sub txbData_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbParcelas_GotFocus()
   On Error Resume Next
   txbParcelas.SelStart = 0
   txbParcelas.SelLength = Len(txbParcelas.Text)
End Sub

Private Sub txbParcelas_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbTelefone_GotFocus()
   On Error Resume Next
   txbTelefone.SelStart = 0
   txbTelefone.SelLength = Len(txbTelefone.Text)
End Sub

Private Sub txbTelefone_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbTexto_GotFocus()
   On Error Resume Next
   txbTexto.SelStart = 0
   txbTexto.SelLength = Len(txbTexto.Text)
End Sub

Private Sub txbTexto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbValor_GotFocus()
   On Error Resume Next
   txbValor.SelStart = 0
   txbValor.SelLength = Len(txbValor.Text)
End Sub

Private Sub txbValor_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbCupom_GotFocus()
   On Error Resume Next
   txbCupom.SelStart = 0
   txbCupom.SelLength = Len(txbCupom.Text)
End Sub

Private Sub txbCupom_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



