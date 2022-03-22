VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B4409115-5405-11D3-943D-0080AD4162AE}#1.0#0"; "ECOMBO.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "BARCOD32.OCX"
Begin VB.Form DetalleArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Artículos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DetalleArticulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Impuesto de ventas"
      Height          =   285
      Left            =   4785
      TabIndex        =   66
      Top             =   2415
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Información del Artículo "
      ForeColor       =   &H00FF0000&
      Height          =   1800
      Left            =   75
      TabIndex        =   57
      Top             =   525
      Width           =   6990
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   0
         Left            =   1185
         MaxLength       =   15
         TabIndex        =   61
         Top             =   405
         Width           =   1905
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   1
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   60
         Top             =   735
         Width           =   5715
      End
      Begin VB.TextBox TLista 
         Height          =   375
         Index           =   0
         Left            =   3375
         TabIndex        =   58
         Top             =   1110
         Width           =   555
      End
      Begin EnhancedCombo.ECombo ECombo1 
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   59
         Top             =   1110
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   225
         Left            =   450
         TabIndex        =   64
         Top             =   435
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   225
         Left            =   75
         TabIndex        =   63
         Top             =   795
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grupo :"
         Height          =   225
         Index           =   0
         Left            =   495
         TabIndex        =   62
         Top             =   1170
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descuento de Temporada "
      ForeColor       =   &H00FF0000&
      Height          =   1050
      Left            =   2625
      TabIndex        =   50
      Top             =   5040
      Width           =   3945
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   3
         Left            =   1335
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   53
         Top             =   420
         Width           =   570
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   2340
         TabIndex        =   51
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         Format          =   67239937
         CurrentDate     =   36832
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2340
         TabIndex        =   52
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Format          =   67239937
         CurrentDate     =   36832
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuento (%) :"
         Height          =   225
         Left            =   105
         TabIndex        =   56
         Top             =   465
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Del:"
         Height          =   255
         Left            =   2010
         TabIndex        =   55
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label14 
         Caption         =   "Al:"
         Height          =   225
         Left            =   2040
         TabIndex        =   54
         Top             =   660
         Width           =   315
      End
   End
   Begin VB.TextBox TLista 
      Height          =   375
      Index           =   5
      Left            =   9825
      TabIndex        =   49
      Top             =   4665
      Width           =   675
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   8
      Left            =   9645
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   3840
      Width           =   510
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluír en Otra      Compañía"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4560
      TabIndex        =   44
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   42
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Outlet"
      Height          =   345
      Left            =   6225
      TabIndex        =   41
      Top             =   4965
      Visible         =   0   'False
      Width           =   855
   End
   Begin BarcodLib.Barcod Barcode1 
      Height          =   315
      Left            =   6405
      TabIndex        =   40
      Top             =   4605
      Width           =   1065
      _Version        =   65543
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   75
      Caption         =   "TH"
      BackColor       =   16777215
      BarWidth        =   0
      Direction       =   0
      Style           =   18
      UPCNotches      =   3
      Alignment       =   0
      Extension       =   ""
   End
   Begin VB.TextBox TEtiquetas 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   7665
      TabIndex        =   39
      ToolTipText     =   "Cantidad de Etiquetas a imprimir"
      Top             =   4650
      Width           =   525
   End
   Begin VB.CommandButton Command5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7200
      Picture         =   "DetalleArticulo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Imprimir las Etiquetas"
      Top             =   5415
      Width           =   1185
   End
   Begin VB.TextBox TLista 
      Height          =   375
      Index           =   2
      Left            =   9825
      TabIndex        =   3
      Top             =   5070
      Width           =   555
   End
   Begin VB.TextBox TLista 
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   1
      Top             =   4290
      Width           =   615
   End
   Begin VB.TextBox PrecioIVI 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3600
      TabIndex        =   37
      Text            =   "0"
      Top             =   3765
      Width           =   1815
   End
   Begin VB.TextBox PreciO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1350
      TabIndex        =   34
      Text            =   "0"
      Top             =   3750
      Width           =   1725
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Color"
      Height          =   345
      Left            =   8430
      TabIndex        =   32
      Top             =   795
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox Factor2 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   5655
      TabIndex        =   8
      Top             =   3930
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Factor1 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   5640
      TabIndex        =   7
      Top             =   3540
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   870
      Left            =   15
      TabIndex        =   18
      Top             =   5145
      Width           =   2565
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comisión"
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Costos"
      ForeColor       =   &H000000FF&
      Height          =   1110
      Left            =   135
      TabIndex        =   19
      Top             =   2340
      Width           =   3615
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   270
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   255
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   2
         Left            =   1230
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   2265
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   10
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   2265
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Local :"
         Height          =   225
         Left            =   660
         TabIndex        =   20
         Top             =   690
         Width           =   510
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Anterior :"
         Height          =   225
         Left            =   405
         TabIndex        =   24
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   465
      Left            =   5565
      Picture         =   "DetalleArticulo.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Eliminar la imágen actual"
      Top             =   4875
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   465
      Left            =   4965
      Picture         =   "DetalleArticulo.frx":0986
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Buscar una imágen"
      Top             =   4875
      Width           =   495
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   11
      Left            =   9555
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   15
      Width           =   270
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   7
      Left            =   7905
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   6
      Left            =   7905
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3930
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   6015
      Picture         =   "DetalleArticulo.frx":0A88
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2820
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   585
      Left            =   4770
      Picture         =   "DetalleArticulo.frx":0BD2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2835
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   4
      Left            =   6840
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3615
      Width           =   465
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   150
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Index           =   1
      Left            =   9315
      TabIndex        =   0
      Top             =   4290
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Index           =   2
      Left            =   9255
      TabIndex        =   2
      Top             =   5070
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Index           =   3
      Left            =   9285
      TabIndex        =   33
      Top             =   765
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Index           =   4
      Left            =   8940
      TabIndex        =   35
      Top             =   375
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Index           =   5
      Left            =   9180
      TabIndex        =   47
      Top             =   4680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Mantenimiento de Artículos"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   930
      TabIndex        =   65
      Top             =   60
      Width           =   5685
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Proveedor :"
      Height          =   225
      Left            =   8235
      TabIndex        =   48
      Top             =   4710
      Width           =   915
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ubicación :"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   8685
      TabIndex        =   46
      Top             =   3900
      Width           =   885
   End
   Begin VB.Label Label16 
      Caption         =   "Cia: "
      Height          =   255
      Left            =   6240
      TabIndex        =   43
      Top             =   5480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "IVI:"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   3135
      TabIndex        =   36
      Top             =   3765
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Precio:"
      Height          =   315
      Left            =   300
      TabIndex        =   31
      Top             =   3780
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Factor 2:"
      Height          =   225
      Index           =   7
      Left            =   4890
      TabIndex        =   29
      Top             =   3990
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Factor 1:"
      Height          =   225
      Index           =   6
      Left            =   4890
      TabIndex        =   28
      Top             =   3660
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Talla :"
      Height          =   225
      Index           =   4
      Left            =   8430
      TabIndex        =   27
      Top             =   450
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Colección :"
      Height          =   225
      Index           =   2
      Left            =   8295
      TabIndex        =   26
      Top             =   5130
      Width           =   885
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sub Grupo :"
      Height          =   225
      Index           =   1
      Left            =   8385
      TabIndex        =   25
      Top             =   4350
      Width           =   945
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Left            =   4575
      Stretch         =   -1  'True
      ToolTipText     =   "Imagen del Artículo"
      Top             =   4350
      Width           =   1620
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Alterno :"
      Height          =   225
      Left            =   8925
      TabIndex        =   23
      Top             =   30
      Width           =   690
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Máximo :"
      Height          =   225
      Left            =   7185
      TabIndex        =   22
      Top             =   3690
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Mínimo :"
      Height          =   225
      Left            =   7185
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "%Util :"
      Height          =   225
      Left            =   6360
      TabIndex        =   17
      Top             =   3735
      Width           =   465
   End
End
Attribute VB_Name = "DetalleArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImageChange$
Dim Temp As Recordset
Dim s$
Dim Item As ListItem
Dim elmini$, elmax$
Dim ElColor$, LaTalla$
Private Sub Check3_Click()      ' color/talla s/n
If Check3 = 1 Then
     ECombo1(3).Visible = True
     ECombo1(4).Visible = True
     Text(6).Visible = True
     Text(7).Visible = True
Else
     ECombo1(3).Visible = False
     ECombo1(4).Visible = False
     Text(6).Visible = False
     Text(7).Visible = False
End If
End Sub

Private Sub Check5_Click()
    If Check5.Value = 1 Then
        Label16.Visible = True
        Text1.Visible = True
    Else
        Label16.Visible = False
        Text1.Visible = False
    End If
End Sub

Private Sub Command1_Click()
If Val(Tag) = 0 Then
          If Inserta Then
               CamBioPrecio
               If Check5.Value = 1 Then Call Inserta2
               Call Articulos.Columnas(True)
          Else
               Selecciona Text(1)
               Exit Sub    ' pa que no la limpie  !!
          End If
ElseIf Val(Tag) = 1 Then
          If Modifica Then
               CamBioPrecio
               If Check5.Value = 1 Then Call Inserta2
               Call Articulos.Columnas(True)
          Else
               Selecciona Text(0)
               Exit Sub    '' pa que no la limpie
          End If
End If
If Articulos.ConTinuO Then
    Call Limpia
    Tag = 0               ' o sea sigue agregando nuevos
Else
    Unload Me
End If
End Sub
Private Function Inserta() As Boolean
'If TLista(0) = "" Or TLista(1) = "" Or IsNull(TLista(0)) Or IsNull(TLista(1)) Then
     'MsgBox "Datos de listas en Blanco", vbCritical, "Agregar Artículo"
     'Inserta = False
     'Exit Function
'End If
     s = "insert into articulos(cia,c_articulo,d_articulo,"
     s = s + "p_compra,porc_desc1,porc_util,"
     s = s + "c_tipo_art,sino_comi,sino_impu,c_prove,"
     s = s + "costoex,minimo,maximo,ubicacion,aux,alterno,estado"
     s = s + ",marca,csubgrupo,ccoleccion,ccolor,ctalla,"
     s = s + "imagen,costoant,fact1,fact2,f_descDel,f_descAl)"
     If Check3 = 1 Then
          ElColor = ECombo1(3).Indice(1)
          LaTalla = ECombo1(4).Indice(1)
          elmini = Text(6)
          elmax = Text(7)
     Else
          ElColor = ""
          LaTalla = ""
          elmini = "0"
          elmax = "0"
     End If
     s = s + " values ('" + CiA + "','" + Text(0) + ElColor + LaTalla + "','" 'El codigo
     s = s + Text(1) + "'," 'La descripcion
     s = s & Doble(Text(2)) & "," 'El costo
     s = s & Doble(Text(3)) & "," 'El % de descuento
     s = s & Doble(Text(4)) & ",'" 'El % de utilidad
     s = s + ECombo1(0).Indice(1) + "'," 'El tipo de articulo
     s = s & Check1.Value & "," 'Si acepta comision o no
     s = s & Check2.Value & ",'" 'Si lleva impuesto o no
     s = s + ECombo1(5).Indice(1) + "',"   ' ECombo2.Indice(1) + "'," 'El provedor
     s = s & Doble(Text(5)) & "," 'El costo en $
     s = s + elmini + "," 'El minimo
     s = s + elmax + ",'" 'El maximo
     s = s + Text(8) + "','"    '      Text(8) + "','" 'ubicacion
     s = s + "*','"              ' Text(9) + " ','" 'auxiliar
     s = s + Text(11) + "',0,'" 'alterno
     s = s + "*','"        '  Text(12) + "','" 'marca
     s = s + "01','"    ' subgrupo
     s = s + "01','"    ' coleccion
     s = s + ElColor + "','"       ' color
     s = s + LaTalla + "'"        ' Talla
     s = s + ",'"   ' ???
     Dim N$
     If ImageChange <> "" Then
          If Dialog1.FilterIndex = 1 Then
               N = Text(0) + ".bmp"
          ElseIf Dialog1.FilterIndex = 2 Then
               N = Text(0) + ".jpg"
          ElseIf Dialog1.FilterIndex = 3 Then
               N = Text(0) + ".tif"
          End If
          ImageChange = ""
     End If
     s = s + N + "'," & Doble(Text(10)) & "," & Factor1 & "," & Factor2
     s = s + ",#" + Format(DTPicker1, "m/d/yyyy") + "#,#" + Format(DTPicker2, "m/d/yyyy") + "#)"
     If Procesa(s) Then
          If Dir(DirTrA + "Imagenes\", vbDirectory) <> "" And N <> "" Then
               Call SavePicture(Image1.Picture, DirTrA + "Imagenes\" + N)
          End If
          Call GBitacora(1, "Artículo: " + Text(1).Text)
          s = "insert into unidades(codart,descripcion,unidades,cia)"
          s = s + "values('" + Text(0) + ElColor + LaTalla + "','Unidad',1,'" + CiA + "')"
          Call Procesa(s)
          s = "insert into ListaPrecios(cia,codart,monto,codigo,descripcion)"
          s = s + "values('" + CiA + "','" + Text(0) + ElColor + LaTalla
          s = s + "'," & Doble(PreciO) & ",1,'Precio 1')"
          Call Procesa(s)
          Inserta = True
     End If
End Function
Private Function Inserta2() As Boolean
If TLista(0) = "" Or TLista(1) = "" Or IsNull(TLista(0)) Or IsNull(TLista(1)) Then
     MsgBox "Datos de listas en Blanco", vbCritical, "Agregar Artículo"
     Inserta2 = False
     Exit Function
End If
     s = "insert into articulos(cia,c_articulo,d_articulo,"
     s = s + "p_compra,porc_desc1,porc_util,"
     s = s + "c_tipo_art,sino_comi,sino_impu,c_prove,"
     s = s + "costoex,minimo,maximo,ubicacion,aux,alterno"
     s = s + ",marca,csubgrupo,ccoleccion,ccolor,ctalla,"
     s = s + "imagen,costoant,fact1,fact2,f_descDel,f_descAl)"
     If Check3 = 1 Then
          ElColor = ECombo1(3).Indice(1)
          LaTalla = ECombo1(4).Indice(1)
          elmini = Text(6)
          elmax = Text(7)
     Else
          ElColor = ""
          LaTalla = ""
          elmini = "0"
          elmax = "0"
     End If
     s = s + " values ('" + Trim(Text1) + "','" + Text(0) + ElColor + LaTalla + "','" 'El codigo
     s = s + Text(1) + "'," 'La descripcion
     s = s & Doble(Text(2)) & "," 'El costo
     '''' descuento fijo 30 % en cia 02     ENC 7/8/02
     s = s & IIf(Text1 = "02", 30, Doble(Text(3))) & "," 'El % de descuento
     s = s & Doble(Text(4)) & ",'" 'El % de utilidad
     s = s + ECombo1(0).Indice(1) + "'," 'El tipo de articulo
     s = s & Check1.Value & "," 'Si acepta comision o no
     s = s & Check2.Value & ",'" 'Si lleva impuesto o no
     s = s + ECombo1(5).Indice(1) + "',"   ' ECombo2.Indice(1) + "'," 'El provedor
     s = s & Doble(Text(5)) & "," 'El costo en $
     s = s + elmini + "," 'El minimo
     s = s + elmax + ",'" 'El maximo
     s = s + Text(8) + "','"    '      Text(8) + "','" 'ubicacion
     s = s + "*','"              ' Text(9) + " ','" 'auxiliar
     s = s + Text(11) + "','" 'alterno
     s = s + "*','"        '  Text(12) + "','" 'marca
     s = s + ECombo1(1).Indice(1) + "','"    ' subgrupo
     s = s + ECombo1(2).Indice(1) + "','"    ' coleccion
     s = s + ElColor + "','"       ' color
     s = s + LaTalla + "'"        ' Talla
     s = s + ",'"   ' ???
     Dim N$
     If ImageChange <> "" Then
          If Dialog1.FilterIndex = 1 Then
               N = Text(0) + ".bmp"
          ElseIf Dialog1.FilterIndex = 2 Then
               N = Text(0) + ".jpg"
          ElseIf Dialog1.FilterIndex = 3 Then
               N = Text(0) + ".tif"
          End If
          ImageChange = ""
     End If
     s = s + N + "'," & Doble(Text(10)) & "," & Factor1 & "," & Factor2
     s = s + ",#" + Format(DTPicker1, "m/d/yyyy") + "#,#" + Format(DTPicker2, "m/d/yyyy") + "#)"
     If Procesa(s, True) Then
          If Dir(DirTrA + "Imagenes\", vbDirectory) <> "" And N <> "" Then
               Call SavePicture(Image1.Picture, DirTrA + "Imagenes\" + N)
          End If
          Call GBitacora(1, "Artículo: " + Text(1).Text)
          s = "insert into unidades(codart,descripcion,unidades,cia)"
          s = s + "values('" + Text(0) + ElColor + LaTalla + "','Unidad',1,'" + Trim(Text1) + "')"
          Call Procesa(s, True)
          s = "insert into ListaPrecios(cia,codart,monto,codigo,descripcion)"
          s = s + "values('" + Trim(Text1) + "','" + Text(0) + ElColor + LaTalla
          s = s + "'," & Doble(PreciO) & ",1,'Precio 1')"
          Call Procesa(s, True)
          Inserta2 = True
     End If
End Function
Private Function Procesa(SQL As String, Optional NoigaNaa As Boolean) As Boolean
     On Error GoTo Errores
     If IsNull(NoigaNaa) Then NoigaNaa = False
     Procesa = True
     DatOS.Execute SQL, 128
Errores:
     If err.Number <> 0 Then
        If NoigaNaa Then
           Procesa = False
           Resume Next
        End If
     End If
     If err.Number = 3022 Then
          MsgBox "Crea una entrada duplicada !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number = 3315 Then
          MsgBox "Al menos uno de los campos vacíos es requerido !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number = 3464 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Procesa"
     End If
     On Error GoTo 0
End Function
Private Function Modifica() As Boolean
     Dim CostoPrev@
     Dim UtilPrev@
     Dim Costo@
     Dim Util@
     Costo = Doble(Text(2))
     Util = Doble(Text(4))
     CostoPrev = Text(2).Tag
     UtilPrev = Text(4).Tag
     If Check3 = 1 Then
          ElColor = ECombo1(3).Indice(1)
          LaTalla = ECombo1(4).Indice(1)
          elmini = Text(6)
          elmax = Text(7)
     Else
          ElColor = ""
          LaTalla = ""
          elmini = "0"
          elmax = "0"
     End If
     s = "update Articulos set c_articulo='" + Text(0) + ElColor + LaTalla
     s = s + "',d_articulo='" + Text(1)
     s = s + "',p_compra=" & Costo
     s = s + ",porc_desc1=" & Doble(Text(3))
     s = s + ",porc_util=" & Util
     s = s + ",c_tipo_art='" & TLista(0).Text
     s = s + "',sino_comi=" + IIf(Check1.Value = 1, "1", "0")
     s = s + ",sino_impu=" + IIf(Check2.Value = 1, "1", "0")
     s = s + ",c_prove='" + ECombo1(5).Indice(1)
     s = s + "',costoex=" & Doble(Text(5))
     s = s + ",minimo=" + elmini
     s = s + ",maximo=" + elmax
     s = s + ",ubicacion='" + Text(8)
     s = s + "',costoant=" & CostoPrev
     s = s + ",aux='*"         '   + Text(9)
     s = s + "',alterno='" + Text(11)
     s = s + "',marca='*'"      '    + Text(12)
     s = s + ",cSubGrupo ='" + TLista(1).Text + "'"
     s = s + ",CColeccion='" + TLista(2).Text + "'"
     s = s + ",CColor    ='" + ElColor + "'"
     s = s + ",CTalla    ='" + LaTalla + "'"
     s = s + ",fact1=" & Factor1
     s = s + ",fact2=" & Factor2
     s = s + ",f_descdel=#" + Format(DTPicker1, "m/d/yyyy") + "#"
     s = s + ",f_descal=#" + Format(DTPicker2, "m/d/yyyy") + "#"
     s = s + ",imagen='"
     Dim N$
     If ImageChange <> "" Then
          If Dialog1.FilterIndex = 1 Then
               N = Text(0) + ".bmp"
          ElseIf Dialog1.FilterIndex = 2 Then
               N = Text(0) + ".jpg"
          ElseIf Dialog1.FilterIndex = 3 Then
               N = Text(0) + ".tif"
          End If
          ImageChange = ""
     End If
     s = s + N + "' where cia='" + CiA + "' and c_articulo='" + Text(0).Tag + "'"
     If Procesa(s) Then
          If Dir(DirTrA + "Imagenes\", vbDirectory) <> "" And N <> "" Then
               Call SavePicture(Image1.Picture, DirTrA + "Imagenes\" + N)
          End If
          Call GBitacora(2, "Articulo : " + Text(1))
          If App.Comments = "Dafesa" And (Costo <> CostoPrev) Then
               s = "El costo del artículo ha sido variado, desea actualizar"
               s = s + Chr(13) + "los precios de venta en base al nuevo costo ?"
               If MsgBox(s, 36, Caption) = 6 Then
                    If Not PreciosVenta(Text(0), Costo) Then Exit Function
               End If
          End If
          Modifica = True
     End If
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Command3_Click()
     On Error GoTo Errores
     Dialog1.CancelError = True
     Dialog1.DialogTitle = "Abrir archivo de imágen"
     Dialog1.InitDir = CurDir
     Dialog1.Filter = "Bitmap files|*.bmp|JPEG files|*.jpg|TIFF files|*.tif"
     Dialog1.ShowOpen
     If Dialog1.FileName <> "" Then
          Image1.Picture = LoadPicture(Dialog1.FileName)
     End If
     ImageChange = Dialog1.FileName
Errores:
     If err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Error"
     End If
End Sub
Private Sub Command4_Click()
     Image1.Picture = LoadPicture("")
     ImageChange = ""
End Sub
Private Sub Command5_Click()
MousePointer = 11
Command5.Enabled = False
If Check4 Then
   If ImpreBarra1(Text(0) + ElColor + LaTalla, Text(1), Doble(PreciO), Doble(TEtiquetas)) Then
      MsgBox "Etiquetas OutLeT Listas", vbInformation, Printer.DeviceName
   End If
ElseIf ImpreBarra(Text(0) + ElColor + LaTalla, Text(1), Doble(PreciO), Doble(TEtiquetas)) Then
     MsgBox "Etiquetas Listas", vbInformation, Printer.DeviceName
End If
MousePointer = 0
Command5.Enabled = True
End Sub
Private Sub ECombo1_Click(Index As Integer)
If Index < 6 Then    ''' ojo se cae en color talla y estilo ????  24/7/02 enc
     TLista(Index) = ECombo1(Index).Indice(1)
End If
End Sub

Private Sub Factor1_Change()
PreciO = Format(CalcU_PreciO(Doble(Text(5)), Doble(Factor1), Doble(Factor2)), "###,###")
PrecioIVI = Format(Doble(PreciO) * (1 + IV), "###,###")
'Text(2) = Format(Doble(Text(5)) * Tipo_Cambio * Factor1, "####,###.##")
End Sub

Private Sub Factor2_Change()
PreciO = Format(CalcU_PreciO(Doble(Text(5)), Doble(Factor1), Doble(Factor2)), "###,###")
PrecioIVI = Format(Doble(PreciO) * (1 + IV), "###,###")
End Sub
Private Sub Form_Load()
     s = "select c_tipo_art,d_tipo_art from tipos "
     s = s + "where cia='" + CiA + "' order by d_tipo_art"
     Set Temp = DatOS.OpenRecordset(s)
     Do Until Temp.EOF
          Call ECombo1(0).AddItem(Temp!d_tipo_art, Trim(Temp!c_tipo_art))
          Temp.MoveNext
     Loop

     ECombo1(3).Visible = False
     ECombo1(4).Visible = False
     Refresh
End Sub
Public Sub DatosArticulo(Codigo$)
     On Error GoTo Errores
     Dim I%
     Dim Articulos As Recordset
     s = "select * from Articulos where cia='" + CiA + "' and c_articulo='" + Codigo + "'"
     Set Articulos = DatOS.OpenRecordset(s)
     If Articulos.EOF() Then
        Call Limpia(True)
        Exit Sub             ' Articulo  nuevo
     End If
     Tag = 1
     Text(0) = Articulos!c_articulo     ' Mid(Articulos!c_articulo, 1, 9)
     Text(0).Tag = Articulos!c_articulo
     Text(1) = Nulo(Articulos!d_articulo)
     Text(2) = FormatNumber(Articulos!p_compra, DeCiMaleS)
     Text(2).Tag = Nulo(Articulos!p_compra)
     Text(3) = Nulo(Articulos!porc_desc1)
     Text(4) = Nulo(Articulos!porc_util)
     Text(4).Tag = Nulo(Articulos!porc_util)
     Text(5) = FormatNumber(Articulos!CostoEx, DeCiMaleS)
     Text(6) = Nulo(Articulos!Minimo)
     Text(7) = Nulo(Articulos!Maximo)
     DTPicker1 = Nulo(Articulos!f_descDel)
     DTPicker2 = Nulo(Articulos!f_descAl)
     Text(8) = Nulo(Articulos!ubicacion)
     'Text(9) = Nulo(Articulos!Aux)
     Text(10) = FormatNumber(Articulos!CostoAnt, DeCiMaleS)
     Text(11) = Nulo(Articulos!Alterno)
     
     For I = 0 To ECombo1(5).ListCount - 1
          If Trim(ECombo1(5).List(I, 1)) = Articulos!c_prove Then
               ECombo1(5).ListIndex = I
               Exit For
          End If
     Next
     
     Factor1.Text = Nulo(Articulos!fact1)
     Factor2.Text = Nulo(Articulos!fact2)
     'Text(12) = Nulo(Articulos!Marca)
     Check1.Value = Nulo(Articulos!sino_comi)
     Check2.Value = Nulo(Articulos!sino_impu)
     ''If Mid(Articulos!c_articulo, 10, 1) <> "" Then
     ''     Check3.Value = 1
     ''     ECombo1(3).Visible = True
     ''     ECombo1(4).Visible = True
     ''     Text(6).Visible = True
     ''     Text(7).Visible = True
     ''Else
          Check3.Value = 0
          ECombo1(3).Visible = False
          ECombo1(4).Visible = False
          Text(6).Visible = False
          Text(7).Visible = False
     ''End If
     Dim eldato$, x As Integer
     For x = 0 To 4
          If x = 0 Then
               eldato = Nulo(Articulos!c_tipo_art)
          ElseIf x = 1 Then
               eldato = Nulo(Articulos!CSubGrupo)
          ElseIf x = 2 Then
               eldato = Nulo(Articulos!cColeccion)
          ElseIf x = 3 Then
               eldato = Nulo(Articulos!ccolor)
          ElseIf x = 4 Then
               eldato = Nulo(Articulos!ctalla)
          ElseIf x = 5 Then
               eldato = Nulo(Articulos!CEstilo)
          End If
          For I = 0 To ECombo1(x).ListCount - 1
               If Trim(ECombo1(x).List(I, 1)) = Trim(eldato) Then
                    ECombo1(x).ListIndex = I
                    Exit For
               End If
          Next I
     Next x
     If Not IsNull(Articulos!Imagen) And Articulos!Imagen <> "" Then
          Dim Ruta$
          Ruta = DirTrA + "Imagenes\" + Articulos!Imagen
          If Dir(Ruta) <> "" Then
               Image1.Picture = LoadPicture(Ruta)
          Else
              ' MsgBox "Imagen no encontrada !", 48, "Detalle de Artículo"
          End If
     End If
     s = "select * from ListaPrecios where cia='" + CiA + "' and codart='" + Codigo + "'"
     s = s + " and codigo=1"
     Set Articulos = DatOS.OpenRecordset(s)
     If Articulos.EOF() Then
        PreciO = "No Existe"
        PreciO.Tag = -1000000
     Else
        PreciO = Format(Articulos!monto, "###,###")
        PreciO.Tag = Articulos!monto
     End If
Errores:
     If err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Carga datos Articulo"
          Resume Next
     End If
     On Error GoTo 0
End Sub
 
Private Sub Precio_Change()
    PrecioIVI = Doble(Doble(PreciO) * (1 + IV))
End Sub
Private Sub PrecioIVI_LostFocus()
    PreciO = Doble(Doble(PrecioIVI) / (1 + IV))
End Sub

Private Sub TEtiquetas_Change()
     Command5.Enabled = Doble(TEtiquetas) > 0
End Sub

Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case Index
     Case 2 To 7
          Select Case KeyAscii
          Case 8, 46, 48 To 57
          Case Else
               KeyAscii = 0
          End Select
     Case Else
          Select Case KeyAscii
          Case 10, 13, 39
               KeyAscii = 0
          End Select
     End Select
End Sub
Private Sub Limpia(Optional EsNuevo As Boolean)
     EsNuevo = IIf(IsNull(EsNuevo), False, EsNuevo)
     For I = 1 To Text.Count - 1      ' OJO EL 0 EL CZRTIICULO NO LO LIMPIA
          If I <> 9 Then Text(I).Text = ""
     Next
     For I = 0 To 5
          ECombo1(I).ListIndex = -1
     Next
     If Not EsNuevo Then
          Text(0) = ""
          Text(0).SetFocus
     End If
     Factor1 = Factor_1
     Factor2 = Factor_2
     'ECombo2.ListIndex = -1
     Check1.Value = 0
     Check2.Value = 1
     Check3.Value = 0
     Text(11) = ""
     ECombo1(3).Enabled = False   'color
     ECombo1(4).Enabled = flase    'talla
     PreciO.Tag = -1000000
     Image1.Picture = LoadPicture("")
     ImageChange = False
End Sub
Private Sub Text_LostFocus(Index As Integer)
Select Case Index
Case 0           ' el articulo
     If Val(Tag) = 0 Then         ' es nuevo
        Text(0) = Trim(Text(0))
        DatosArticulo (Text(0))
     End If
Case 5                 ' el fob $
     PreciO = Format(CalcU_PreciO(Doble(Text(5)), Doble(Factor1), Doble(Factor2)), "###,###")
     PrecioIVI = Format(Doble(PreciO) * (1 + IV), "###,###")
     Text(2) = Format(Doble(Text(5)) * Tipo_Cambio * Factor1, "####,###.##")
End Select
End Sub
Private Sub TLista_LostFocus(Index As Integer)
TLista(Index) = Trim(TLista(Index))
Dim Hay As Boolean
Hay = False
For I = 0 To ECombo1(Index).ListCount - 1
     If Trim(ECombo1(Index).List(I, 1)) = TLista(Index) Then
          ECombo1(Index).ListIndex = I
          Hay = True
          Exit For
     End If
Next I
If Not Hay Then
   ECombo1(Index).ListIndex = -1
End If
End Sub
Private Sub CamBioPrecio()
If PreciO.Tag <> Doble(PreciO) Then
     s = "Update ListaPrecios set monto=" & Doble(PreciO)
     s = s + " Where cia='" + CiA + "' and codart='" + Text(0) + ElColor + LaTalla
     s = s + "' and codigo=1"
     DatOS.Execute (s)
End If
End Sub
Private Function ImpreBarra(elCarticuLo$, elDarticulO$, elMonTo As Double, LasCopias%) As Boolean
On Error GoTo CMDErr
Printer.Copies = LasCopias
Printer.ScaleMode = vbMillimeters
Printer.Font = "Tahoma"       '"Courier New"
Printer.FontSize = 10
Printer.Print ""
Barcode1.PrinterScaleMode = Printer.ScaleMode
'''''''''''''''''''''''''''''''''''''''''
Dim Start%, Xx$
Start = 2
LabelTop = 1
Barcode1.PrinterLeft = 2        'Start
Barcode1.PrinterTop = 10        'LabelTop
Barcode1.PrinterWidth = 35
Barcode1.PrinterHeight = 8
Barcode1.Caption = "*"
Barcode1.Caption = elCarticuLo
Barcode1.PrinterHDC = Printer.hDC
' La Barra
Printer.FontSize = 10
Printer.CurrentX = Barcode1.PrinterLeft
Printer.CurrentY = LabelTop     ' Barcode1.PrinterTop + Barcode1.PrinterHeight + 2
Printer.Print Barcode1.Displayed
'El Precio
Printer.FontSize = 8
Printer.CurrentX = Start + 30
Printer.CurrentY = LabelTop
Xx = "¢" + Format(elMonTo, "###,###")
Printer.Print Xx
'El precio con iv
Printer.FontBold = True
Printer.FontSize = 12
Printer.CurrentX = Start + 17
Printer.CurrentY = LabelTop + 4
Xx = "¢" + Format(elMonTo * (IV + 1), "###,###") + " IVI"
Printer.Print Xx
'La fecha de hoy
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Start + 37
Printer.CurrentY = 16
Printer.Print Format(Date, "mmyy") + " "
'La descripcion
Printer.FontSize = 8
Printer.CurrentX = Start
Printer.CurrentY = LabelTop + 19
Xx = Mid(elDarticulO, 1, 48)
Printer.Print Xx
'
Printer.FontBold = False
Printer.FontSize = 10
'Printer.NewPage
Printer.EndDoc
ImpreBarra = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CMDErr:
If err.Number = 482 Then
   s = "La impresora no está lista!, verifique que esté encendida y con ETIQUETAS."
   If MsgBox(s, 21, Printer.DeviceName) = 4 Then
      Resume
   Else
      Exit Function
   End If
ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
End If
On Error GoTo 0
End Function

Private Function ImpreBarra1(elCarticuLo$, elDarticulO$, elMonTo As Double, LasCopias%) As Boolean
On Error GoTo CMDErr
Printer.Copies = LasCopias
Printer.ScaleMode = vbMillimeters
Printer.Font = "Arial"         '' "Courier New"  ''  "Tahoma"
Printer.FontSize = 10
Printer.Print ""
Barcode1.PrinterScaleMode = Printer.ScaleMode
'''''''''''''''''''''''''''''''''''''''''
Dim Start%, Xx$, ArriBa As Integer, ConDescu As Double
Start = 2
LabelTop = 1
ArriBa = 33
Barcode1.Direction = msDTopToBottom
Barcode1.PrinterLeft = 0        'Start
Barcode1.PrinterTop = 42        'LabelTop
Barcode1.PrinterWidth = 6
Barcode1.PrinterHeight = 35
Barcode1.Caption = "*"
Barcode1.Caption = elCarticuLo
Barcode1.PrinterHDC = Printer.hDC
'La descripcion
Printer.FontSize = 8
Printer.CurrentX = Start
Printer.CurrentY = ArriBa
Xx = Mid(elDarticulO, 1, 48)
Printer.Print Xx
' La Barra EL CODIGO
Printer.FontBold = True
Printer.CurrentX = Start
Printer.CurrentY = ArriBa + 5
Printer.Print Barcode1.Displayed
Printer.FontBold = False
''
Printer.CurrentX = 18
Printer.CurrentY = ArriBa + 10
Printer.Print "Precio"
Printer.CurrentX = 18
Printer.CurrentY = ArriBa + 14
Printer.Print "Regular IVI"
'El Precio
Printer.FontSize = 10
Printer.CurrentX = 14
Printer.CurrentY = ArriBa + 18
Xx = "¢" + Format(elMonTo * (IV + 1), "###,###.00")
Printer.Print Xx
' EL DESCU
Printer.FontBold = True
Printer.FontSize = 14
Printer.CurrentX = 17
Printer.CurrentY = ArriBa + 23
Xx = Format(Text(3), "##") + "%"
Printer.Print Xx
'''''''
ArriBa = ArriBa + 3
Printer.FontBold = False
Printer.FontSize = 10
Printer.CurrentX = 15
Printer.CurrentY = ArriBa + 27
Printer.Print "NUESTRO"
Printer.CurrentX = 15
Printer.CurrentY = ArriBa + 31
Printer.Print "PRECIO IVI"
'El precio con DESCUENTO
Printer.FontBold = True
Printer.FontSize = 12
Printer.CurrentX = 11
Printer.CurrentY = ArriBa + 35
ConDescu = Val(Text(3)) / 100   '' el descuento   30 %
ConDescu = elMonTo * ConDescu   ''  2500*.3 = 750
ConDescu = elMonTo - ConDescu   ''  2500 - 750
Xx = "¢" + Format(ConDescu * (IV + 1), "###,###.00")
Printer.Print Xx
'La fecha de hoy
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = 2
Printer.CurrentY = ArriBa + 42
Printer.Print Format(Date, "mmyy") + " "
'
Printer.FontBold = False
Printer.FontSize = 10
'Printer.NewPage
Printer.EndDoc
ImpreBarra1 = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CMDErr:
If err.Number = 482 Then
   s = "La impresora no está lista!, verifique que esté encendida y con ETIQUETAS."
   If MsgBox(s, 21, Printer.DeviceName) = 4 Then
      Resume
   Else
      Exit Function
   End If
ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
End If
On Error GoTo 0
End Function

