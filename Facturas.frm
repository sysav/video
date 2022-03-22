VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{B4409115-5405-11D3-943D-0080AD4162AE}#1.0#0"; "ECOMBO.OCX"
Begin VB.Form Facturas 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Facturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   10965
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   225
      Left            =   2175
      TabIndex        =   71
      Top             =   5940
      Width           =   810
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Left            =   1110
      TabIndex        =   70
      Top             =   5955
      Width           =   660
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   240
      Left            =   1455
      TabIndex        =   69
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nueva"
      Height          =   615
      Left            =   3825
      Picture         =   "Facturas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   5640
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Agregar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   5010
      Picture         =   "Facturas.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5640
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Actualizar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6180
      Picture         =   "Facturas.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   5610
      Width           =   1155
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Borrar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7380
      Picture         =   "Facturas.frx":0748
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   5610
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8565
      Picture         =   "Facturas.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   5610
      Width           =   1155
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   9750
      Picture         =   "Facturas.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5610
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Términos de la Factura"
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
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   3330
      TabIndex        =   56
      Top             =   4020
      Width           =   7635
      Begin VB.TextBox pedido 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5610
         TabIndex        =   65
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   57
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Número Pedido:"
         Height          =   225
         Left            =   4260
         TabIndex        =   66
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Condiciones:"
         Height          =   225
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   1260
      End
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   795
      Left            =   930
      TabIndex        =   51
      Top             =   2970
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1402
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   13
      Left            =   1110
      MaxLength       =   15
      TabIndex        =   52
      ToolTipText     =   "Digite con cuanto paga el cliente !"
      Top             =   5220
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   1545
      Left            =   15
      ScaleHeight     =   1485
      ScaleWidth      =   10845
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   -15
      Width           =   10905
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   15
         Left            =   7395
         MaxLength       =   15
         TabIndex        =   68
         ToolTipText     =   "El código de cliente"
         Top             =   750
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808000&
         Caption         =   "Factura Exenta"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   9150
         TabIndex        =   8
         ToolTipText     =   "Determina si esta factura es exenta en su totalidad"
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   12
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   50
         Top             =   30
         Width           =   1545
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Proforma :"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2430
         TabIndex        =   49
         ToolTipText     =   "Traer una proforma generada previamente"
         Top             =   60
         Width           =   1155
      End
      Begin EnhancedCombo.ECombo ECombo 
         Height          =   345
         Index           =   2
         Left            =   5865
         TabIndex        =   48
         ToolTipText     =   "Como se cancela esta fcatura"
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   11
         Left            =   1890
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Comentarios adicionales"
         Top             =   1110
         Width           =   7065
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   9750
         MultiLine       =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Monto por flete"
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   9
         Left            =   870
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "El plazo de esta factura"
         Top             =   1110
         Width           =   465
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   8
         Left            =   2340
         TabIndex        =   41
         ToolTipText     =   "El nombre al cual sale la factura"
         Top             =   750
         Width           =   5040
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   10170
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   390
         Width           =   645
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   2
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Fecha de la factura"
         Top             =   30
         Width           =   1245
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   1
         Left            =   870
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "El código de cliente"
         Top             =   750
         Width           =   1455
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   3
         Left            =   5850
         TabIndex        =   3
         ToolTipText     =   "Quien hace esta factura"
         Top             =   390
         Width           =   615
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   0
         Left            =   870
         MaxLength       =   2
         TabIndex        =   2
         ToolTipText     =   "La bodega de la cual se procede a vender"
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   67
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label NomBod 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1320
         TabIndex        =   47
         Top             =   390
         Width           =   3465
      End
      Begin VB.Label NomVend 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   6480
         TabIndex        =   46
         Top             =   390
         Width           =   2865
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1395
         TabIndex        =   45
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flete :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   9030
         TabIndex        =   43
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   300
         TabIndex        =   42
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc (%) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   9405
         TabIndex        =   40
         Top             =   450
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   39
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pago :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5355
         TabIndex        =   38
         Top             =   90
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   165
         TabIndex        =   37
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   8280
         TabIndex        =   36
         Top             =   90
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4935
         TabIndex        =   35
         Top             =   450
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   135
         TabIndex        =   34
         Top             =   480
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   915
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   10845
      TabIndex        =   26
      Top             =   1560
      Width           =   10905
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   14
         Left            =   7470
         MaxLength       =   15
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   1710
      End
      Begin EnhancedCombo.ECombo ECombo 
         Height          =   345
         Index           =   5
         Left            =   3180
         TabIndex        =   12
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "La cantidad a facturar"
         Top             =   420
         Width           =   885
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   8490
         MaxLength       =   5
         TabIndex        =   14
         Top             =   420
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         Caption         =   "IV"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6990
         TabIndex        =   13
         Top             =   480
         Width           =   525
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   4
         Left            =   1095
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "El código del articulo"
         Top             =   60
         Width           =   1695
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   345
         Left            =   2010
         TabIndex        =   27
         Top             =   425
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   609
         _Version        =   327681
         BuddyControl    =   "Text(5)"
         BuddyDispid     =   196620
         BuddyIndex      =   5
         OrigLeft        =   9150
         OrigTop         =   1980
         OrigRight       =   9345
         OrigBottom      =   2295
         Max             =   999999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6990
         TabIndex        =   55
         Top             =   150
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label NomLOte 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   9195
         TabIndex        =   54
         Top             =   60
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label DescArt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2820
         TabIndex        =   44
         ToolTipText     =   "La descripción del artículo"
         Top             =   60
         Width           =   4065
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   345
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2295
         TabIndex        =   32
         Top             =   510
         Width           =   840
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4365
         TabIndex        =   31
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5025
         TabIndex        =   30
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   270
         TabIndex        =   15
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC(%) :"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7560
         TabIndex        =   29
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5025
         TabIndex        =   28
         ToolTipText     =   "El precio de venta"
         Top             =   420
         Width           =   1860
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1605
      Left            =   0
      TabIndex        =   16
      Top             =   2460
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Unid"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Precio"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SubTotal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Desc(%)"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descuento"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "IV(%)"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Impuesto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Total"
         Object.Width           =   2117
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   300
      Top             =   5670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":0BD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label PagaCon 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Paga con :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   255
      TabIndex        =   53
      Top             =   5250
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Impuesto :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   25
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Subtotal :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   330
      TabIndex        =   24
      Top             =   4200
      Width           =   750
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descuento :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Descuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1110
      TabIndex        =   22
      ToolTipText     =   "El descuento de la factura"
      Top             =   4500
      Width           =   2175
   End
   Begin VB.Label Subtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1110
      TabIndex        =   21
      ToolTipText     =   "El subtotal de la factura"
      Top             =   4140
      Width           =   2175
   End
   Begin VB.Label Impuesto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1110
      TabIndex        =   20
      ToolTipText     =   "El impuesto de la factura"
      Top             =   4860
      Width           =   2175
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   3525
      TabIndex        =   19
      Top             =   4935
      Width           =   1230
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Left            =   4740
      TabIndex        =   18
      ToolTipText     =   "El total de la factura"
      Top             =   4920
      Width           =   6225
   End
End
Attribute VB_Name = "Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TieneVuelto%
Dim Bodegas As Recordset
Dim Lotes As Recordset
Dim Temp As Recordset
Dim Vendedores As Recordset
Public Cliente As Recordset
Dim Genero As Boolean
Public articulos As Recordset
Public Leyenda$, La_cedula$
Dim S$
Dim UsoProforma As Boolean
Dim Item As ListItem
Private Sub Check2_Click()
     Static Valor%
     If ListView1.ListItems.Count > 0 And Valor = 0 Then
          S = "Todos los articulos incluidos se perderán, desea continuar ?"
          If MsgBox(S, 36, "Facturación") = 6 Then
               ListView1.ListItems.Clear
               Call CalculaSaldo
          Else
               Valor = 1
               Check2.Value = IIf(Check2.Value = 0, 1, 0)
               Valor = 0
          End If
     End If
     If Text(4).Enabled Then Text(4).SetFocus
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     If ListView1.ListItems.Count > 0 And Command5.Caption = "&Guardar" Then
          S = "No se ha guardado la factura actual, desea salir ?"
          If MsgBox(S, 36, Caption) <> 6 Then Cancel = True
     End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Posiciona(Me, 1)
End Sub
Private Sub Command7_Click()
     S = "select n_factura as Número,clientes.nombre as Cliente,"
     S = S + "f_factura as Fecha "
     S = S + "from proformas left join clientes "
     S = S + "on clientes.cia=proformas.cia "
     S = S + "and clientes.codigo=proformas.c_cliente "
     S = S + "where proformas.estado=0 order by clientes.nombre"
     Call Lista.Carga(ByVal Text(12), S, "Proformas")
     Lista.Show 1
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     Dim Titulo$
     If KeyCode = vbKeyF3 And Index = 0 Then
          S = "select c_bodega as Código,d_bodega as Nombre "
          S = S + "from bodegas where cia='" + CiA + "' order by d_bodega"
          Titulo = "Bodegas"
     ElseIf KeyCode = vbKeyF3 And Index = 1 Then
          S = "select codigo as Código,Nombre "
          S = S + "from clientes where cia='" + CiA + "' order by nombre"
          Titulo = "Clientes"
     ElseIf KeyCode = vbKeyF3 And Index = 3 Then
          S = "select codigo as Código,Nombre "
          S = S + "from agentes where cia='" + CiA + "' order by nombre"
          Titulo = "Vendedores"
     ElseIf KeyCode = vbKeyF3 And Index = 4 Then
          S = "select c_articulo as Código,d_articulo as Descripción "
          S = S + "from articulos where cia='" + CiA + "' order by d_articulo"
          Titulo = "Artículos"
     ElseIf KeyCode = vbKeyF3 And Index = 14 Then
          S = "select codigo as Código,Nombre "
          S = S + "from lotes where cia='" + CiA + "' order by nombre"
          Titulo = "Lotes"
     ElseIf KeyCode = vbKeyF7 And Index = 4 Then
          Call Existencias.Carga(Text(4))
          Exit Sub
     ElseIf KeyCode = vbKeyF7 And Index = 14 Then
          If DescArt <> "" Then
               S = "select codigo as Código,Nombre,Vence,Existencia "
               S = S + "from lotes left join existencias "
               S = S + "on lotes.cia=existencias.cia "
               S = S + "and lotes.codigo=existencias.lote "
               S = S + "where existencias.cia='" + CiA
               S = S + "' and existencias.c_articulo='" + Text(4)
               S = S + "' and existencias.c_bodega='" + Text(0)
               S = S + "' order by vence"
               Titulo = "Existencias por Lote"
          Else
               MsgBox "Debe seleccionar un articulo !", 48, Caption
          End If
     Else
          Exit Sub
     End If
     Call Lista.Carga(Text(Index), S, Titulo)
     Lista.Show 1
End Sub
Private Sub Command1_Click()
     If ListView1.ListItems.Count > 0 And Command5.Caption = "&Guardar" Then
          If MsgBox("Desea guardar la factura actual ?", 36, "Facturación") = 6 Then
               Command5_Click
               Exit Sub
          End If
     End If
     Call Limpia
     Dim Temp As Recordset
     S = "select * from parametros where cia='" + CiA + "'"
     Set Temp = DatOS.OpenRecordset(S, 4)
     If Not Temp.EOF Then
          Label7.Caption = Temp!Consec
          If Label7 = "" Then Label7 = "00000001"
          TieneVuelto = Temp!Vuelto
          Leyenda = Nulo(Temp!rotulo)
     Else
          S = "No se han definido los parámetros de la compañía actual !"
          MsgBox S, 48, Caption
     End If
     Picture1.Enabled = True
     Picture2.Enabled = True
     ListView1.Enabled = True
     Bodegas.FindFirst "default=1"
     If Not Bodegas.NoMatch Then Text(0) = Bodegas!c_bodega
     ECombo(2).ListIndex = Val(ECombo(2).Tag)
     Text(2) = Format(Date, "dd/mm/yyyy")
     If ECombo(2).Enabled Then ECombo(2).SetFocus
End Sub
Private Sub Command2_Click()
     Call Llena(1)
End Sub
Private Sub Command3_Click()
     Call Llena(2)
End Sub
Private Sub Llena(Tipo As Integer)
     If ListView1.ListItems.Count <= 18 Then
          If Nulo(Cliente!credito) = 1 Then
               S = "El cliente seleccionado tiene el crédito bloqueado !"
               MsgBox S, 16, Caption
               Selecciona Text(1)
               Exit Sub
          End If
          Dim Cant@
          Cant = Doble(Text(5)) * Doble(Label16)
          If Not AgregaItem(Text(4), DescArt.Caption, Cant, _
               Doble(Label16.Caption), Doble(Text(6)), Check1.Value, _
               Doble(Label22.Caption), Tipo, articulos!Minimo, _
               articulos!p_compra, Text(14), 0) Then
               Exit Sub
          End If
          Command3.Enabled = False
          Command4.Enabled = False
          DescArt.Caption = ""
          Text(4) = ""
          Text(6) = ""
          Check1.Value = 0
          Text(4).SetFocus
     Else
          MsgBox "Ya se alcanzó el máximo de artículos por factura !", 64, "Facturación"
     End If
End Sub
Private Sub Command4_Click()
     Set Item = ListView1.SelectedItem
     ListView2.ListItems.Remove Item.Index
     ListView1.ListItems.Remove Item.Index
     Command3.Enabled = False
     Command4.Enabled = False
     Call CalculaSaldo
     Text(14) = ""
     Text(4) = ""
     Text(4).SetFocus
     Set Item = Nothing
End Sub
Private Sub Command5_Click()
    'Pide clave para autorizar credito
    'If Command5.Caption = "&Guardar" Then
    '   Avail = ValidaSaldo(Total)
    '   If Avail <= 0 Then
    '      Clave.Show 1
    '      If claveD <> "SI" Then
    '          Exit Sub
    '      End If
    '   End If
    'End If
  
     If Command5.Caption = "&Guardar" Then
          If Nulo(Cliente!credito) = 1 Then
               S = "El cliente seleccionado tiene el crédito bloqueado !"
               MsgBox S, vbCritical, Caption
               Text(1).SetFocus
               Exit Sub
          End If
          EspaCio.BeginTrans
          If Inserta Then
               Call ActualizaConsec
               S = "delete from proformas where cia='" + CiA + "' and n_factura='"
               S = S + Format(Text(12), "00000000") + "'"
               DatOS.Execute S, 128
               S = "delete from desgprof where cia='" + CiA + "' and n_factura='"
               S = S + Format(Text(12), "00000000") + "'"
               DatOS.Execute S, 128
               EspaCio.CommitTrans
               If TieneVuelto = 1 Then
                    Dim Paga#
                    Dim mTotal#
                    Paga = Doble(Text(13))
                    mTotal = Doble(Total.Caption)
                    If Paga > 0 Then
                         Do While True
                              If Paga > mTotal Then
                                   Call Vuelto.Carga(Paga - mTotal)
                                   Exit Do
                              Else
                                   S = "El pago debe ser mayor al total de la factura !"
                                   S = S + Chr(13) + "Digite el monto con el cual se cancela :"
                                   Paga = Doble(InputBox(S, "Facturación", 0))
                              End If
                         Loop
                    End If
               End If
               Command4.Enabled = False
               ListView1.Enabled = False
               Picture1.Enabled = False
               Picture2.Enabled = False
               Command5.Caption = "&Imprimir"
               Command5.Picture = ImageList1.ListImages(1).Picture
          Else
               EspaCio.Rollback
          End If
     Else
          MousePointer = 11
          If Not Impresion(ReadKey("factura")) Then
               MsgBox "Error al imprimir la factura !", 48, "Facturación"
          Else
               Call Limpia
          End If
          MousePointer = 0
     End If
End Sub
Private Sub Command6_Click()
     Unload Me
End Sub
Private Sub ECombo_Click(Index As Integer)
     If Index = 5 Then 'Los precios
          If ECombo(5).ListIndex > -1 Then
               Label16.Caption = ECombo(5).Indice(1)
               Label22.Caption = FormatNumber(ECombo(5).Indice(2), DeCiMaleS)
          Else
               Label16.Caption = ""
               Label22.Caption = ""
          End If
          Call Valida
     ElseIf Index = 2 Then
          If ECombo(2).Indice(5) = 1 Then
               PagaCon.Visible = True
               Text(13).Visible = True
          Else
               PagaCon.Visible = False
               Text(13).Visible = False
          End If
     End If
End Sub
Private Sub Form_Load()
     Call Posiciona(Me, 0)
     'Bodegas
     S = "select d_bodega,c_bodega,default from bodegas "
     S = S + "where cia='" + CiA + "' order by d_bodega"
     Set Bodegas = DatOS.OpenRecordset(S)
     'Tipos de pago
     S = "select d_tipo_fac,c_tipo_fac,cxc,c_conta,default,dolares "
     S = S + "from tipofact where cia='" + CiA + "' order by d_tipo_fac"
     Call CargaCombo(ECombo(2), S)
     'Vendedores
     S = "select nombre,codigo from agentes where cia='" + CiA + "' order by nombre"
     Set Vendedores = DatOS.OpenRecordset(S, dbOpenSnapshot)
     S = "select * from clientes where cia='" + CiA + "'"
     Set Cliente = DatOS.OpenRecordset(S, dbOpenSnapshot)
     'Articulos
     S = "select * from articulos  where cia='" + CiA + "'"
     Set articulos = DatOS.OpenRecordset(S, dbOpenSnapshot)
     If UsaLotes = 1 Then
          Label21.Visible = True
          Text(14).Visible = True
          NomLOte.Visible = True
          S = "select * from lotes  where cia='" + CiA + "'"
          Set Lotes = DatOS.OpenRecordset(S, dbOpenSnapshot)
     End If
     If PreCBoD = 0 Then
          Label13 = "Precio : "
     Else
          Label13 = "Unidades : "
     End If
     Show
     Refresh
     Command1_Click
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     Command3.Enabled = True
     Command4.Enabled = True
     Call CargaItem(Item)
End Sub
Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then
          Select Case KeyCode
          Case vbKeyDelete
               Command4_Click
          End Select
     End If
End Sub
Private Sub Text_Change(Index As Integer)
     Select Case Index
     Case 0 'Bodegas
          Text(4) = ""
          Text(14) = ""
          NomBod.Caption = ""
          ListView1.ListItems.Clear
          Call CalculaSaldo
          Bodegas.FindFirst "c_bodega='" + Text(0).Text + "'"
          If Not Bodegas.NoMatch Then
               NomBod.Caption = Bodegas!d_bodega
               If Text(1) <> "" And Not Cliente.EOF And NomBod <> "" Then Picture1.Enabled = True
          Else
               Picture1.Enabled = False
          End If
     Case 1 'Cliente
          Text(7) = ""
          Text(8) = ""
          Text(8).Tag = ""
          Text(9) = ""
          La_cedula = "**"
          Text1 = "**"
          
          S = "select * from clientes where cia='" + CiA + "'"
          S = S + " and codigo='" + Text(1) + "'"
          Set Cliente = DatOS.OpenRecordset(S)
          
          If Not Cliente.EOF Then
               Text(7) = Nulo(Cliente!Descuento)
               Text(8) = Cliente!Nombre
               Text(9) = Cliente!Plazo
               Text(8).Tag = Nulo(Cliente!Direccion)
               Text(3) = Cliente!agente
               La_cedula = Nulo(Cliente!Cedula)
               Text1 = Nulo(Cliente!LINEA1)
               Text(15) = Nulo(Cliente!Telefono)
               If Text(0) <> "" And Not Bodegas.EOF Then Picture1.Enabled = True
               If IsNull(Cliente!numprecio) Then
                    S = "update clientes set numprecio=1 where codigo='"
                    S = S + Cliente!Codigo + "'"
                    DatOS.Execute S, 128
                    Cliente.Requery
               End If
          Else
               Picture1.Enabled = False
          End If
     Case 3 'vendedor
          NomVend.Caption = ""
          Vendedores.FindFirst "codigo='" + Text(3) + "'"
          If Not Vendedores.NoMatch Then NomVend.Caption = Vendedores!Nombre
     Case 4 'Articulo
          DescArt.Caption = ""
          Label16.Caption = ""
          Label22.Caption = ""
          ECombo(5).Clear
          Text(6) = ""
          Check1.Value = 0
                              
          articulos.FindFirst "c_articulo='" + Text(4) + "'"
          If Not articulos.NoMatch Then
               DescArt.Caption = articulos!d_articulo
               Check1.Value = articulos!sino_impu 'El impuesto
               Text(6) = Nulo(Cliente!Descuento) 'El descuento
               Dim DolaR@
               If FacTDoL = 1 Then
                    DolaR = TCambio
               Else
                    DolaR = 1
               End If
               If PreCBoD = 1 Then
                    S = "select monto,unidades.descripcion,numuni "
                    S = S + "from precios,unidades "
                    S = S + "where unidades.cia='" + CiA
                    S = S + "' and precios.cia=unidades.cia "
                    S = S + "' and unidades.codart=precios.codart "
                    S = S + "and unidades.unidades=precios.numuni "
                    S = S + "and precios.codart='" + articulos!c_articulo
                    S = S + "' and precios.codbod='" + Text(0) + "'"
                    Set Temp = DatOS.OpenRecordset(S)
                    With Temp
                    Do Until .EOF
                         ECombo(5).AddItem Nulo(!Descripcion), !NumUni, (!Monto * DolaR)
                         .MoveNext
                         w% = DoEvents
                    Loop
                    End With
                    ECombo(5).ListIndex = IIf(ECombo(5).ListCount > 0, 0, -1)
               Else
                    Dim Ind%
                    Ind = -1
                    S = "select * from listaprecios where cia='" + CiA
                    S = S + "' and codart='" + articulos!c_articulo
                    S = S + "' order by codigo"
                    Set Temp = DatOS.OpenRecordset(S)
                    With Temp
                    'Chequea Que el Articulo Exista en la Tabla de Existencias
                     S = "select count(*)  from existencias "
                     S = S + "where cia='" + CiA
                     S = S + "' and C_BODEGA= '" + Text(0) + "'"
                     S = S + " and c_articulo='" + Text(4) + "'"
                     Set Temp = DatOS.OpenRecordset(S)
                     If Temp(0) = 0 Then
                         MsgBox "El Articulo no Tiene Historial de Existencias"
                         Text(4).SetFocus
                         Text(4) = ""
                     End If
                    Do Until .EOF
                         ECombo(5).AddItem !Descripcion, 1, (!Monto * DolaR)
                         If Nulo(Cliente!numprecio) = Trim(!Codigo) Then
                              Ind = ECombo(5).ListCount - 1
                         End If
                         .MoveNext
                         w% = DoEvents
                    Loop
                    End With
                    If Ind = -1 Then
                         S = "No se encontró el precio asociado al cliente!"
                         Call Barra(S, 4)
                    End If
                    ECombo(5).ListIndex = IIf(Ind > -1, Ind, IIf(ECombo(5).ListCount > 0, 0, -1))
               End If
          End If
     Case 7, 10
          If ListView1.ListItems.Count > 0 Then Call CalculaSaldo
     Case 12
          If Text(12) <> "" Then Call CargaProforma(Text(12))
     Case 14 'Lotes
          NomLOte.Caption = ""
          Lotes.FindFirst "codigo='" + Text(14) + "'"
          If Not Lotes.NoMatch Then NomLOte.Caption = Lotes!Nombre
     End Select
     If Index = 4 Or Index = 5 Or Index = 14 Then Call Valida
End Sub
Private Function Valida() As Boolean
     Command2.Enabled = False
     'Command3.Enabled = False
     If Text(4) = "" Then Exit Function
     If Label22 = "" Then Exit Function
     'If UsaLotes = 1 And NomLote = "" Then Exit Function
     If Text(5) = "" Then Exit Function
     Command2.Enabled = True
     'Command3.Enabled = True
     Valida = True
End Function
Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
     Select Case Index
     Case 0, 1, 3
          Call Barra("F3 para buscar por lista")
     Case 4, 14
          Call Barra("F3 para buscar por lista - F7 existencias")
     Case 12
          Call Barra("F3 para lista de proformas")
     End Select
End Sub
Private Sub CalculaSaldo()
     If ListView1.ListItems.Count > 0 Then
          Dim IvI#
          Dim SubTot#
          Dim Costo#
          Dim Tot#
          Dim DesC#
          Dim Item As ListItem
          For I = 1 To ListView1.ListItems.Count
               Set Item = ListView1.ListItems(I)
               SubTot = SubTot + Doble(Item.SubItems(5))
               DesC = DesC + Doble(Item.SubItems(7))
               IvI = IvI + Doble(Item.SubItems(9))
               Tot = Tot + Doble(Item.SubItems(10))
               Costo = Costo + (Doble(Item.Tag) * Doble(Item.SubItems(4)))
          Next
          Subtotal.Tag = Costo
          Subtotal.Caption = Format(SubTot, "standard")
          'Le resta el descuento
          Tot = SubTot - DesC
          'Le suma el impuesto
          Tot = Tot + IvI
          'Le suma el flete
          Tot = Tot + Doble(Text(10).Text)
          Descuento.Caption = Format(DesC, "standard")
          Impuesto.Caption = Format(IvI, "standard")
          'Tot = redo_05(Tot)
          Total.Caption = Format(Tot, "standard")
          Command5.Enabled = True
     Else
          Subtotal.Caption = ""
          Descuento.Caption = ""
          Impuesto.Caption = ""
          Total.Caption = ""
          Command5.Enabled = False
     End If
     Frame1.Enabled = Command5.Enabled
End Sub
Private Function Inserta() As Boolean
     On Error GoTo INSErr
     Static Consecutivo%
     Call Barra("Generando Factura ...")
     'Inserta la factura
     DatOS.Execute Arma, 129
     'Inserta el desglose de la factura y actualiza existencias y estadisticas
     Call Barra("Actualizando desglose de la factura ...")
     For I = 1 To ListView1.ListItems.Count
          Set Item = ListView1.ListItems(I)
          Set Item2 = ListView2.ListItems(Item.Index)
          'Desglose
          S = "insert into desgfact(c_articulo,n_factura,cantidad,"
          S = S + "unidades,precio,total_brut,descuento,porc_imp,"
          S = S + "total_neto,c_bodega,costo,lote,cia,serie) values ('"
          S = S + Item.Text + "','" 'El codigo de articulo
          S = S + Label7.Caption + "'," 'El numero de factura
          S = S + Item.SubItems(4) + "," 'La cantidad vendida
          S = S + "1," 'El numero de unidades
          S = S + Format(Doble(Item.SubItems(3))) + "," 'El precio de cada unidad
          S = S + Format(Doble(Item.SubItems(5))) + "," 'El total sin IV y sin descuento
          S = S + Format(Doble(Item.SubItems(6))) + "," 'El % de descuento
          S = S + Format(Doble(Item.SubItems(8))) + "," 'El % de impuesto
          S = S + Format(Doble(Item.SubItems(10))) + ",'" 'El total con IV y con descuento
          S = S + Text(0).Text + "'," 'LA BODEGA
          S = S + Trim(Item.Tag) + ",'" 'El costo
          S = S + Item2.Tag + "','" + CiA + "','')" 'El Lote
          DatOS.Execute S, 128
          'Actualiza existencias
          S = "update existencias set existencia=existencia-" & Doble(Item.SubItems(4))
          S = S + ", f_ult_sal=#" + Format(Text(2), "m/d/yyyy")
          S = S + "#,doc_ult_sa='" + Label7.Caption
          S = S + "',tiposalida=0 " + "where cia='" + CiA + "' and c_bodega='" + Text(0)
          S = S + "' and c_articulo='" + Item.Text + "'"
          If UsaLotes = 1 Then S = S + " and lote='" + Item2.Tag + "'"
          DatOS.Execute S, 128
          'Estadisticas
          S = "update estadisticas "
          S = S + "set monto=monto+" + Format(Doble(Item.SubItems(10)))
          S = S + ",unidades=unidades+" + Item.SubItems(4)
          S = S + " where cia='" + CiA + "' and  codart='" + Item.Text
          S = S + "' and cliente='" + Text(1)
          S = S + "' and codbod='" + Text(0)
          S = S + "' and fecha='" + Format(Date, "yyyymm") + "'"
          DatOS.Execute S, 128
          If DatOS.RecordsAffected = 0 Then
               S = "insert into estadisticas (codart,fecha,"
               S = S + "monto,unidades,codbod,cliente,cia)"
               S = S + " values ('" + Item.Text + "','" 'El articulo
               S = S + Format(Date, "yyyymm") + "'," 'La fecha
               S = S + Format(Doble(Item.SubItems(10))) + "," 'El monto
               S = S + Item.SubItems(4) + ",'" 'La cantidad
               S = S + Text(0) + "','" 'La bodega
               S = S + Text(1) + "','" + CiA + "')" 'El cliente
               DatOS.Execute S, 128
          End If
          
     Next
'     If ECombo(2).Indice(2) = 1 And Text(1) <> "" Then  pasa todas 30/3/00
          Call GeneraCXC(Label7.Caption, Text(2), Text(1), _
          Total.Caption, 1, Val(Text(9)), Label7)
'     End If
     Inserta = True
     Call GBitacora(1, "Factura No. " + Label7.Caption)
     On Error GoTo 0

INSErr:
     If err.Number = 3022 Then
          Label7.Caption = Format(Val(Label7) + 1, "00000000")
          Resume
     ElseIf err.Number = 3315 Then
          MsgBox "Al menos uno de los campos vacíos es requerido !", 16, "Error"
     ElseIf err.Number = 3201 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Inserta"
     End If
     Call Barra("")
     On Error GoTo 0
End Function
Private Sub Limpia()
     For I = 0 To Text.Count - 1
          Text(I) = ""
     Next
     ECombo(2).ListIndex = -1
     ECombo(5).ListIndex = -1
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     
     Total.Caption = ""
     Subtotal.Caption = ""
     Impuesto.Caption = ""
     Descuento.Caption = ""
     Label7.Caption = ""
     Label16.Caption = ""
     pedido.Text = ""
     Check1.Value = 0
     Command5.Caption = "&Guardar"
     Command5.Enabled = False
     Command5.Picture = ImageList1.ListImages(2).Picture
     
     Frame1.Enabled = False
End Sub
Private Function AgregaItem(CodArt$, Articulo$, Cant@, Unidades#, _
     PorcDesc#, PorcImp%, PreciO#, Novo%, Minimo&, Costo#, _
     Lote$, Calculado%) As Boolean
     On Error GoTo Errores
     Dim IvI@
     Dim Descuento@
     Dim Subtotal@
     Dim Total@
     Dim Impuesto@
     Dim Avail@
     Dim ExisTot@
     Dim Llave$
     Dim Item2 As ListItem
     If Check2.Value = 0 Then IvI = IIf(PorcImp = 1, IV, 0)
     If Cant = 0 Then err.Raise 9996
     If Costo >= PreciO Or PreciO = 0 Then err.Raise 9994
     Subtotal = Cant * PreciO
     Subtotal = Subtotal - (Subtotal * (PorcDesc / 100))
     Subtotal = Subtotal + (Subtotal * (IvI / 100))
     'Saldo del cliente
     Avail = ValidaSaldo(Subtotal)
     If Avail <= 0 Then err.Raise 9995
     Subtotal = 0
     'Existencias
     If Calculado = 0 Then
          Avail = ValidaExis(CodArt, Text(0), Lote)
          If Avail <= 0 Then err.Raise 9999
          ExisTot = ValidaMinimo
          If (ExisTot - Cant) < Minimo Then err.Raise 9998
     End If
     'Minimo del articulo
     Llave = "*" + CodArt
     If UsaLotes = 1 Then Llave = Llave + Lote
     If Novo = 1 Then
          Set Item = ListView1.ListItems.Add(, Llave)
          Set Item2 = ListView2.ListItems.Add
     ElseIf Novo = 2 Then
          Set Item = ListView1.SelectedItem
          Item.Key = Llave
          Set Item2 = ListView2.ListItems(Item.Index)
     End If
     'La Lista Visible
     Item.Tag = Costo
     Item.Text = CodArt 'La descripcion
     Item.SubItems(1) = Articulo 'La descripcion
     Item.SubItems(2) = Unidades 'Las unidades
     Item.SubItems(3) = Format(PreciO, "standard") 'El costo
     Item.SubItems(4) = Cant 'La cantidad
     Subtotal = Cant * PreciO 'Calcula el subtotal
     Item.SubItems(5) = Format(Subtotal, "standard") 'El subtotal
     Descuento = Subtotal * (PorcDesc / 100) 'Calcula el descuento
     Item.SubItems(6) = PorcDesc 'El % de descuento
     Item.SubItems(7) = Format(Descuento, "standard") 'El descuento
     Impuesto = (Subtotal - Descuento) * IvI 'Calcula el impuesto
     Item.SubItems(8) = Format(IvI * 100, "standard") 'El % de impuesto
     Item.SubItems(9) = Format(Impuesto, "standard") 'El impuesto
     Total = (Subtotal - Descuento) + Impuesto 'Calcula el total
     'Total = redo_05(Str(Total))
     Item.SubItems(10) = Format(Total, "standard") 'El total
     'La Lista Invisible
     Item2.Tag = Lote 'El Lote
     Item2.Text = ECombo(5).ListIndex 'El numero de precio
     Set Item = Nothing
     AgregaItem = True
     Call CalculaSaldo
     Command5.Enabled = True
Errores:
     If err.Number = 35602 Then
          MsgBox "Ya incluyó este artículo !", 48, CodArt
          Text(4).SetFocus
     ElseIf err.Number = 9999 Then
          If Seguridad("factura1") Then
               S = "La cantidad a facturar  es mayor a la existencia"
               S = S + Chr(13) + "   en la bodega seleccionada, desea continuar ?"
               If MsgBox(S, 36, "Existencia: " + FormatNumber(Avail)) = 6 Then
                    Resume Next
               Else
                    Selecciona Text(5)
               End If
          Else
               S = "La cantidad a facturar es mayor a la existencia"
               S = S + Chr(13) + "          en la bodega seleccionada!"
               MsgBox S, 64, "Existencia: " + FormatNumber(Avail)
               Selecciona Text(5)
          End If
     ElseIf err.Number = 9996 Then
          MsgBox "La cantidad debe ser mayor a cero !", 16, "Error"
          Text(5).SetFocus
     ElseIf err.Number = 9998 Then
          If Seguridad("factura2") Then
               S = "La cantidad solicitada menos la existencia es menor"
               S = S + Chr(13) + "     al mínimo del articulo, desea continuar ?"
               If MsgBox(S, 36, "Existencia: " + FormatNumber(Avail) + " Mínimo: " + FormatNumber(articulos!Minimo)) = 6 Then
                    Resume Next
               Else
                    Exit Function
               End If
          Else
               S = "La cantidad solicitada menos la existencia es menor al mínimo del articulo !"
               MsgBox S, 64, "Existencia: " + FormatNumber(Avail) + " Mínimo: " + FormatNumber(articulos!Minimo)
               Selecciona Text(5)
          End If
     ElseIf err.Number = 9995 Then
          If Seguridad("factura3") Then
               S = "El monto a facturar sobrepasa el disponible de crédito del cliente,"
               S = S + Chr(13) + "                         Desea continuar ?"
               If MsgBox(S, 36, "Disponible: " + FormatNumber(Avail)) = 6 Then
                    Resume Next
               Else
                    Exit Function
               End If
          Else
               S = "El monto a facturar sobrepasa el límite de crédito del cliente !"
               MsgBox S, 64, "Disponible: " + FormatNumber(Avail)
               Selecciona Text(5)
          End If
     ElseIf err.Number = 9994 Then
          S = "El costo del artículo seleccionado es mayor o igual "
          S = S + Chr(13) + "   al precio de venta, no es posible facturarlo."
          MsgBox S, 16, "Costo: " + FormatNumber(Costo)
          Text(4).SetFocus
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Agrega Articulo"
     End If
End Function
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case Index
     Case 5, 6, 7, 10
          Select Case KeyAscii
          Case 8, 46, 48 To 57
          Case Else
               KeyAscii = 0
          End Select
     End Select
End Sub
Private Sub CargaCombo(Lista As Object, SQL As String)
     On Error GoTo Errores
     Set Temp = DatOS.OpenRecordset(SQL)
     Dim I%
     Indice = -1
     With Temp
     Do Until .EOF
          Lista.AddItem Temp(0)
          For I = 1 To .Fields.Count - 1
               Lista.List(Lista.ListCount, I) = Nulo(Temp(I))
               If LCase(Temp(I).Name) = "default" Then
                    If Temp(I) = 1 Then Indice = Lista.ListCount - 1
               End If
          Next
          .MoveNext
          w% = DoEvents
     Loop
     End With
     Lista.Tag = Indice
Errores:
     If err.Number = 3044 Then
          Exit Sub
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "CargaCombo"
     End If
     On Error GoTo 0
End Sub
Private Sub CargaItem(Item As ListItem)
     ECombo(5).ListIndex = Val(ListView2.ListItems(Item.Index).Text)
     Text(4) = Item.Text
     Text(5) = Item.SubItems(4)
     Text(6) = Item.SubItems(6)
     Text(14) = ListView2.ListItems(Item.Index).Tag
     Check1.Value = IIf(Val(Item.SubItems(8)) = 0, 0, 1)
End Sub
Private Function ValidaExis(Codigo$, Bodega$, Lote$) As Double
     S = "select existencia from existencias where cia='" + CiA + "' and c_articulo='"
     S = S + Codigo + "' and c_bodega='" + Bodega + "'"
     If UsaLotes = 1 Then S = S + " and lote='" + Lote + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Not Temp.EOF Then ValidaExis = Temp(0)
     Set Temp = Nothing
End Function
Private Function ValidaMinimo() As Double
     S = "select sum(existencia) from existencias "
     S = S + "where cia='" + CiA + "' and c_articulo='" + Text(4) + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Not IsNull(Temp(0)) Then
          ValidaMinimo = Temp(0)
     End If
     Set Temp = Nothing
End Function
Private Sub Text_LostFocus(Index As Integer)
     Call Barra("")
     If Index = 1 Then
          If Nulo(Cliente!credito) = 1 Then
               S = "El cliente seleccionado tiene el crédito bloqueado !"
               MsgBox S, vbCritical, Caption
               'Text(1).SetFocus
          End If
     End If
End Sub
Private Sub ActualizaConsec()
     Dim num$
     num = Format(Val(Label7.Caption) + 1, "00000000")
     S = "update parametros set consec='" + num + "' where cia='" + CiA + "'"
     DatOS.Execute S, 128
     If DatOS.RecordsAffected = 0 Then
          S = "insert into parametros(cia,consec) values('" + CiA + "','" + num + "')"
          DatOS.Execute S, 128
     End If
End Sub
Private Function Arma() As String
     S = "insert into facturas (n_factura,c_bodega,impuesto,"
     S = S + "monto,monto_real,f_factura,plazo,c_cliente,descuento,"
     S = S + "c_tipo_fac,vendedor,estado,flete,costo,nota,"
     S = S + "dolares,tipocambio,cia,autoriza) "
     S = S + "values('" + Label7.Caption + "','" 'El numero de factura
     S = S + Text(0) + "'," 'El codigo de bodega
     S = S & Doble(Impuesto) & "," 'El monto del impuesto
     S = S & Doble(Subtotal) & "," 'El subtotal
     S = S & Doble(Total) & ",#" 'El total de la factura
     S = S + Format(Text(2), "m/d/yyyy") + "#," 'La fecha
     S = S & Val(Text(9)) & ",'" 'El plazo
     S = S + Text(1) + "'," 'El cliente
     S = S & Doble(Descuento) & ",'" 'El descuento
     S = S + ECombo(2).Indice(1) + "','" 'El tipo de pago
     S = S + Text(3) + "'," 'El vendedor
     S = S + "0," 'El estado 0= No nula 1=Nula
     S = S & Doble(Text(10)) & "," 'El flete
     S = S & Subtotal.Tag & ",'" 'El costo
     S = S & Text(11) + "',0," 'La nota
     S = S & TCambio & ",'" + CiA + "','" + LoGiN + "')"  'Autoriza
     Arma = S
     Debug.Print S
End Function
Public Sub Inicial(Index As Integer)
     If Index = 2 Then
          If PreCBoD = 0 Then
               Label14.Caption = "Monto: "
               Label13.Caption = "Precio: "
          End If
          Show
          Refresh
     Else
          Unload Me
          Proformas.Show
     End If
End Sub
Private Sub CargaProforma(Numero$)
     Dim Temp As Recordset
     S = "select * from proformas where cia='" + CiA + "' and n_factura='" + Numero + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Not Temp.EOF Then
          With Temp
          Text(3).Text = Nulo(!Vendedor)
          Text(1).Text = Nulo(!c_cliente)
          Text(8).Text = Nulo(!Nombre)
          Text(0).Text = Nulo(!c_bodega)
          Text(11).Text = Nulo(!observaciones)
          Text(9).Text = Nulo(!Plazo)
          End With
          S = "select desgprof.c_articulo,d_articulo,costo,precio,"
          S = S + "cantidad,descuento,porc_imp "
          S = S + "from desgprof left join articulos "
          S = S + "on articulos.cia=desgprof.cia "
          S = S + "and articulos.c_articulo=desgprof.c_articulo "
          S = S + "where desgprof.cia='" + CiA + "' and n_factura='" + Numero + "'"
          Set Temp = DatOS.OpenRecordset(S)
          With Temp
          Do Until .EOF
               Set Item = ListView2.ListItems.Add
               Item.Text = Nulo(Cliente!numprecio)
               Set Item = ListView1.ListItems.Add(, "*" + Nulo(!c_articulo))
               Item.Tag = Nulo(!Costo)
               Item.Text = Nulo(!c_articulo) 'La descripcion
               Item.SubItems(1) = Nulo(!d_articulo) 'La descripcion
               Item.SubItems(2) = 1 'Las unidades
               Item.SubItems(3) = Format(!PreciO, "standard") 'El costo
               Item.SubItems(4) = !Cantidad 'La cantidad
               Subtotal = !Cantidad * !PreciO 'Calcula el subtotal
               Item.SubItems(5) = Format(Subtotal, "standard") 'El subtotal
               Descuento = Subtotal * (!Descuento / 100) 'Calcula el descuento
               Item.SubItems(6) = !Descuento 'El % de descuento
               Item.SubItems(7) = Format(Descuento, "standard") 'El descuento
               Impuesto = (Subtotal - Descuento) * (!PORC_IMP / 100) 'Calcula el impuesto
               Item.SubItems(8) = !PORC_IMP 'El % de impuesto
               Item.SubItems(9) = Format(Impuesto, "standard") 'El impuesto
               Total = (Subtotal - Descuento) + Impuesto 'Calcula el total
               Item.SubItems(10) = Format(Total, "standard") 'El total
               .MoveNext
          Loop
          End With
          If ListView1.ListItems.Count > 0 Then
               Command5.Enabled = True
               UsoProforma = True
          End If
          Call CalculaSaldo
     End If
End Sub
Private Function ValidaSaldo(Monto@) As Currency
     Dim Temp As Recordset
     Dim Debitos@
     Dim Creditos@
     Dim Disponible@
     Disponible = Cliente!limi_cred - Cliente!saldo
     If Disponible < Monto Then Exit Function
     S = "select tip_doc,sum(saldo_doc) from diariocxc "
     S = S + "where cia='" + CiA + "' and codigo='" + Text(1)
     S = S + "' group by tip_doc"
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          Select Case Temp!tip_doc
          Case 0 'Adelantos
               Debitos = Debitos + Temp(1)
          Case 1 'Facturas
               Creditos = Creditos + Temp(1)
          Case 2 'Notas cred
               Debitos = Debitos + Temp(1)
          Case 3 'Notas deb
               Creditos = Creditos + Temp(1)
          Case 4 'Recibos
               Debitos = Debitos + Temp(1)
          End Select
          Temp.MoveNext
     Loop
     Set Temp = Nothing
     For I = 1 To ListView1.ListItems.Count
          Set Item = ListView1.ListItems(I)
          Disponible = Disponible - Doble(Item.SubItems(10))
     Next
     Disponible = Disponible + (Debitos - Creditos)
     Disponible = Disponible - (Monto + Doble(Total.Caption))
     ValidaSaldo = Disponible
End Function
Public Function redo_05(xx_monto As Double) As Double
     Dim Xx As Double
     Dim mon As Double
     Dim mon1 As Double
     mon = xx_monto
     Xx = xx_monto - Int(xx_monto)
     'xx = xx * 10
     'mon1 = xx - Int(xx)
     'mon1 = Round(mon1, 1)
     If Xx > 0 Then xx_monto = Int(xx_monto) + 1
     redo_05 = xx_monto
     'redo_05 = Round(xx_monto * 2, 1) / 2
End Function
Public Sub Imprimefactura(FactExp$, Bodega$, Reporte As CrystalReport, clien$)
     On Error GoTo Errores
     FactExp = Format(FactExp, "00000000")
     S = "{facturas.cia}='" + CiA + "' and {facturas.c_bodega}='" + Bodega
     S = S + "' and {facturas.n_factura}='" + FactExp + "'"
     Reporte.ReportFileName = DirTrA + "reportes\factexpor.rpt"
     Reporte.DataFiles(0) = DatOS.Name
     Reporte.Formulas(5) = "nomclie='" + clien + "'"
     'Reporte.Formulas(6) = "comodin3='Compañía : " + CiA + " " + Inicio.StatusBar1.Panels(2) + "'"
     'Reporte.WindowTitle = "Entrada No. " + Entrada
     Reporte.SelectionFormula = S
     Reporte.Action = 1
Errores:
     If err.Number = 20526 Then
          MsgBox "Debe existir al menos una impresora en el sistema !", 48, "Entradas"
     ElseIf err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "ImprimeFactura"
     End If
     On Error GoTo 0
End Sub

