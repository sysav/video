VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Entradas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Entradas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   10725
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   8100
      TabIndex        =   4
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "nueva"
            Object.ToolTipText     =   "Generar una entrada nueva"
            Object.Tag             =   "Nueva"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "inserta"
            Object.ToolTipText     =   "Agregar artículos"
            Object.Tag             =   "Insertar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modifica"
            Object.ToolTipText     =   "Modificar el artículo seleccionado"
            Object.Tag             =   "Modificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "borra"
            Object.ToolTipText     =   "Borrar el artículo seleccionado"
            Object.Tag             =   "Borrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "graba"
            Object.ToolTipText     =   "Guardar los cambios"
            Object.Tag             =   "Guardar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sale"
            Object.ToolTipText     =   "Cerrar esta ventana"
            Object.Tag             =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Entradas.frx":030A
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   150
      Top             =   5370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   2895
      Left            =   0
      TabIndex        =   30
      Top             =   2340
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lote"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P. Compra"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "N. Costo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Desc"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "IV"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Total"
         Object.Width           =   2117
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   345
      Left            =   270
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4770
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView ListView4 
      Height          =   405
      Left            =   1710
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3870
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   714
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   885
      Left            =   60
      ScaleHeight     =   825
      ScaleWidth      =   10575
      TabIndex        =   38
      Top             =   1380
      Width           =   10635
      Begin VB.TextBox Costos 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   1
         Left            =   3360
         TabIndex        =   26
         Top             =   30
         Width           =   1245
      End
      Begin VB.TextBox Costos 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   0
         Left            =   2370
         TabIndex        =   28
         Top             =   390
         Width           =   1815
      End
      Begin ComctlLib.ListView ListView3 
         Height          =   825
         Left            =   4680
         TabIndex        =   27
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1455
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Nombre"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "T. Cambio"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4230
         Picture         =   "Entradas.frx":0624
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Aplicar el monto al costo seleccionado"
         Top             =   420
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1950
         TabIndex        =   42
         Top             =   90
         Width           =   1365
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1740
         TabIndex        =   40
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Costeo de Importaciones"
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
         Height          =   465
         Left            =   60
         TabIndex        =   39
         Top             =   0
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   855
      Left            =   60
      ScaleHeight     =   795
      ScaleWidth      =   10575
      TabIndex        =   15
      Top             =   480
      Width           =   10635
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   8
         Left            =   8775
         MultiLine       =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Monto del flete"
         Top             =   390
         Width           =   1755
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   7
         Left            =   6930
         TabIndex        =   22
         ToolTipText     =   "Monto del flete"
         Top             =   390
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   0
         Left            =   4170
         MaxLength       =   10
         TabIndex        =   17
         ToolTipText     =   "Factura del proveedor"
         Top             =   30
         Width           =   1725
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   2
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   18
         ToolTipText     =   "Fecha del documento"
         Top             =   30
         Width           =   1185
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   3
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   21
         ToolTipText     =   "Código del proveedor"
         Top             =   390
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1680
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Número de orden de compra"
         Top             =   30
         Width           =   1635
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   5
         Left            =   8670
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Descuento global"
         Top             =   30
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   6
         Left            =   9900
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Plazo del proveedor"
         Top             =   30
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7830
         TabIndex        =   37
         Top             =   450
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flete :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6420
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factura :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3390
         TabIndex        =   35
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6000
         TabIndex        =   34
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   33
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   9360
         TabIndex        =   32
         Top             =   120
         Width           =   510
      End
      Begin VB.Label NomProv 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2310
         TabIndex        =   31
         ToolTipText     =   "Nombre del proveedor"
         Top             =   390
         Width           =   4035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de compra  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc (%) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7860
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   4
      Left            =   900
      MaxLength       =   2
      TabIndex        =   0
      ToolTipText     =   "Código de la bodega"
      Top             =   90
      Width           =   615
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   5970
      MaxLength       =   10
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Número del documento"
      Top             =   90
      Width           =   1245
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Descuento :"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   960
      TabIndex        =   14
      Top             =   5640
      Width           =   1680
   End
   Begin VB.Label Descuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2715
      TabIndex        =   13
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label NomBod 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1530
      TabIndex        =   12
      ToolTipText     =   "Nombre de la bodega"
      Top             =   90
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bodega :"
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6090
      TabIndex        =   10
      Top             =   5640
      Width           =   870
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7035
      TabIndex        =   9
      Top             =   5640
      Width           =   3675
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "I.V. :"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6330
      TabIndex        =   8
      Top             =   5250
      Width           =   630
   End
   Begin VB.Label Impuesto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7035
      TabIndex        =   7
      Top             =   5250
      Width           =   3675
   End
   Begin VB.Label SubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2715
      TabIndex        =   6
      Top             =   5250
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "SubTotal :"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1290
      TabIndex        =   5
      Top             =   5280
      Width           =   1350
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   570
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":0726
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":0838
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":094A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":0A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":0D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":1090
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":13AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":16C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Entradas.frx":17D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Número  :"
      Height          =   225
      Left            =   5130
      TabIndex        =   2
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "Entradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BackOrd As Boolean
Dim S$
Dim Compra%
Dim Ordenes As Recordset
Dim Proveedor As Recordset
Dim Bodega As Recordset
Dim Item As ListItem
Private Sub Combo1_Change()
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     Text(0) = ""
     Text(3) = ""
     Text(6) = ""
     Impuesto.Caption = ""
     Subtotal.Caption = ""
     Total.Caption = ""
     Ordenes.FindFirst "numero='" + Combo1.Text + "'"
     If Not Ordenes.NoMatch Then
          Dim Temp As Recordset
          Proveedor.FindFirst "codigo='" + Ordenes!Proveedor + "'"
          If Proveedor.NoMatch Then
               MsgBox "El proveedor no existe !", 48, Caption
               Text(3).SetFocus
          Else
               Text(3) = Ordenes!Proveedor
          End If
          S = "select codart,cantidad,precio,d_articulo,impuesto,"
          S = S + "descuento,p_compra,maximo,porc_util "
          S = S + "from desgord left join articulos "
          S = S + "on articulos.cia=desgord.cia "
          S = S + "and articulos.c_articulo=desgord.codart "
          S = S + "where desgord.cia='" + CiA + "' and orden='" + Ordenes!Numero + "'"
          Set Temp = DatOS.OpenRecordset(S)
          Dim Item2 As ListItem
          Dim Monto#
          Dim ImpU#
          Dim DesC#
          Do Until Temp.EOF
               If IsNull(Temp!d_articulo) Then
                    S = "El artículo '" + Temp!CodArt + "' no existe en le sistema !"
                    MsgBox S, 48, Caption
               Else
                    Set Item2 = ListView2.ListItems.Add(, "*" + Temp!CodArt, , , 5)
                    Set Item = ListView1.ListItems.Add
                    'La lista visible
                    Monto = Temp!Cantidad * Temp!PreciO
                    DesC = Monto * (Temp!Descuento / 100)
                    ImpU = (Temp!Impuesto / 100) * (Monto - DesC)
                    Item2.Tag = Temp!p_compra 'El precio de costo
                    Item2.Text = Temp!CodArt 'El codigo
                    Item2.SubItems(1) = Temp!d_articulo 'La descripcion
                    Item2.SubItems(3) = FormatNumber(Temp!p_compra, DeCiMaleS) 'El costo
                    Item2.SubItems(4) = FormatNumber(Temp!PreciO, DeCiMaleS) 'El costo
                    Item2.SubItems(5) = Temp!Cantidad 'La cantidad
                    Item2.SubItems(6) = Temp!Descuento 'La cantidad
                    Item2.SubItems(7) = Temp!Impuesto 'La cantidad
                    Item2.SubItems(8) = FormatNumber(Monto - DesC + ImpU) 'El total
                    'La lista invisible
                    Item.Text = ImpU 'El % de impuesto
                    Item.SubItems(1) = DesC 'El descuento
                    Item.SubItems(2) = Nulo(Temp!Maximo) 'El nuevo maximo
                    Item.SubItems(3) = Nulo(Temp!porc_util) 'El nuevo utilidad
               End If
               Temp.MoveNext
               w% = DoEvents
          Loop
          Call Barra("")
          MousePointer = 0
          Call CalculaSaldo
          If ListView2.ListItems.Count > 0 Then
               Toolbar1.Buttons("graba").Enabled = True
               BackOrd = True
          Else
               BackOrd = False
          End If
     Else
          BackOrd = False
     End If
End Sub
Private Sub Combo1_Click()
     If Combo1.Text <> "" Then Combo1_Change
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
     If Len(Combo1.Text) >= 8 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub Command1_Click()
     If ListView2.ListItems.Count > 0 Then
          S = "Todos los artículos incluidos se perderán, desea continuar ?"
          If MsgBox(S, 36, "Entradas") = 6 Then
               ListView2.ListItems.Clear
               ListView1.ListItems.Clear
               Call CalculaSaldo
          Else
               Exit Sub
          End If
     End If
     Set Item = ListView3.SelectedItem
     Item.SubItems(1) = FormatNumber(Doble(Costos(0)))
     Item.SubItems(2) = FormatNumber(Doble(Costos(1)))
     Costos(0).Text = ""
     ListView3.SetFocus
End Sub
Private Sub Costos_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case KeyAscii
     Case 8, 46, 48 To 57
     Case 13
          If Costos(0) <> "" And Costos(1) <> "" Then Command1_Click
     Case Else
          KeyAscii = 0
     End Select
End Sub
Private Sub Form_Activate()
     Call Menus(True, Me)
End Sub
Private Sub Form_Deactivate()
     Call Menus(False, Me)
End Sub
Private Sub Form_Load()
     Call Posiciona(Me, 0)
     S = "select codigo,nombre,plazo,moneda "
     S = S + "from proveedores where compania='" + CiA + "'"
     Set Proveedor = DatOS.OpenRecordset(S)
     S = "select c_bodega,d_bodega from bodegas where cia='" + CiA + "'"
     Set Bodega = DatOS.OpenRecordset(S)
     If CosTeO <> 1 Then
          ListView2.Top = Picture1.Top + Picture1.Height + 10
          ListView2.Height = ListView2.Height + Picture2.Height + 10
          Picture2.Enabled = False
     Else
          S = "select * from costos  where cia='" + CiA + "' order by nombre"
          Set Tabla = DatOS.OpenRecordset(S)
          Do Until Tabla.EOF
               Set Item = ListView3.ListItems.Add
               Item.Tag = Tabla!Codigo
               Item.Text = Tabla!Nombre
               Item.SubItems(1) = "0.00"
               Item.SubItems(2) = "1.00"
               Tabla.MoveNext
          Loop
          Costos(1) = TCambio
     End If
     If UsaLotes = 0 Then ListView2.ColumnHeaders(3).Width = 0
     Show
     Refresh
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     If ListView1.ListItems.Count > 0 Then
          S = "No se ha guardado la entrada en proceso, desea salir ?"
          If MsgBox(S, 292, Caption) <> 6 Then
               Cancel = True
          Else
               On Error Resume Next
               S = "delete from series where cia='" + CiA
               S = S + "' and entrada='" + Format(Text(1), "00000000")
               S = S + "' and estado=0"
               DatOS.Execute S, 128
               On Error GoTo 0
          End If
     End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Menus(False, Me)
     Call Posiciona(Me, 1)
     Call Barra("")
End Sub
Private Sub ListView2_DblClick()
     If ListView2.SelectedItem Is Nothing Then Exit Sub
     If ListView2.SelectedItem.Selected Then
          Call AddArt.Carga(Me, 1, ListView2.SelectedItem, Compra)
          AddArt.Show 1
     End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
     Toolbar1.Buttons("modifica").Enabled = True
     Inicio.SubPop(2).Enabled = True
     If Compra < 2 Then
          Toolbar1.Buttons("borra").Enabled = True
          Inicio.SubPop(3).Enabled = True
     End If
End Sub
Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
          If ListView2.SelectedItem Is Nothing Then Exit Sub
          If ListView2.SelectedItem.Selected Then
               Call Toolbar1_ButtonClick(Toolbar1.Buttons(5))
          End If
     End If
End Sub
Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim I%
     For I = 0 To Inicio.SubPop.Count - 1
          If Inicio.SubPop(I).Enabled And Inicio.SubPop(I).Caption <> "-" Then Exit For
     Next
     If Button = 2 Then PopupMenu Inicio.Popup, , , , Inicio.SubPop(I)
End Sub
Private Sub ListView3_ItemClick(ByVal Item As ComctlLib.ListItem)
     Costos(0).Text = Item.SubItems(1)
     Costos(0).SelStart = 0
     Costos(0).SelLength = Len(Costos(0))
End Sub
Private Sub Text_Change(Index As Integer)
     Select Case Index
     Case 1
          Picture1.Enabled = IIf(Text(1).Text = "", False, True)
          If Seguridad("entrada1") Then
               Picture2.Enabled = IIf(Text(1).Text = "", False, True)
          End If
     Case 3 'El proveedor
          NomProv.Caption = ""
          Toolbar1.Buttons("inserta").Enabled = False
          Inicio.SubPop(Toolbar1.Buttons("inserta").Index - 1).Enabled = False
          Text(6).Text = ""
          Proveedor.FindFirst "codigo='" + Format(Text(3), "0000000000") + "'"
          If Not Proveedor.NoMatch Then
               NomProv.Caption = Proveedor!Nombre
               Text(6).Text = Proveedor!Plazo
               If Compra < 2 Then
                    Toolbar1.Buttons("inserta").Enabled = True
                    Inicio.SubPop(2).Enabled = True
               End If
          End If
     Case 4 'La bodega
          NomBod = ""
          Text(0) = ""
          Text(3) = ""
          Text(5) = ""
          Text(6) = ""
          Combo1.Text = ""
          Toolbar1.Buttons(1).Enabled = False
          Inicio.SubPop(0).Enabled = False
          Picture1.Enabled = False
          Bodega.FindFirst "c_bodega='" + Text(4) + "'"
          If Not Bodega.NoMatch Then
               NomBod.Caption = Bodega!d_bodega
               If Compra < 2 Then
                    Toolbar1.Buttons(1).Enabled = True
                    Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                    Inicio.SubPop(0).Enabled = True
                    If Combo1.Enabled Then Combo1.SetFocus
               End If
          End If
     Case 5
          Call CalculaSaldo
     End Select
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Select Case Index
     Case 3, 4
          Call Barra("Presione F3 para buscar por lista")
     Case 1
          If Compra = 2 Then Call Barra("Presione F3 para buscar por lista")
     End Select
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case Index
     Case 2
          Select Case KeyAscii
          Case 8, 47 To 57
          Case Else
               KeyAscii = 0
          End Select
     Case 1
          Select Case KeyAscii
          Case 8, 48 To 57
          Case Else
               KeyAscii = 0
          End Select
     Case 5, 7
          Select Case KeyAscii
          Case 8, 46, 48 To 57
          Case Else
               KeyAscii = 0
          End Select
     End Select
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     Select Case Index
     Case 1
          If Compra = 2 And KeyCode = vbKeyF3 Then
               S = "select n_entrada as Número,Nombre as Proveedor,f_entrada as Fecha "
               S = S + "from entradas left join proveedores "
               S = S + "on proveedores.compania=entradas.cia "
               S = S + "and proveedores.codigo=entradas.c_prove "
               S = S + "where entradas.cia='" + CiA
               S = S + "' and entradas.tipo=0"
               If Text(3).Text <> "" Then S = S + " and c_prove='" + Text(3) + "'"
               S = S + " order by f_entrada,n_entrada "
               Call Lista.Carga(Text(Index), S, "Entradas por Compra")
               Lista.Show 1
          End If
     Case 3
          If KeyCode = vbKeyF3 Then
               S = "select codigo as Código,Nombre "
               S = S + "from proveedores where compania='" + CiA + "' order by nombre "
               Call Lista.Carga(Text(Index), S, "Proveedores")
               Lista.Show 1
               Text(Index).SetFocus
          End If
     Case 4
          If KeyCode = vbKeyF3 Then
               S = "select c_bodega as Código,d_bodega as Descripcion "
               S = S + "from bodegas where cia='" + CiA + "' order by d_bodega "
               Call Lista.Carga(Text(Index), S, "Bodegas")
               Lista.Show 1
               Text(Index).SetFocus
          End If
     End Select
End Sub
Private Sub Text_LostFocus(Index As Integer)
     If Index = 3 Or Index = 1 Then Text(Index) = Format(Text(Index), "0000000000")
     Call Barra("")
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
     Select Case Button.Key
     Case "nueva"
          Call Limpia(False)
          If Compra < 2 Then
               Text(1) = NewNum
               Text(2) = Format(Date, "dd/mm/yyyy")
          End If
     Case "inserta"
          Call AddArt.Carga(Me, 0)
          AddArt.Show 1
          If ListView1.ListItems.Count > 0 Then
               Toolbar1.Buttons("graba").Enabled = True
               Inicio.SubPop(4).Enabled = True
               Text(7).Enabled = False
          End If
     Case "modifica"
          ListView2_DblClick
     Case "borra"
          If MsgBox("Desea borrar la línea seleccionada ?", 36, "Entradas") = 6 Then
               Set Item = ListView2.SelectedItem
               S = "delete from series where cia='" + CiA
               S = S + "' and entrada='" + Format(Text(1), "00000000")
               S = S + "' and codigo='" + Item.Text + "' and estado=0"
               DatOS.Execute S, 128
               ListView1.ListItems.Remove Item.Index
               ListView2.ListItems.Remove Item.Index
               Call CalculaSaldo
               If ListView1.ListItems.Count = 0 Then
                    Toolbar1.Buttons("modifica").Enabled = False
                    Toolbar1.Buttons("borra").Enabled = False
                    Toolbar1.Buttons("graba").Enabled = False
                    Inicio.SubPop(6).Enabled = False
                    If Compra = 0 Then Text(7).Enabled = True
               Else
                    ListView2.SelectedItem.Selected = True
               End If
               ListView2.SetFocus
          End If
     Case "graba"
          MousePointer = 11
          w% = DoEvents
          Select Case Compra
          Case 0, 1
               If ListView1.ListItems.Count > 0 Then
                    EspaCio.BeginTrans
                    If Inserta Then
                         EspaCio.CommitTrans
                         MsgBox "Entrada generada en Inventario y Facturación", 64, "Entradas"
                         If Compra = 0 Then
                              If BackOrd Then
                                   EspaCio.BeginTrans
                                   If BackOrder Then
                                        EspaCio.CommitTrans
                                   Else
                                        EspaCio.Rollback
                                   End If
                                   S = "delete from ordenes where cia='" + CiA
                                   S = S + "' and numero='" + Combo1.Text + "'"
                                   DatOS.Execute S, 128
                                   S = "delete from desgord where cia='" + CiA
                                   S = S + "' and orden='" + Combo1.Text + "'"
                                   DatOS.Execute S, 128
                              End If
                              If Val(Text(6)) > 0 Then
                                   'EspaCio.BeginTrans
                                   'If GeneraCXP Then
                                   '     EspaCio.CommitTrans
                                   '     MsgBox "Entrada generada en Cuentas por Pagar", 64, "Entradas"
                                   'Else
                                   '     EspaCio.Rollback
                                   'End If
                              End If
                         End If
                         Call ImprimeEntrada(Text(1), Text(4), Report1, NomProv)
                         If App.Comments = "Dafesa" Or App.Comments = "Sejim" Then
                              S = "Desea imprimir las etiquetas de la entrada ?"
                              If MsgBox(S, 36, Caption) = 6 Then
                                   Call Barras.DesdeEntradas(Text(1), Text(4))
                              End If
                         End If
                         Call Barra("")
                         Call Limpia(False)
                    Else
                         EspaCio.Rollback
                    End If
               End If
          End Select
          MousePointer = 0
     Case "sale"
          Unload Me
     End Select
End Sub
Public Sub CalculaSaldo()
     If ListView1.ListItems.Count > 0 Then
          Dim SubTot@
          Dim IvE@
          Dim Gravado@
          Dim Exento@
          Dim DesC@
          Dim DescLinea@
          Dim Temp@
          Dim Incremento@
          Dim Item As ListItem
          Dim Cuantos%
          Dim Cantidad@
          Dim Flete@
          Dim Costo@
          Dim NuevoCosto@
          Dim Linea@
          'Calcula el costo total con el precio nuevo
          'de los articulos incluidos
          Cuantos = ListView2.ListItems.Count
          For I = 1 To Cuantos
               Set Item = ListView2.ListItems(I)
               Costo = Doble(Item.SubItems(3))
               Cantidad = Item.SubItems(5)
               SubTot = SubTot + (Costo * Cantidad)
          Next
          'Si el flete es mayor a cero
          Flete = Doble(Text(7))
          If Flete > 0 Then Incremento = Flete / SubTot
          'Si la lista de costos directos es mayor a cero
          If ListView3.ListItems.Count > 0 Then
               Dim Costos@
               Dim Conv@
               For I = 1 To ListView3.ListItems.Count
                    Set Item = ListView3.ListItems(I)
                    'El monto por el tipo de cambio
                    Conv = Doble(Item.SubItems(1)) * Doble(Item.SubItems(2))
                    Costos = Costos + Conv
               Next
               'El incremento seria el flete mas el total
               'de los costos directos entre el total bruto de la entrada
               Incremento = ((Flete + Costos) / SubTot)
          End If
          SubTot = 0
          Dim Res@ 'El Nuevo costo
          For I = 1 To Cuantos
               Set Item = ListView2.ListItems(I)
               Costo = Doble(Item.SubItems(3))
               NuevoCosto = Item.Tag 'Item.SubItems(4)
               Cantidad = Item.SubItems(5)
               'Incrementa el nuevo costo segun los porcentajes y/o el flete
               Res = NuevoCosto + (NuevoCosto * Incremento)
               Item.SubItems(4) = FormatNumber(Res, DeCiMaleS)
               'Calcula el subtotal por linea con el precio del proveedor
               Temp = Costo * Cantidad
               Item.SubItems(8) = FormatNumber(Temp) 'Total
               'Calcula el descuento por linea
               DescLinea = Temp * (Item.SubItems(6) / 100)
               'El impuesto por linea
               IvE = IvE + ((Temp - DescLinea) * (Item.SubItems(7) / 100))
               SubTot = SubTot + Temp
               DesC = DesC + DescLinea
          Next
          Descuento.Caption = Format(DesC, "standard")
          Subtotal.Caption = Format(SubTot, "standard")
          Impuesto.Caption = Format(IvE, "standard")
          Total.Caption = Format(SubTot - DesC + IvE + Flete, "standard")
     Else
          Descuento.Caption = ""
          Subtotal.Caption = ""
          Impuesto.Caption = ""
          Total.Caption = ""
     End If
End Sub
Private Function Inserta() As Boolean
     On Error GoTo Errores
     Dim Temp As Recordset
     Dim NumEnt$
     Dim Confirma As Boolean
     Dim ActualizaPrecios As Boolean
     NumEnt = Format(Text(1), "00000000")
     Call Barra("Generando Entrada ...")
     'Inserta la entrada
     'If Text(0) = "" Then err.Raise 9999
     If NomProv = "" Then err.Raise 9997
     S = "select 1 from entradas where cia='" + CiA + "' and c_prove='" + Text(3)
     S = S + "' and fact_comp='" + Text(0) + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Not Temp.EOF Then err.Raise 9998
     Set Temp = Nothing
     DatOS.Execute Arma, 128
     'Inserta el desglose de la entrada
     Dim Costo@
     Dim CostoExt@
     Dim CostoNew@
     Dim Mdesc@
     Dim MImp@
     If App.Comments = "Dafesa" Then
          S = "Desea actualizar los precios de venta en base al nuevo costo ?"
          If MsgBox(S, 36, Caption) = 6 Then
               ActualizaPrecios = True
          End If
     End If
     For I = 1 To ListView2.ListItems.Count
          Set Item = ListView2.ListItems(I)
          If UsaLotes = 1 And Item.SubItems(2) = "" Then
               S = "Al menos uno de los artículos por incluir tiene un lote inválido !"
               MsgBox S, 16, "Error"
               Item.EnsureVisible
               Item.Selected = True
               Exit Function
          End If
          Costo = Doble(Item.SubItems(3))
          CostoNew = Doble(Item.SubItems(4))
          CostoExt = Doble(ListView1.ListItems(Item.Index).Tag)
          Mdesc = (Costo * Doble(Item.SubItems(5)) * (Item.SubItems(6) / 100))
          MImp = ((Costo * Doble(Item.SubItems(5)) - Mdesc) * (Item.SubItems(7) / 100))
          
          'Inserta el desglose de la entrada
          S = "insert into desgent (c_articulo,n_entrada,cantidad,"
          S = S + "p_compra,total_brut,total_neto,porc_imp,"
          S = S + "bodega,descuento,lote,cia) values ('"
          S = S + Item.Text + "','" 'El articulo
          S = S + NumEnt + "'," 'La entrada
          S = S & Doble(Item.SubItems(5)) & "," 'La cantidad
          S = S & Costo & "," 'El costo
          S = S & (Costo * Doble(Item.SubItems(5))) & "," 'El total sin IV
          S = S & (Doble(Item.SubItems(8)) - Mdesc + MImp) & "," 'El total con IV
          S = S & Doble(Item.SubItems(7)) 'El % de impuesto
          S = S + ",'" + Text(4) + "'," 'La bodega
          S = S & Doble(Item.SubItems(6)) & ",'" 'El descuento
          S = S + Item.SubItems(2) + "','" + CiA + "')" 'El Lote
          DatOS.Execute S, 128
          'Actualiza el precio de compra del articulo
          S = "update articulos set costoant=" & Costo
          S = S + ",p_compra=" & CostoNew
          S = S + ",costoex=" & CostoExt
          S = S + " where cia='" + CiA + "' and c_articulo='" + Item.Text + "'"
          DatOS.Execute S, 128
          'Precios de Venta
          If ActualizaPrecios And CostoNew <> CostoAnt Then
               If PreciosVenta(Item.Text, CostoNew, Item.SubItems(1), Confirma) Then
                    If Not Confirma Then
                         S = "Desea seguir confirmando cada artículo ?"
                         If MsgBox(S, 36, Caption) <> 6 Then Confirma = True
                    End If
               Else
                    Exit Function
               End If
          End If
          'Actualiza las existencias
          S = "update existencias set existencia=existencia+"
          S = S & Doble(Item.SubItems(5))  'La cantidad
          S = S + ",f_ult_ent=#" + Format(Text(2), "m/d/yyyy")
          S = S + "#, doc_ult_en='" + NumEnt
          S = S + "',tipoentrada=" & Compra
          S = S + ",costo=" & Costo
          S = S + ",costopro=" & CostoNew
          S = S + " where cia='" + CiA + "' and c_bodega='" + Text(4)
          S = S + "' and c_articulo='" + Item.Text + "'"
          If UsaLotes = 1 Then S = S + " and lote='" + Item.SubItems(2) + "'"
          DatOS.Execute S, 128
          If DatOS.RecordsAffected = 0 Then
               S = "insert into existencias(c_bodega,c_articulo,existencia,"
               S = S + "f_ult_ent,doc_ult_en,tipoentrada,lote,costo,costopro,cia)"
               S = S + "values('" + Text(4) + "','" 'La bodega
               S = S + Item.Text + "'," 'El articulo
               S = S & Doble(Item.SubItems(5)) & ",#" 'La cantidad
               S = S + Format(Text(2), "m/d/yyyy") + "#,'" 'La fecha
               S = S + NumEnt + "'," 'La ultima entrada
               S = S & Compra   'El tipo de entrada
               S = S + ",'" + Item.SubItems(2) 'El lote
               S = S + "'," & Costo 'El costo del proveedor
               S = S + "," & CostoNew & ",'" + CiA + "')" 'Costo promedio
               DatOS.Execute S, 128
          End If
          'Actualiza la estadistica de compras si es una entrada por compra
     Next
     'Inserta los costos de importacion
     Dim CT#
     For I = 1 To ListView3.ListItems.Count
          Set Item = ListView3.ListItems(I)
          CT = Doble(Item.SubItems(1))
          If CT > 0 Then
               S = "insert into costosbodega(bodega,codigo,entrada,"
               S = S + "tcambio,monto,cia) "
               S = S + "values('" + Text(4) + "','" + Item.Tag
               S = S + "','" + NumEnt
               S = S + "'," & Doble(Item.SubItems(2))
               S = S + "," & Doble(Item.SubItems(1)) & ",'" + CiA + "')"
               DatOS.Execute S, 128
          End If
     Next
     'Inserta los costos directos por articulo
     For I = 1 To ListView4.ListItems.Count
          Set Item = ListView4.ListItems(I)
          CT = Doble(Item.SubItems(1))
          If CT > 0 Then
               S = "insert into costosarticulo(bodega,codigo,entrada,"
               S = S + "tcambio,monto,articulo,cia) values('"
               S = S + Text(4).Text + "','" + Item.Tag
               S = S + "','" + NumEnt + "'," + Trim(Doble(Item.SubItems(1)))
               S = S + "," + Trim(Doble(Item.Text))
               S = S + ",'" + Item.SubItems(2) + "','" + CiA + "')"
               DatOS.Execute S, 128
          End If
     Next
     w% = DoEvents
     On Error GoTo 0
     Inserta = True
     Call GBitacora(1, "Entrada No. " + NumEnt)
Errores:
     If err.Number = 3022 Then
          MsgBox "Entrada ya existente en la bodega seleccionada !", 16, "Error"
          Selecciona Text(1)
     ElseIf err.Number = 9999 Then
          MsgBox "La factura del proveedor es requerida !", 16, "Error"
          Text(0).SetFocus
     ElseIf err.Number = 9998 Then
          MsgBox "Factura ya existente para el proveedor seleccionado !", 16, "Error"
          Selecciona Text(0)
     ElseIf err.Number = 3188 Or err.Number = 3202 Then
          Call Barra("Tabla bloqueada, por favor espere ...")
          For I = 0 To 500000
               w% = DoEvents
          Next
          Resume
     ElseIf err.Number = 9998 Then
          MsgBox "Ya existe esta factura para el proveedor seleccionado !", 16, "Error"
          Text(0).SetFocus
     ElseIf err.Number = 3201 Then
          MsgBox "Bodega no existente !", 16, "Error"
          Selecciona Text(4)
      ElseIf err.Number = 3315 Or err.Number = 3134 Then
          MsgBox "Al menos uno de los campos vacíos es requerido !", 16, "Error"
          Selecciona Text(0)
     ElseIf err.Number = 3464 Or err.Number = 3075 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
          Selecciona Text(0)
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Inserta"
     End If
     On Error GoTo 0
     Call Barra("")
End Function
Private Sub Limpia(Modo As Boolean)
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     Combo1.Text = ""
     Subtotal = ""
     Descuento = ""
     Impuesto = ""
     Total = ""
     NomProv = ""
     For I = 0 To 8
          If I <> 4 Then Text(I) = ""
     Next
     For I = 1 To ListView3.ListItems.Count
          ListView3.ListItems(I).SubItems(1) = "0.00"
     Next
     On Error Resume Next
     For I = 3 To 7
          Toolbar1.Buttons(I).Enabled = Modo
          Inicio.SubPop(I - 1).Enabled = Modo
     Next
     On Error GoTo 0
End Sub
Private Function GeneraCXP() As Boolean
     'On Error GoTo Errores
     Dim NumEnt$
     Dim cuentacxp As String
     Dim tipoprov As String
     NumEnt = Format(Text(1), "00000000")
     Call Barra("Actualizando Cuentas por Pagar !")
     Dim CodPRov$
     Dim Numero$
     Dim SysParam As Recordset
     S = "select * from parametros where cia='" + CiA + "'"
     Set SysParam = DatOS.OpenRecordset(S, 4)
     'Inserta en el diario de cuentas por pagar
     CodPRov = Format(Text(3), "0000000000")
     Numero = Format(Text(0), "0000000000")
     S = "insert into diario_cxp(compania,n_docum,tip_doc,fecha,codigo,plazo,"
     S = S + "monto,moneda,tipo_cambi,concepto,banco,doc_afec,cierra) "
     S = S + "values('" + CiA + "','" 'La compania
     S = S + Numero + "','" 'El numero de entrada
     S = S + "1',#" 'El tipo de documento
     S = S + Format(Text(2), "m/d/yyyy") + "#,'" 'La fecha
     S = S + CodPRov + "'," 'El codigo del provedor
     S = S + Text(6) + "," 'El plazo de la factura
     S = S & Doble(Total.Caption) & ","  'El monto de la factura
     S = S + "'" + Proveedor!Moneda + "'," 'La moneda
     S = S & TCambio & ","  'El tipo de cambio
     S = S + "'Generada por IyF',"  'El concepto
     S = S + "'01','"  'El banco
     S = S + Numero + "',1)" 'El documento afectado
     DatOS.Execute S, 128
     Dim Tabla As Recordset
     
     'Inserta el detalle contable del inventario
     S = "insert into movicon_cxp(n_docum,tip_doc,C_CONTA,"
     S = S + "debe,haber,codigo,compania) "
     S = S + " select '" + Numero + "','" 'El numero de factura del proveedor
     S = S + "1'," 'El tipo de documento
     S = S + "tipos.c_conta," 'El codigo contable del tipo de articulo
     S = S + "sum(desgent.total_brut)," 'El monto al debe
     S = S + "0,'" 'El monto al haber
     S = S + CodPRov + "','" + CiA + "' " 'El codigo del provedor y la compania
     S = S + "  from tipos,articulos,desgent "
     S = S + "  where desgent.cia='" + CiA
     S = S + "' and desgent.bodega='" + Text(4)
     S = S + "' and desgent.n_entrada='" + NumEnt
     S = S + "' and articulos.cia=desgent.cia "
     S = S + "  and articulos.c_articulo=desgent.c_articulo "
     S = S + "  and tipos.cia=desgent.cia "
     S = S + "  and tipos.c_tipo_art=articulos.c_tipo_art "
     S = S + "  group by tipos.c_conta"
     DatOS.Execute S, 128
     Debug.Print "Linea 1 : " + S
     
     'Inserta el detalle contable del impuesto
     If Val(Impuesto.Caption) <> 0 Then
        If IsNull(SysParam!cuentaisv) Or SysParam!cuentaisv = "" Then
             S = "La cuenta de impuestos en Parámetros es incorrecta !"
             MsgBox S, 16, "Error"
             Exit Function
        End If
        S = "insert into movicon_cxp(n_docum,tip_doc,c_conta,"
        S = S + "debe,haber,codigo,compania) values('"
        S = S + Numero + "','" 'El numero de factura del proveedor
        S = S + "1','" 'El tipo de documento
        S = S + SysParam!cuentaisv + "'," 'El codigo contable del impuesto
        S = S & Doble(Impuesto.Caption) & "," 'El monto al debe
        S = S + "0,'" 'El monto al haber
        S = S + CodPRov + "','" + CiA + "')" 'El codigo del provedor
        DatOS.Execute S, 128
     End If
     
     'Inserta el movimiento contable de la cuenta por pagar lo Toma de Parámetros
     
     'If IsNull(SysParam!cuentacxp) Or SysParam!cuentacxp = "" Then
     '     S = "La cuenta de CXP en Parámetros es incorrecta !"
     '     MsgBox S, 16, "Error"
     '     Exit Function
     'End If
     'S = "insert into movicon_cxp(n_docum,tip_doc,c_conta,"
     'S = S + "debe,haber,codigo,compania) values('"
     'S = S + Numero + "','" 'El numero de factura
     'S = S + "1','" 'El tipo de documento
     'S = S + SysParam!cuentacxp + "',0," 'El monto al debe
     'S = S & Doble(Total.Caption) 'El monto al haber
     'S = S + ",'" + CodPRov + "','" + CiA + "')" 'El codigo del provedor
     'DatOS.Execute S, 128
     
     'Inserta el Detalle Contable de la CXP, lo toma segun el tipo de Proveedor
     S = "select c_tipo_pro from proveedores where compania='" + CiA
     S = S + "' and  codigo='" + CodPRov + " '"
     Set Temp = DatOS.OpenRecordset(S, 4)
     
     tipoprov = Temp!c_tipo_pro
     
     S = "select * from tipoprov where compania='" + CiA
     S = S + "' and  c_tipo_pro='" + tipoprov + " '"
     Set Temp = DatOS.OpenRecordset(S, 4)
     
     cuentacxp = Temp!c_conta

     S = "insert into movicon_cxp(n_docum,tip_doc,c_conta,"
     S = S + "debe,haber,codigo,compania) values('"
     S = S + Numero + "','" 'El numero de factura
     S = S + "1','" 'El tipo de documento
     S = S + cuentacxp + "',0," 'El monto al debe
     S = S & Doble(Total.Caption) 'El monto al haber
     S = S + ",'" + CodPRov + "','" + CiA + "')" 'El codigo del provedor
     DatOS.Execute S, 128
          
     'Inserta el movimiento contable del descuento
     If Val(Descuento.Caption) <> 0 Then
        If IsNull(SysParam!cuentadsc) Or SysParam!cuentadsc = "" Then
             S = "La cuenta de Descuentos sobre Compras en Parámetros es incorrecta !"
             MsgBox S, 16, "Error"
             Exit Function
        End If
     
        S = "insert into movicon_cxp(n_docum,tip_doc,c_conta,"
        S = S + "debe,haber,codigo,compania) values('"
        S = S + Numero + "','" 'El numero de factura
        S = S + "1','" 'El tipo de documento
        S = S + SysParam!cuentadsc + "',0,"
        S = S & Doble(Descuento.Caption) & ",'" 'El monto al debe
        S = S + CodPRov + "','" + CiA + "')" 'El codigo del provedor
        DatOS.Execute S, 128
     End If
     
     GeneraCXP = True
     Set Tabla = Nothing
Errores:
     If err.Number = 3022 Then
          S = "El número de factura ya existe para el proveedor seleccionado !"
          S = S + Chr(13) + "Deberá incluir la entrada manualmente en Cuentas por Pagar."
          MsgBox S, 64, "GeneraCXP"
     ElseIf err.Number > 0 Then
          S = "El siguiente error ocurrió al generar el movimiento en Cuentas por Pagar :"
          S = S + Chr(13) + err.Description
          S = S + "Error No. : " + Str(err.Number)
          MsgBox S, 16, "GeneraCXP"
     End If
     On Error GoTo 0
     MousePointer = 0
     Call Barra("")
End Function
Public Sub Carga(Tipo As Integer)
     Compra = Tipo
     If Compra = 0 Then 'Entrada por compra
          Caption = "Entradas por Compra"
          Combo1.Enabled = True
          S = "select * from ordenes where cia='" + CiA + "' order by numero"
          'Set Ordenes = DatOS.OpenRecordset(S, dbOpenSnapshot)
          'Do Until Ordenes.EOF
          '     Combo1.AddItem Ordenes!Numero
          '     Ordenes.MoveNext
          'Loop
     ElseIf Compra = 1 Then 'Entrada por ajuste
          Caption = "Entradas por Ajuste"
          Text(5).Enabled = False
          Text(7).Enabled = False
     ElseIf Compra = 2 Then 'Devoluciones sobre compras
          MsgBox "Error", 16
     End If
     Load Me
End Sub
Private Function BackOrder() As Boolean
     Call Barra("Generando Back Order ...")
     Dim NumBack$
     Dim COnta&
     NumBack = GeneraNumero("BK")
     S = "insert into backorder(numero,orden,fecha,entrada,proveedor,bodega,cia)"
     S = S + "values('" + NumBack
     S = S + "','" + Combo1.Text 'El numero de orden de compra
     S = S + "',#" + Format(Date, "m/d/yyyy") 'La fecha
     S = S + "#,'" + Text(1) 'El numero de entrada
     S = S + "','" + Text(3) 'El provedor
     S = S + "','" + Text(4) + "','" + CiA + "')" 'La bodega
     DatOS.Execute S, 128
     S = "select * from desgord where cia='" + CiA + "' and orden='" + Combo1.Text + "'"
     Dim Desgord As Recordset
     Set Desgord = DatOS.OpenRecordset(S)
     For I = 1 To ListView2.ListItems.Count
          Set Item = ListView2.ListItems(I)
          Desgord.FindFirst "codart='" + Item.Text + "'"
          If Not Desgord.EOF Then
               If Desgord!Cantidad > Doble(Item.SubItems(5)) Then
                    S = "insert into desgorder(numero,codart,cantini,cantfin,cia)"
                    S = S + "values('" + NumBack
                    S = S + "','" + Item.Text
                    S = S + "'," & Desgord!Cantidad
                    S = S + "," & Doble(Item.SubItems(5)) & ",'" + CiA + "')"
                    DatOS.Execute S, 128
                    COnta = COnta + DatOS.RecordsAffected
               End If
          End If
          w% = DoEvents
     Next
     Set Desgord = Nothing
     If COnta > 0 Then BackOrder = True
End Function
Private Function Arma() As String
     S = "insert into entradas (n_entrada,c_bodega,impuesto,"
     S = S + "monto,monto_real,f_entrada,plazo,fact_comp,c_prove,"
     S = S + "tipo,descuento,referencia,cia)"
     S = S + "values('" + Format(Text(1), "00000000") + "','" 'El numero de entrada
     S = S + Text(4) + "'," 'El codigo de bodega
     S = S + Trim(Doble(Impuesto.Caption)) + "," 'El monto del impuesto
     S = S + Trim(Doble(Subtotal.Caption)) + "," 'El subtotal
     S = S + Trim(Doble(Total.Caption)) + ",#" 'El total de la entrada
     S = S + Format(Text(2), "m/d/yyyy") + "#," 'La fecha
     S = S + Text(6) + ",'"  'El plazo
     S = S + Text(0) + "','" 'La factura del provedor
     S = S + Format(Text(3), "0000000000") + "'," 'El codigo del provedor
     S = S + Trim(Compra) + "," 'El tipo de entrada, 0=compra,1=ajuste
     S = S + Trim(Doble(Descuento.Caption)) + ",'" 'El descuento
     S = S + Text(8) + "','" + CiA + "')" 'referencia
     Arma = S
End Function
Private Function NewNum() As String
     Dim Temp As Recordset
     S = "select max(n_entrada) from entradas "
     S = S + "where cia='" + CiA + "' and c_bodega='" + Text(4) + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If IsNull(Temp(0)) Then
          NewNum = "00000001"
     Else
          NewNum = Format(Val(Temp(0)) + 1, "00000000")
     End If
     Set Temp = Nothing
End Function
