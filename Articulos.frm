VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Articulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Articulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   6300
   ScaleWidth      =   11100
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Articulos.frx":030A
      Height          =   3315
      Left            =   60
      OleObjectBlob   =   "Articulos.frx":031E
      TabIndex        =   10
      Top             =   420
      Width           =   10965
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   8640
      TabIndex        =   11
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "inserta"
            Object.ToolTipText     =   "Incluir un artículo"
            Object.Tag             =   "Insertar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modifica"
            Object.ToolTipText     =   "Modificar el artículo seleccionado"
            Object.Tag             =   "Modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "borra"
            Object.ToolTipText     =   "Borrar el artículo seleccionado"
            Object.Tag             =   "Borrar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "imprime"
            Object.ToolTipText     =   "Imprimir la lista de artículos"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "refresca"
            Object.ToolTipText     =   "Refrescar la lista"
            Object.Tag             =   "Refrescar Lista"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sale"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   3750
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   4154
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Características"
      TabPicture(0)   =   "Articulos.frx":0D11
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Utilidad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Comision"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Tipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Descuento"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "IV"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Image1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SubGrupo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Talla"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label15"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Color"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Estilo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label20"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Coleccion"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "&Adicional"
      TabPicture(1)   =   "Articulos.frx":0D2D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "CostoEx"
      Tab(1).Control(2)=   "Costo"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(5)=   "Auxiliar"
      Tab(1).Control(6)=   "Label17"
      Tab(1).Control(7)=   "Alterno"
      Tab(1).Control(8)=   "Marca"
      Tab(1).Control(9)=   "Label16"
      Tab(1).Control(10)=   "Label14"
      Tab(1).Control(11)=   "CostoAnt"
      Tab(1).Control(12)=   "Provedor"
      Tab(1).Control(13)=   "Label4"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "&Buscar"
      TabPicture(2)   =   "Articulos.frx":0D49
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Check1"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(3)=   "Text1"
      Tab(2).Control(4)=   "Command1"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Unidades de &Venta"
      TabPicture(3)   =   "Articulos.frx":0D65
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label1"
      Tab(3).Control(2)=   "ListView1"
      Tab(3).Control(3)=   "UniBoton(2)"
      Tab(3).Control(4)=   "UniBoton(1)"
      Tab(3).Control(5)=   "UniBoton(0)"
      Tab(3).Control(6)=   "DescUni"
      Tab(3).Control(7)=   "NumUni"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "&Precios de Venta"
      TabPicture(4)   =   "Articulos.frx":0D81
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Existencias"
      TabPicture(5)   =   "Articulos.frx":0D9D
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ListView3"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton Command1 
         Height          =   675
         Left            =   -69120
         Picture         =   "Articulos.frx":0DB9
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Realizar la búsqueda seleccionada"
         Top             =   870
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   -72210
         TabIndex        =   28
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar por"
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
         Height          =   915
         Left            =   -74610
         TabIndex        =   25
         Top             =   600
         Width           =   2295
         Begin VB.OptionButton Opcion1 
            Caption         =   "Código de artículo"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   300
            Width           =   1815
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Descripción"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   540
            Value           =   -1  'True
            Width           =   1905
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Buscar en todo el texto"
         Height          =   255
         Left            =   -68250
         TabIndex        =   24
         Top             =   870
         Width           =   2145
      End
      Begin VB.TextBox NumUni 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   -68790
         TabIndex        =   20
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox DescUni 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -68790
         TabIndex        =   19
         Top             =   480
         Width           =   3285
      End
      Begin VB.CommandButton UniBoton 
         Caption         =   "&Nueva"
         Height          =   555
         Index           =   0
         Left            =   -67560
         Picture         =   "Articulos.frx":10C3
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1230
         Width           =   1125
      End
      Begin VB.CommandButton UniBoton 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   -66390
         Picture         =   "Articulos.frx":11C5
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1230
         Width           =   1125
      End
      Begin VB.CommandButton UniBoton 
         Caption         =   "&Borrar"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   -65220
         Picture         =   "Articulos.frx":12C7
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1230
         Width           =   1125
      End
      Begin ComctlLib.ListView ListView3 
         Height          =   1935
         Left            =   -74940
         TabIndex        =   14
         Top             =   360
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Bodega"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Lote"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Existencia"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Ult. Entrada"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "F.U. Entrada"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Tipo E"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "U. Salida"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "F.U. Salida"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Tipo S"
            Object.Width           =   529
         EndProperty
      End
      Begin ComctlLib.ListView ListView2 
         Height          =   1905
         Left            =   -74940
         TabIndex        =   15
         Top             =   360
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   3360
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Bodega"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Descripción"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Precio"
            Object.Width           =   1764
         EndProperty
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   1905
         Left            =   -74940
         TabIndex        =   21
         Top             =   360
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   3360
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Descripción"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Cantidad"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Coleccion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3360
         TabIndex        =   54
         Top             =   1080
         Width           =   4155
      End
      Begin VB.Label Label20 
         Caption         =   "Colección:"
         Height          =   255
         Left            =   2430
         TabIndex        =   53
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   225
         Left            =   -70950
         TabIndex        =   52
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Provedor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -69870
         TabIndex        =   51
         Top             =   900
         Width           =   4155
      End
      Begin VB.Label Estilo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4290
         TabIndex        =   50
         Top             =   1860
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Estilo:"
         Height          =   225
         Left            =   3720
         TabIndex        =   49
         Top             =   1950
         Width           =   525
      End
      Begin VB.Label Color 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5970
         TabIndex        =   48
         Top             =   1470
         Width           =   1545
      End
      Begin VB.Label Label15 
         Caption         =   "Color:"
         Height          =   315
         Left            =   5400
         TabIndex        =   47
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label Talla 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3360
         TabIndex        =   46
         Top             =   1470
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Talla:"
         Height          =   315
         Left            =   2850
         TabIndex        =   45
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label SubGrupo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3360
         TabIndex        =   44
         Top             =   750
         Width           =   4155
      End
      Begin VB.Label Label9 
         Caption         =   "SubGrupo:"
         Height          =   255
         Left            =   2430
         TabIndex        =   43
         Top             =   750
         Width           =   885
      End
      Begin VB.Label CostoAnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -73320
         TabIndex        =   42
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costo Anterior :"
         Height          =   225
         Left            =   -74655
         TabIndex        =   41
         Top             =   660
         Width           =   1290
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   225
         Left            =   -70560
         TabIndex        =   40
         Top             =   1500
         Width           =   570
      End
      Begin VB.Label Marca 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -69870
         TabIndex        =   39
         Top             =   1410
         Width           =   2865
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   8910
         Stretch         =   -1  'True
         ToolTipText     =   "Imagen del artículo seleccionado"
         Top             =   390
         Width           =   1995
      End
      Begin VB.Label Alterno 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -73320
         TabIndex        =   38
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código Alterno :"
         Height          =   225
         Left            =   -74715
         TabIndex        =   37
         Top             =   1740
         Width           =   1320
      End
      Begin VB.Label Auxiliar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -69855
         TabIndex        =   36
         Top             =   510
         Width           =   5595
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descripción Auxiliar :"
         Height          =   225
         Left            =   -71640
         TabIndex        =   35
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costo Local :"
         Height          =   225
         Left            =   -74400
         TabIndex        =   34
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Costo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -73320
         TabIndex        =   33
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label CostoEx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -73320
         TabIndex        =   32
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costo Extranjero :"
         Height          =   225
         Left            =   -74835
         TabIndex        =   31
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto a buscar :"
         Height          =   225
         Left            =   -72210
         TabIndex        =   30
         Top             =   930
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unidades :"
         Height          =   225
         Left            =   -69690
         TabIndex        =   23
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   225
         Left            =   -69900
         TabIndex        =   22
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label IV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1395
         TabIndex        =   13
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Impuesto :"
         Height          =   225
         Left            =   465
         TabIndex        =   12
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Descuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1395
         TabIndex        =   8
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuento (%) :"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Tipo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3360
         TabIndex        =   6
         Top             =   390
         Width           =   4155
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grupo :"
         Height          =   225
         Left            =   2700
         TabIndex        =   5
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comisión :"
         Height          =   225
         Left            =   450
         TabIndex        =   4
         Top             =   1140
         Width           =   870
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Utilidad (%)  :"
         Height          =   225
         Left            =   300
         TabIndex        =   3
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label Comision 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1395
         TabIndex        =   2
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Utilidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1395
         TabIndex        =   1
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5670
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":13C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":14DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":15ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":16FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":1A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":1D33
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":1E45
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":1F57
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":2271
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":258B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":28A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Articulos.frx":2BBF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Articulo As Recordset
Dim Boton%
Dim Item As ListItem
Dim S$
Public ConTinuO As Boolean
Private Function Borra() As Boolean
     On Error GoTo Errores
     Dim Uno As Boolean
     Dim Imagen$
     Dim Cod$
     Cod = DBGrid1.Columns(0)
     If DBGrid1.SelBookmarks.Count = 1 Then
          If MsgBox("Desea borrar el artículo seleccionado ?", 36, DBGrid1.Columns(1)) = 6 Then
               Imagen = DirTrA + "Imagenes\" + Nulo(Articulo!Imagen)
               S = "delete from articulos where cia='" + CiA
               S = S + "' and c_articulo='" + Cod + "'"
               Data1.Database.Execute S, 128
               Kill Imagen
               Call GBitacora(3, "Articulo: " + Cod + " " + DBGrid1.Columns(1))
               Borra = True
               MsgBox "Un (1) artículo eliminado !", 64, "Artículos"
          End If
          If Data1.Recordset.EOF Then Call Limpia(False)
     ElseIf DBGrid1.SelBookmarks.Count > 1 Then
          If MsgBox("Desea borrar los artículos seleccionados ?", 36, "Artículos") = 6 Then
               Uno = True
               MousePointer = 11
               Call Barra("Borrando registros ...")
               Bar1.Visible = True
               Bar1.Max = DBGrid1.SelBookmarks.Count
               For I = 0 To DBGrid1.SelBookmarks.Count - 1
                    DBGrid1.Bookmark = DBGrid1.SelBookmarks(I)
                    Imagen = DirTrA + "Imagenes\" + Nulo(Articulo!Imagen)
                    S = "delete from articulos where c_articulo='" + Cod + "'"
                    Data1.Database.Execute S, 128
                    Kill Imagen
                    Call GBitacora(3, "Articulo: " + Cod + " " + DBGrid1.Columns(1))
                    Bar1.Value = I + 1
                    w% = DoEvents
               Next
               Call Columnas(True)
               Borra = True
               MsgBox (I + 1) & " artículos eliminados !", 64, "Artículos"
               MousePointer = 0
               Call Barra("")
               Bar1.Visible = False
          End If
          If Data1.Recordset.EOF Then Call Limpia(False)
     Else
          S = "Debe seleccionar el artículo haciendo click sobre la flecha negra a la izquierda"
          MsgBox S, 48, "Artículos"
          DBGrid1.SetFocus
     End If
Errores:
     If err.Number = 3200 Then
          S = "Existen registros asociados a este artículo,"
          S = S + Chr(13) + "    No es posible su eliminación."
          MsgBox S, 48, Item.Text
          If Uno Then Resume Next
     ElseIf err.Number = 53 Then
          Resume Next
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Borra"
     End If
     On Error GoTo 0
End Function
Private Sub Command1_Click()
     MousePointer = 11
     Call Barra("Un momento por favor ...")
     Dim Criterio$
     Criterio = IIf(Opcion1(0).Value, "código", "descripción")
     Criterio = Criterio + " like '"
     If Check1.Value Then Criterio = Criterio + "*"
     Criterio = Criterio + Text1 + "*'"
     Data1.Recordset.FindFirst Criterio
     If Not Data1.Recordset.NoMatch Then
          DBGrid1.Bookmark = Data1.Recordset.Bookmark
          MousePointer = 0
          Call Barra("")
          Exit Sub
     End If
     Selecciona Text1
     MousePointer = 0
     Call Barra("No se encontró el registro !", 1)
End Sub
Private Sub DBGrid1_DblClick()
     Call Modifica
End Sub
Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
     S = "select c_articulo as Código,d_articulo as Descripción,"
     S = S + "iif(sino_impu=0,'No','Si') as IV,p_compra as Costo,"
     S = S + "minimo as Mínimo,maximo as Máximo "
     S = S + "from articulos where cia='" + CiA + "' and estado<>1"
     If ColIndex = 0 Then
          S = S + " order by c_articulo"
     ElseIf ColIndex = 1 Then
          S = S + " order by d_articulo"
     End If
     Data1.RecordSource = S
     Data1.Refresh
     Call Columnas(False)
End Sub
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then PopupMenu Inicio.Popup, , , , Inicio.SubPop(0)
End Sub
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     If MousePointer = 11 Then Exit Sub
     If Data1.Recordset.EOF Then Exit Sub
     MousePointer = 11
     Data1.Recordset.Bookmark = DBGrid1.Bookmark
     Dim Temp As Recordset
     'Inicio.SubPop(1).Enabled = True
     'Inicio.SubPop(2).Enabled = True
     Toolbar1.Buttons(2).Enabled = True
     Toolbar1.Buttons(3).Enabled = True
     Articulo.FindFirst "c_articulo='" + DBGrid1.Columns(0) + "'"
     Comision.Caption = ""
     Utilidad.Caption = ""
     Descuento.Caption = ""
     Tipo.Caption = ""
     Provedor.Caption = ""
     CostoEx.Caption = ""
     Costo.Caption = ""
     IV.Caption = ""
     Auxiliar.Caption = ""
     Alterno.Caption = ""
     Marca.Caption = ""
     CostoAnt.Caption = ""
     SubGrupo.Caption = ""
     Color.Caption = ""
     Talla.Caption = ""
     Estilo.Caption = ""
     If Not Articulo.NoMatch Then
          Tipo.Caption = Nulo(Articulo!d_tipo_art)
          CostoEx.Caption = FormatNumber(Articulo!CostoEx, DeCiMaleS)
          Costo.Caption = FormatNumber(Articulo!p_compra, DeCiMaleS)
          CostoAnt.Caption = FormatNumber(Articulo!CostoAnt, DeCiMaleS)
          IV.Caption = IIf(Articulo!sino_impu = 0, "No", "Si")
          If Not IsNull(Articulo!Imagen) And Articulo!Imagen <> "" Then
               If Dir(DirTrA + "Imagenes\" + Articulo!Imagen) <> "" Then
                    Image1.Picture = LoadPicture(DirTrA + "Imagenes\" + Articulo!Imagen)
               Else
                    Call Barra("Imagen no encontrada ! - " + Articulo!Imagen, 1)
               End If
          Else
               Image1.Picture = LoadPicture("")
          End If
     End If
     Call ListPrecios(DBGrid1.Columns(0))
     Call ListaExistencias(DBGrid1.Columns(0))
     MousePointer = 0
End Sub
Private Sub Form_Activate()
     Call Menus(True, Me)
End Sub
Private Sub Form_Deactivate()
     Call Menus(False, Me)
End Sub
Private Sub Form_Load()
     Call Posiciona(Me, 0)
     Set Actual = Me
     Data1.DatabaseName = DatOS.Name
     Data1.Refresh
     MousePointer = 11
     Call Barra("Un momento ... Cargando artículos !")
     If UsaLotes = 0 Then
          With ListView3
          .ColumnHeaders(1).Width = .ColumnHeaders(1).Width + .ColumnHeaders(2).Width
          .ColumnHeaders(2).Width = 0
          End With
     End If
     Show
     Refresh
     Call Barra(Carga & " artículos encontrados", 1)
     If Data1.Recordset.EOF Then SSTab1.Enabled = False
     MousePointer = 0
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Boton = 2 Then
          PopupMenu Inicio.Popup, , x, y, Inicio.SubPop(0)
     End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
     On Error Resume Next
     Call Posiciona(Me, 1)
     Articulo.Close
     Call Menus(False, Me)
End Sub
Private Sub Modifica()
     If DBGrid1.SelBookmarks.Count > 1 Then Exit Sub
     MousePointer = 11
     Dim Marca As Variant
     Marca = DBGrid1.Row
     Call Load(DetalleArticulo)
     Call DetalleArticulo.DatosArticulo(DBGrid1.Columns(0))
     DetalleArticulo.Show 1
     'Articulo.Requery
     DBGrid1.Row = Marca
     MousePointer = 0
End Sub

Private Sub Listview1_DblClick()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then
          Call UniBoton_Click(1)
     End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     UniBoton(1).Enabled = True
     UniBoton(2).Enabled = True
End Sub
Private Sub NumUni_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
     Case 13
          Call UniBoton_Click(1)
     Case 8, 46, 48 To 57
     Case Else
          KeyAscii = 0
     End Select
End Sub
Private Sub Opcion1_Click(Index As Integer)
     Selecciona Text1
End Sub
Private Sub ListPrecios(Codigo As String)
     Dim Temp As Recordset
     ListView2.ListItems.Clear
     If PreCBoD = 1 Then
          S = "select d_bodega,monto,unidades.descripcion "
          S = S + "from bodegas,unidades,precios "
          S = S + "where precios.cia='" + CiA
          S = S + "' and precios.codart='" + Codigo
          S = S + "' and bodegas.cia=precios.cia "
          S = S + "  and unidades.cia=precios.cia "
          S = S + "  and bodegas.c_bodega=precios.codbod "
          S = S + "and unidades.unidades=precios.numuni "
          S = S + "and unidades.codart=precios.codart "
          ListView2.ColumnHeaders(1).Text = "Bodega"
          ListView2.ColumnHeaders(2).Text = "Descripcion"
          ListView2.ColumnHeaders(3).Text = "Precio"
     Else
          S = "select descripcion,monto from listaprecios "
          S = S + "where cia='" + CiA + "' and codart='"
          S = S + Codigo + "' order by codigo"
          ListView2.ColumnHeaders(1).Text = "Descripción"
          ListView2.ColumnHeaders(2).Text = "Precio"
          ListView2.ColumnHeaders(3).Text = Chr(32)
     End If
     Set Temp = DatOS.OpenRecordset(S, dbOpenSnapshot)
     Do Until Temp.EOF
          Set Item = ListView2.ListItems.Add(, , , , 8)
          Item.Text = Temp(0)
          If PreCBoD = 1 Then
               Item.SubItems(2) = Format(Temp!Monto, "standard")
          Else
               Item.SubItems(1) = Format(Temp!Monto, "standard")
          End If
          w% = DoEvents
          Temp.MoveNext
     Loop
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
     If SSTab1.Tab = 1 Then Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then Command1_Click
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF3 Then
          S = "select c_articulo as Código,d_articulo as Descripción "
          S = S + "from articulos where cia='" + CiA + "' and estado<>1 order by d_articulo"
          Call Lista.Carga(Text1, S, "artículos")
          Lista.Show 1
          Text1.SetFocus
     End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
     Select Case Button.Key
     Case "inserta"
          If ReadKey("IYFDV") = "11" Then
               Dim Temp As Recordset
               S = "select count(*) from articulos where cia='" + CiA + "'"
               Set Temp = DatOS.OpenRecordset(S)
               If Not IsNull(Temp(0)) Then
                    If Temp(0) >= 15 Then
                         S = "El máximo de artículos ya se alcanzó en esta versión"
                         MsgBox S, 16, "Inventario y Facturación - Versión Demo "
                         Exit Sub
                    End If
               End If
          End If
          ConTinuO = True        ' para que siga hasta que unload con cancel
          DetalleArticulo.Show 1
          Call Limpia(False)
          Data1.Recordset.Requery
          With Data1.Recordset
          If Not .EOF Then
               .MoveLast
               If .RecordCount > 0 Then SSTab1.Enabled = True
          End If
          End With
     Case "modifica"
          DetalleArticulo.Check5.Visible = False
          ConTinuO = False
          Call Modifica
     Case "borra"
          If Borra Then Call Columnas(True)
     Case "sale"
          Unload Me
     Case "refresca"
          Call Columnas(True)
     Case "imprime"
          For Each Form In Forms
               If Form.Caption = "Reporte de Artículos" Then
                    Form.ZOrder
                    Exit Sub
               End If
          Next
          Dim forma As New RepPrecios
          Call forma.Carga("articulos")
     End Select
End Sub
Private Sub Limpia(Modo As Boolean)
     Comision.Caption = ""
     Utilidad.Caption = ""
     Descuento.Caption = ""
     Provedor.Caption = ""
     Costo.Caption = ""
     CostoEx.Caption = ""
     CostoAnt.Caption = ""
     Auxiliar.Caption = ""
     Alterno.Caption = ""
     Marca.Caption = ""
     IV.Caption = ""
     Tipo.Caption = ""
     Toolbar1.Buttons(2).Enabled = Modo
     Toolbar1.Buttons(3).Enabled = Modo
     Inicio.SubPop(1).Enabled = Modo
     Inicio.SubPop(2).Enabled = Modo
End Sub
Private Function Carga() As Long
     S = "select c_articulo,sino_comi,porc_util,porc_desc1,costoex,"
     S = S + "d_tipo_art,sino_impu,p_compra,aux,"
     S = S + "alterno,imagen,marca,c_prove,costoant,CSubgrupo,CColeccion,Ctalla,CColor,CEstilo "
     S = S + "from articulos left join tipos "
     S = S + "on tipos.cia=articulos.cia "
     S = S + "and tipos.c_tipo_art=articulos.c_tipo_art "
     S = S + "where articulos.cia='" + CiA + "' and estado<>1"
     Set Articulo = DatOS.OpenRecordset(S)
     If Not Articulo.EOF Then
          Articulo.MoveLast
          Carga = Articulo.RecordCount
          Articulo.MoveFirst
     End If
     S = "select c_articulo as Código,d_articulo as Descripción,"
     S = S + "iif(sino_impu=0,'No','Si') as IV,p_compra as Costo,"
     S = S + "minimo as Mínimo,maximo as Máximo "
     S = S + "from articulos where cia='" + CiA + "' and estado<>1 order by C_articulo "
     Data1.RecordSource = S
     Data1.Refresh
     Call Columnas
End Function
Public Sub Columnas(Optional Refresca As Boolean, Optional Xx As Integer, Optional CodArt As String)
If Refresca Then Data1.Recordset.Requery
    
    If Xx = 1 Then
        Criterio = "código"
        Criterio = Criterio + " = '"
        Criterio = Criterio + CodArt + "'"
        Data1.Recordset.FindFirst Criterio
        If Not Data1.Recordset.NoMatch Then
          DBGrid1.Bookmark = Data1.Recordset.Bookmark
        End If
    End If
    
     DBGrid1.Columns(0).Width = 1900
     DBGrid1.Columns(1).Width = 4000
     DBGrid1.Columns(2).Width = 500
     DBGrid1.Columns(3).Width = 1600
     S = "###,###,##0"
     If DeCiMaleS > 0 Then
          S = S + "."
          For I = 1 To DeCiMaleS
               S = S + "0"
          Next
     End If
     DBGrid1.Columns(3).NumberFormat = S
     DBGrid1.Columns(3).Alignment = 1
     DBGrid1.Columns(4).NumberFormat = "###,###,##0"
     DBGrid1.Columns(4).Alignment = 1
     DBGrid1.Columns(4).Width = 1200
     DBGrid1.Columns(5).NumberFormat = "###,###,##0"
     DBGrid1.Columns(5).Width = 1200
     DBGrid1.Columns(5).Alignment = 1
End Sub
Private Sub ListUni(Codigo$)
     Dim Temp As Recordset
     ListView1.ListItems.Clear
     S = "select descripcion,unidades "
     S = S + "from unidades where cia='" + CiA + "' and codart='" + Codigo + "'"
     Set Temp = DatOS.OpenRecordset(S, dbOpenSnapshot)
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add(, , , , 9)
          Item.Text = Temp(0)
          Item.SubItems(1) = Temp(1)
          w% = DoEvents
          Temp.MoveNext
     Loop
End Sub
Private Sub ListaExistencias(Codigo As String)
     Dim Temp As Recordset
     ListView3.ListItems.Clear
     S = "select d_bodega,existencia,doc_ult_en,f_ult_ent,"
     S = S + "doc_ult_sa,f_ult_sal,tipoentrada,tiposalida,"
     S = S + "existencias.c_bodega,lote "
     S = S + "from existencias left join bodegas "
     S = S + "on  bodegas.cia=existencias.cia "
     S = S + "and bodegas.c_bodega=existencias.c_bodega "
     S = S + "where existencias.cia='" + CiA + "' and c_articulo='" + Codigo
     S = S + "' order by existencias.c_bodega,lote"
     Set Temp = DatOS.OpenRecordset(S, dbOpenSnapshot)
     Do Until Temp.EOF
          Set Item = ListView3.ListItems.Add '(, , , , 10)
          Item.Text = ConVierte(Temp!c_bodega + " " + Temp!d_bodega, "T")
          Item.SubItems(1) = Nulo(Temp!Lote)
          Item.SubItems(2) = Format(Temp(1), "standard")
          Item.SubItems(3) = Nulo(Temp(2))
          Item.SubItems(4) = Format(Temp(3), "dd/mm/yyyy")
          Item.SubItems(5) = IIf(Temp(6) = 0, "EC", "EA")
          Item.SubItems(6) = Nulo(Temp(4))
          Item.SubItems(7) = Format(Temp(5), "dd/mm/yyyy")
          Select Case Temp(7)
          Case 0
               Item.SubItems(8) = "FA" 'Facturacion
          Case 1
               Item.SubItems(8) = "RE" 'Salida
          Case 2
               Item.SubItems(8) = "TR" 'Traslado
          End Select
          w% = DoEvents
          Temp.MoveNext
     Loop
     If ListView3.ListItems.Count > 0 Then
          Dim Cant#
          For I = 1 To ListView3.ListItems.Count
               Cant = Cant + Doble(ListView3.ListItems(I).SubItems(2))
          Next
          Set Item = ListView3.ListItems.Add '(, , , , 11)
          Item.Text = "Existencia Total : "
          Item.SubItems(2) = Format(Cant, "standard")
     End If
End Sub
Private Sub UniBoton_Click(Index As Integer)
     Dim Item As ListItem
     Select Case Index
     Case 0
          Call LimpiaUni(True)
          UniBoton(1).Tag = 0
          DescUni.SetFocus
     Case 1
          If UniBoton(1).Caption = "&Modificar" Then
               Call LimpiaUni(True)
               UniBoton(1).Tag = 1
               Set Item = ListView1.SelectedItem
               DescUni.Text = Item.Text
               NumUni.Tag = Item.SubItems(1)
               NumUni.Text = Item.SubItems(1)
               DescUni.SetFocus
          ElseIf UniBoton(1).Caption = "&Aceptar" Then
               If UniBoton(1).Tag = 0 Then
                    S = "insert into unidades(codart,descripcion,unidades,cia)"
                    S = S + " values ('" + DBGrid1.Columns(0)
                    S = S + "','" + DescUni.Text + "'," 'La descripcion
                    S = S + NumUni.Text + ",'" + CiA + "')" 'El numero de unidades
                    If Procesa(S) Then
                         Set Item = ListView1.ListItems.Add(, , , , 9)
                         Item.Text = DescUni.Text
                         Item.SubItems(1) = NumUni.Text
                         S = "Unidad del artículo: " + DBGrid1.Columns(0)
                         S = S + ", Unidad: " + NumUni.Text
                         Call GBitacora(1, S)
                         DescUni.Text = ""
                         NumUni.Text = ""
                         DescUni.SetFocus
                    Else
                         Selecciona DescUni
                    End If
               ElseIf UniBoton(1).Tag = 1 Then
                    S = "update unidades set descripcion='" + DescUni.Text
                    S = S + "',unidades=" + NumUni.Text
                    S = S + " where cia='" + CiA + "' and codart='" + DBGrid1.Columns(0)
                    S = S + "' and unidades=" + NumUni.Tag
                    If Procesa(S) Then
                         S = "update precios set numuni=" + NumUni.Text
                         S = S + " where cia='" + CiA
                         S = S + "' and codart='" + DBGrid1.Columns(0)
                         S = S + "' and numuni=" + NumUni.Tag
                         If Procesa(S) Then
                              Set Item = ListView1.SelectedItem
                              Item.Text = DescUni.Text
                              Item.SubItems(1) = NumUni.Text
                              S = "Unidad del artículo: " + DBGrid1.Columns(0)
                              S = S + ", Unidad: " + NumUni.Text
                              Call GBitacora(2, S)
                              Call LimpiaUni(False)
                         End If
                    Else
                         Selecciona DescUni
                    End If
               End If
          End If
     Case 2
          If UniBoton(2).Caption = "&Borrar" Then
               Set Item = ListView1.SelectedItem
               S = "Desea borrar el registro seleccionado ?"
               If MsgBox(S, 36, Item.Text) = 6 Then
                    S = "delete from unidades where cia='" + CiA + "' and codart='"
                    S = S + DBGrid1.Columns(0) + "' and unidades=" + Item.SubItems(1)
                    If Procesa(S) Then
                         ListView1.ListItems.Remove Item.Index
                         Call LimpiaUni(False)
                         If ListView1.ListItems.Count > 0 Then
                              ListView1.SelectedItem.Selected = True
                              ListView1.SetFocus
                         End If
                    End If
               End If
          ElseIf UniBoton(2).Caption = "&Cancelar" Then
               Call LimpiaUni(False)
          End If
     End Select
End Sub
Private Function Procesa(SQL As String) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     Procesa = True
Errores:
     If err.Number = 3022 Then
          MsgBox "Crea una entrada duplicada !", 16, "Error"
     ElseIf err.Number = 3315 Or err.Number = 3134 Then
          MsgBox "Al menos uno de los campos vacíos es requerido !", 16, "Error"
     ElseIf err.Number = 3464 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Procesa"
     End If
     On Error GoTo 0
End Function
Private Sub LimpiaUni(Modo As Boolean)
     UniBoton(1).Tag = -1
     UniBoton(0).Enabled = Not Modo
     UniBoton(1).Enabled = Modo
     UniBoton(2).Enabled = Modo
     ListView1.Enabled = Not Modo
     DescUni.Enabled = Modo
     NumUni.Enabled = Modo
     DescUni.Text = ""
     NumUni.Text = ""
     If Not Modo Then
          UniBoton(1).Caption = "&Modificar"
          UniBoton(2).Caption = "&Borrar"
          UniBoton(1).Picture = ImageList1.ListImages(2).Picture
          UniBoton(2).Picture = ImageList1.ListImages(3).Picture
     Else
          UniBoton(1).Caption = "&Aceptar"
          UniBoton(2).Caption = "&Cancelar"
          UniBoton(1).Picture = ImageList1.ListImages(6).Picture
          UniBoton(2).Picture = ImageList1.ListImages(7).Picture
     End If
End Sub
