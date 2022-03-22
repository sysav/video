VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Inicio 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Facturación------>"
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8595
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "submenu11"
            Object.ToolTipText     =   "Mantenimiento de Estaciones"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "submenu10"
            Object.ToolTipText     =   "Mantenimiento de Artículos"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "espacio1"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "submenu20"
            Object.ToolTipText     =   "Entradas por Compras"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "submenu21"
            Object.ToolTipText     =   "Entradas por Ajuste"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "submenu22"
            Object.ToolTipText     =   "Facturación"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "varios12"
            Object.ToolTipText     =   "Monitoreo"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "submenu27"
            Object.ToolTipText     =   "Notas de Crédito"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "submenu31"
            Object.ToolTipText     =   "Consulta de Documentos"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "submenu30"
            Object.ToolTipText     =   "Precios de los Artículos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "submenu114"
            Object.ToolTipText     =   "Seleccionar Compañía"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "submenu115"
            Object.ToolTipText     =   "Parámetros del Sistema"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "submenu213"
            Object.ToolTipText     =   "Punto de Venta"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sale"
            Object.ToolTipText     =   "Salir del sistema"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Inicio.frx":030A
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Left            =   2790
      Top             =   1470
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   8535
      TabIndex        =   1
      Top             =   2475
      Width           =   8595
      Begin VB.Label Marquee 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marquee"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   -60
         Width           =   795
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2700
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   10301
            MinWidth        =   8819
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Object.Tag             =   ""
            Object.ToolTipText     =   "Compañía seleccionada"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Object.Tag             =   ""
            Object.ToolTipText     =   "Usuario registrado"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Object.Tag             =   ""
            Object.ToolTipText     =   "Tipo de Cambio"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Object.Tag             =   ""
            Object.ToolTipText     =   "Impuesto de ventas"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   706
            MinWidth        =   706
            Object.Tag             =   ""
            Object.ToolTipText     =   "Número de Caja"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   1350
      Top             =   1470
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   1830
      Top             =   1470
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2310
      Top             =   1470
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   690
      Top             =   1920
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
            Picture         =   "Inicio.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":07FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   690
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":0CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":100C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":1326
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":1640
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":195A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":1C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":1F8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":22A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":25C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":279C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":2976
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":2B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":2D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":2F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":321E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":3538
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":3852
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":3B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":3E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":41A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":44BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":47D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":4AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Inicio.frx":4E08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menu 
      Caption         =   "&Sistema"
      Index           =   0
      Begin VB.Menu Submenu1 
         Caption         =   "&Vendedores"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Artículos"
         Index           =   1
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Bodegas"
         Index           =   2
      End
      Begin VB.Menu Submenu1 
         Caption         =   "C&ompañías"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Estaciones"
         Index           =   4
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Costos &Directos"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "T&ipos de Pago"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Tipos de artículos"
         Index           =   7
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Proveedores"
         Index           =   8
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Sub Grupos artículos"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Tipo Proveedor"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Colecciones"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Lista de &Usos"
         Index           =   12
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Tipos de &Precios"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Estaciones"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Tallas"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Seleccionar Co&mpañía"
         Index           =   17
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Parám&etros del Sistema"
         Index           =   18
      End
      Begin VB.Menu Submenu1 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu Submenu1 
         Caption         =   "&Salir"
         Index           =   20
         Shortcut        =   ^S
      End
      Begin VB.Menu Submenu1 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu Submenu1 
         Caption         =   "Cambia&r de Usuario"
         Index           =   22
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Procesos"
      Index           =   1
      Begin VB.Menu SubMenu2 
         Caption         =   "Entradas por &Compras"
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "Entradas por &Ajuste"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "&P O S"
         Index           =   2
         Shortcut        =   ^F
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "&Requisiciones"
         Index           =   3
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "&Traslados"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "&Notas de Crédito"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "Notas de &Débito"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "&Proformas"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "De&voluciones sobre Compras"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "&Ordenes de Compra"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "P.&O.S."
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "Monitoreo Video"
         Index           =   14
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "Cierre &Mensual"
         Index           =   15
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Utilitarios"
      Index           =   2
      Begin VB.Menu SubMenu3 
         Caption         =   "&Precios"
         Index           =   0
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Consulta de Documentos"
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Consulta de E&xistencias"
         Index           =   2
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Tipo &de Cambio"
         Index           =   3
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Estadísticas de &Ventas"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Estadística de C&ompras"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Toma Física"
         Index           =   7
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Co&sto del Inventario"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Movimientos por Artículo"
         Index           =   9
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Etiquetas"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Lotes"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Usuarios"
         Index           =   14
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Niveles de Seguridad"
         Index           =   15
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Bitácora del Sistema"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "&Importar Artículos"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "De&finición de Cajero"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   20
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "T&raspaso Contable"
         Index           =   21
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Configuracion Tras&paso"
         Index           =   22
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   23
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "Respaldo&s"
         Index           =   24
         Visible         =   0   'False
      End
      Begin VB.Menu SubMenu3 
         Caption         =   "-"
         Index           =   25
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Reportes"
      Index           =   3
      Begin VB.Menu Reportes 
         Caption         =   "&Ventas"
         Index           =   0
         Begin VB.Menu subMenu4 
            Caption         =   "Ventas por &Agente"
            Index           =   0
            Tag             =   "agentes"
            Visible         =   0   'False
         End
         Begin VB.Menu subMenu4 
            Caption         =   "Ventas por &Artículo"
            Index           =   1
            Tag             =   "VentArt"
         End
         Begin VB.Menu subMenu4 
            Caption         =   "Ventas por &Cliente"
            Index           =   2
            Tag             =   "VentCli"
         End
         Begin VB.Menu subMenu4 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu subMenu4 
            Caption         =   "Estadísticas de &Ventas"
            Index           =   4
            Tag             =   "Estadis"
         End
         Begin VB.Menu subMenu4 
            Caption         =   "&Rotación de Artículos"
            Index           =   5
            Tag             =   "rotacion"
            Visible         =   0   'False
         End
         Begin VB.Menu subMenu4 
            Caption         =   "Artículos &no Vendidos"
            Index           =   6
            Tag             =   "Novend"
            Visible         =   0   'False
         End
         Begin VB.Menu subMenu4 
            Caption         =   "Acumulados Ventas"
            Index           =   7
            Tag             =   "AcumVentas"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Reportes 
         Caption         =   "&Documentos"
         Index           =   1
         Begin VB.Menu Submenu5 
            Caption         =   "&Entradas entre Fechas"
            Index           =   0
            Tag             =   "Entfec"
         End
         Begin VB.Menu Submenu5 
            Caption         =   "&Facturas entre Fechas"
            Index           =   1
            Tag             =   "Factfec"
         End
         Begin VB.Menu Submenu5 
            Caption         =   "Notas de &Débito"
            Index           =   2
            Tag             =   "notasdeb"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu5 
            Caption         =   "Notas de &Crédito"
            Index           =   3
            Tag             =   "Notcre"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu5 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu Submenu5 
            Caption         =   "&Artículos por Uso"
            Index           =   5
            Tag             =   "artuso"
         End
         Begin VB.Menu Submenu5 
            Caption         =   "&Requisiciones entre Fechas"
            Index           =   6
            Tag             =   "SalFec"
         End
         Begin VB.Menu Submenu5 
            Caption         =   "Requisiciones por &Orden"
            Index           =   7
            Tag             =   "Sallot"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu5 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu Submenu5 
            Caption         =   "&Traslados entre Fechas"
            Index           =   9
            Tag             =   "Trasl"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu5 
            Caption         =   "&Apartados y Crédito"
            Index           =   10
            Tag             =   "apartados"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Reportes 
         Caption         =   "&Adicionales"
         Index           =   2
         Begin VB.Menu Submenu6 
            Caption         =   "&Back Orders"
            Index           =   0
            Tag             =   "Back"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Resumen de Pagos"
            Index           =   1
            Tag             =   "Cierre"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu6 
            Caption         =   "C&omisiones"
            Index           =   2
            Tag             =   "Comisi"
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Pedidos por Tienda"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Existencias por Bodega"
            Index           =   4
            Tag             =   "Exis"
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Mínimos y Máximos"
            Index           =   5
            Tag             =   "minimo"
         End
         Begin VB.Menu Submenu6 
            Caption         =   "Reporte de Costo&s"
            Index           =   6
            Tag             =   "costos"
         End
         Begin VB.Menu Submenu6 
            Caption         =   "Mo&vimientos por Artículo"
            Index           =   7
            Tag             =   "movimientos"
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Lista de Precios"
            Index           =   8
            Tag             =   "Precios"
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Reporte de Precios"
            Index           =   9
         End
         Begin VB.Menu Submenu6 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu Submenu6 
            Caption         =   "E&xistencia por Lote"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu Submenu6 
            Caption         =   "&Notas de Crédito"
            Index           =   12
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Reportes 
         Caption         =   "&Matriz"
         Index           =   3
         Visible         =   0   'False
         Begin VB.Menu Submenu7 
            Caption         =   "&Procesar"
            Index           =   0
         End
         Begin VB.Menu Submenu7 
            Caption         =   "Ventas &Unidades"
            Index           =   1
         End
         Begin VB.Menu Submenu7 
            Caption         =   "Existencias xBodegas"
            Index           =   2
         End
         Begin VB.Menu Submenu7 
            Caption         =   "Ventas Montos"
            Index           =   3
         End
         Begin VB.Menu Submenu7 
            Caption         =   "RESUMEN"
            Index           =   4
         End
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Reportes de Documentos"
      Index           =   4
      Visible         =   0   'False
   End
   Begin VB.Menu Menu 
      Caption         =   "Reportes Adicionales"
      Index           =   5
      Visible         =   0   'False
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu SubPop 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Varios 
      Caption         =   "&Varios"
      Visible         =   0   'False
      Begin VB.Menu Varios1 
         Caption         =   "Cambio &Precios"
         Index           =   0
      End
      Begin VB.Menu Varios1 
         Caption         =   "&Articulos"
         Index           =   1
      End
      Begin VB.Menu Varios1 
         Caption         =   "Aparta&dos y Créditos"
         Index           =   2
      End
      Begin VB.Menu Varios1 
         Caption         =   "&Pagos Apartados y Crédito "
         Index           =   3
      End
      Begin VB.Menu Varios1 
         Caption         =   "Cierre de &Caja"
         Index           =   4
      End
      Begin VB.Menu Varios1 
         Caption         =   "&Registro de Certifcados"
         Index           =   5
      End
      Begin VB.Menu Varios1 
         Caption         =   "&Boleta de Inscripción"
         Index           =   6
      End
      Begin VB.Menu Varios1 
         Caption         =   "Clientes "
         Index           =   7
         Begin VB.Menu Frecuen 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu Frecuen2 
            Caption         =   "Ocupaciones"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu Frecuen3 
            Caption         =   "Estado Civil"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu Frecuen4 
            Caption         =   "Frecuencia de Visitas"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu Frecuen6 
            Caption         =   "Encuesta Classic"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu Frecuen5 
            Caption         =   "Encuesta"
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Varios1 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu Varios1 
         Caption         =   "&Enviar==> Datos"
         Index           =   9
         Begin VB.Menu enviar 
            Caption         =   "== ENVIAR =="
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu enviar 
            Caption         =   "Cierre de Caja"
            Index           =   1
         End
         Begin VB.Menu enviar 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu enviar 
            Caption         =   "Traslados"
            Index           =   3
         End
         Begin VB.Menu enviar 
            Caption         =   "Compras"
            Index           =   4
         End
         Begin VB.Menu enviar 
            Caption         =   "Precios"
            Index           =   5
         End
         Begin VB.Menu enviar 
            Caption         =   "Descuentos"
            Index           =   6
         End
      End
      Begin VB.Menu Varios1 
         Caption         =   "&Recibir<==Datos"
         Index           =   10
         Begin VB.Menu recibir 
            Caption         =   "== RECIBIR= ="
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu recibir 
            Caption         =   ".Cierre de Caja"
            Index           =   1
         End
         Begin VB.Menu recibir 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu recibir 
            Caption         =   ".Traslados"
            Index           =   3
         End
         Begin VB.Menu recibir 
            Caption         =   ".Compras"
            Index           =   4
         End
         Begin VB.Menu recibir 
            Caption         =   ".Precios"
            Index           =   5
         End
         Begin VB.Menu recibir 
            Caption         =   ".Descuentos"
            Index           =   6
         End
      End
      Begin VB.Menu Varios1 
         Caption         =   "Congelar Artículos"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu Varios1 
         Caption         =   "Códigos de Barra"
         Index           =   12
      End
   End
   Begin VB.Menu Ventana 
      Caption         =   "&Ventana"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu Ventanas 
         Caption         =   "&Cascada"
         Index           =   0
      End
      Begin VB.Menu Ventanas 
         Caption         =   "Mosaico &Horizontal"
         Index           =   1
      End
      Begin VB.Menu Ventanas 
         Caption         =   "Mosaico &Vertical"
         Index           =   2
      End
      Begin VB.Menu Ventanas 
         Caption         =   "&Organizar Iconos"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim forma As Form
Dim Status As Boolean

Private Sub EnViar_Click(Index As Integer)
EnviarD.Tag = Index
Select Case Index
     Case 1     ' enviar cierre de caja
          EnviarD.Label1.Caption = "CIERRE DE CAJA"
          EnviarD.TBodega.Enabled = False
          EnviarD.Show 1
     Case 3     ' enviar traslados
          EnviarD.Label1.Caption = "ENVIAR TRASLADOS A:"
          EnviarD.Show 1
     Case 4     ' enviar compras
          EnviarD.Label1.Caption = "ENVIAR COMPRAS"
          EnviarD.Show 1
     Case 5     ' enviar precios
     Case 6     ' enviar descuentos
End Select
End Sub

Private Sub Frecuen_Click(Index As Integer)
     IngCli = 1
     Clientes.Show
End Sub

Private Sub Frecuen2_Click(Index As Integer)
     Profesion.Show
End Sub

Private Sub Frecuen3_Click(Index As Integer)
     Preferen.Show
End Sub

Private Sub Frecuen4_Click(Index As Integer)
     Frecuente.Show
End Sub

Private Sub Frecuen5_Click(Index As Integer)
     encuesta.Show
End Sub

Private Sub Frecuen6_Click(Index As Integer)
    EncClassic.Show
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then
          PopupMenu Menu(1), , x, y, SubMenu2(2)
     End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Trim(LoGiN) = "admin" Then
          If Button = 2 Then GeneraMenus.Show
     End If
End Sub
Private Sub ReciBir_Click(Index As Integer)
RecibirD.Tag = Index
RecibirD.TBodega = LaBodega
RecibirD.DBodega.Caption = LaDBodega
Select Case Index
     Case 1     ' recibir cierre de caja
          RecibirD.Label1.Caption = "RECIBIR CIERRE DE CAJA"
          RecibirD.TBodega = ""
          RecibirD.DBodega.Caption = "Tienda a Traer ?"
          RecibirD.Show 1
     Case 3     ' recibir traslados
          RecibirD.TBodega.Enabled = False
          RecibirD.Label1.Caption = "RECIBIR TRASLADOS"
          RecibirD.Show 1
     Case 4     ' recibir compras
          RecibirD.TBodega.Enabled = False
          RecibirD.Label1.Caption = "RECIBIR COMPRAS"
          RecibirD.Show 1
     Case 5     ' recibir precios
     Case 6     ' recibir descuentos
End Select
End Sub

Private Sub SubMenu2_Click(Index As Integer)
     Select Case Index
     Case 0
          For Each Form In Forms
               If LCase(Form.Caption) = "entradas por compra" Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Set forma = New Entradas
          Call forma.Carga(Index)
     Case 1
          For Each Form In Forms
               If LCase(Form.Caption) = "entradas por ajuste" Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Set forma = New Entradas
          Call forma.Carga(Index)
     Case 2
          'call SetWindowPos(Facturas.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
          FactProd.Show
     Case 9
          Call SetWindowPos(Proformas.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
     Case 3
          Salidas.Show
     Case 5
          Traslados.Show
     Case 7
          For Each Form In Forms
               If LCase(Form.Caption) = "notas de crédito" Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Set forma = New NotasDeb
          Call forma.Init(Index)
     Case 8
          For Each Form In Forms
               If LCase(Form.Caption) = "notas de débito" Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Set forma = New NotasDeb
          Call forma.Init(Index)
     Case 10
          Devolucion.Show
     Case 12
          Ordenes.Show
     Case 13
          Punto.Show
     Case 14
          Monitoreo.Show
     Case 15
          CierreM.Show
     End Select
End Sub
Private Sub mdiForm_load()
     Caption = Caption + ReadKey("nomcia")
     Call Posiciona(Me, 0)
     If CiA = "" Then
          LaBodega = "01"
          For I = 1 To Menu.Count - 1
               Menu(I).Enabled = False
          Next
          On Error Resume Next
          For I = 0 To Submenu1.Count - 1
               Select Case I
               Case 2, 14, 17
               Case Else
                   Submenu1(I).Enabled = False
               End Select
          Next
          On Error GoTo 0
          Seleccion.Show
     Else
          S = "select c_bodega,d_bodega From Bodegas where default=1 and cia='" + CiA + "'"
          Set Temp = DatOS.OpenRecordset(S)
          If Temp.EOF Then
               LaBodega = "01"
               LaDBodega = "SinDefinir bodega PRINCIPAL !!"
          Else
               LaBodega = Temp!c_bodega
               LaDBodega = Trim(Temp!d_bodega)
          End If
          Temp.Close
          Call ValidaCambio
     End If
  
End Sub
Private Sub MDIForm_Resize()
     On Error Resume Next
     Centra Parche
     Parche.Show
     On Error GoTo 0
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
     Unload Lista
     If Not Status Then
          'Set SisTemA = Nothing
          Set DatOS = Nothing
          End
     End If
End Sub
Private Sub submenu1_Click(Index As Integer)
     Select Case Index
     Case 0
          Vendedores.Show
     Case 1
          Articulos.Show
     Case 2
          Bodegas.Show
     Case 3
          Companias.Show
     Case 4, 5
          Estaciones.Show
     Case 6
          TipoPago.Show
     Case 7
          TipoArt.Show
     Case 8
          Provedores.Show
     Case 9
          SubGrupo.Show
     Case 10
          tipos_de_proveedor.Show
     Case 11
          Coleccion.Show
     Case 12
          Usos.Show
     Case 13
          TipoPrecios.Show
     Case 14
          Estaciones.Show
     Case 15
          Talla.Show
     Case 17
          Seleccion.Show
     Case 18
          Parametros.Show
     Case 20
          Unload Me
     Case 22             ' cambiar de usuario
          For Each Form In Forms
               If Form.Name <> "Inicio" And Form.Name <> "Parche" Then
                    Unload Form
               End If
          Next
          Status = True
          Unload Me
          Entrada.Show
     End Select
End Sub
Private Sub LlamaReporte(Llave As String)
     Select Case LCase(Llave)
     Case "ventart", "ventcli", "comisi", "back", "acumventas"
          For Each Form In Forms
               If Form.Tag = Llave Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Dim Re As New Reportes
          Re.Tag = Llave
          Call Re.Carga(Llave)
     Case "novend", "factfec", "notcre", "movimientos", "rotacion"
          For Each Form In Forms
               If Form.Tag = Llave Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Dim Re1 As New RepNoVendidos
          Re1.Tag = Llave
          Call Re1.Carga(Llave)
     Case "entfec"
          RepEntFec.Show
     Case "trasl"
          RepTrasFec.Show
     Case "salfec", "saluso"
          For Each Form In Forms
               If Form.Tag = Llave Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Dim Panta As New RepSalFec
          Panta.Tag = Llave
          Call Panta.Carga(Llave)
     Case "estadis"
          RepEstad.Show
     Case "cierre"
          RepCierre.Show
     Case "agentes"
          RepAgentes.Show
     Case "artuso"
          RepArtUso.Show
     Case "exis", "precios", "minimo"
          For Each Form In Forms
               If Form.Tag = Llave Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Dim Panta2 As New RepPrecios
          Panta2.Tag = Llave
          Call Panta2.Carga(Llave)
     Case "costos"
          Call RepNoVendidos.Carga("costos")
     Case "apartados"
          RepApar.Show
     Case Else
          MsgBox "Reporte no definido", 16, Llave
     End Select
End Sub

Private Sub SubMenu4_Click(Index As Integer)
     Call LlamaReporte(subMenu4(Index).Tag)
End Sub

Private Sub SubMenu5_Click(Index As Integer)
     If Index <> 7 Then
          Call LlamaReporte(Submenu5(Index).Tag)
     Else
          RepSalOrd.Show
     End If
End Sub
Private Sub SubMenu6_Click(Index As Integer)
     If Index = 3 Then
        RepPedidos.Show
     ElseIf Index < 8 And Index <> 4 Then
          Call LlamaReporte(Submenu6(Index).Tag)
     ElseIf Index = 8 Then
          RepLista.Show
     ElseIf Index = 4 Then
        'RepExSub.Show
        Dim Forma2 As New RepExSub
        Call Forma2.Carga("exis")
      ElseIf Index = 9 Then
          For Each Form In Forms
               If Form.Name = "Reporte de Precios" Then
                    Form.ZOrder
                    Exit Sub
               End If
          Next
          Dim forma As New RepPrecios
          Call forma.Carga("precios")
     ElseIf Index = 11 Then
          LotExis.Show
     End If
End Sub

Private Sub Submenu7_Click(Index As Integer)
Dim Temp As Recordset
Select Case Index
Case 0           ' Procesar Ventas entre fechas carga MATRIZ
     repmatriz.Show
     repmatriz.Tag = 0
     repmatriz.Label4.Caption = "PROCESAR LAS VENTAS"
     repmatriz.Command1.Caption = "Procesar"
     repmatriz.Frame2.Visible = False
     repmatriz.Frame3.Visible = False
     repmatriz.Frame4.Visible = False
     repmatriz.Frame5.Visible = False
     repmatriz.Combo3.Visible = False
     repmatriz.Label6.Visible = False
Case 1        ' VENTAS ENTRE FECHAS por unidades
     S = "SELECT c_articulo,ccolor FROM MATRIZ WHERE c_bodega='AA' and cia='" + CiA + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Temp.EOF Then
        MsgBox "No se han procesado las ventas", vbCritical, "Reportes de Matriz"
     Else
        repmatriz.DTPicker1 = Temp!c_articulo
        repmatriz.DTPicker2 = Temp!ccolor
        repmatriz.DTPicker1.Enabled = False
        repmatriz.DTPicker2.Enabled = False
        repmatriz.Tag = 1
        repmatriz.Label4.Caption = "VENTAS EN UNIDADES"
        repmatriz.Show
     End If
Case 2
     repmatriz.Tag = 2
     repmatriz.Label4.Caption = "EXISTENCIAS POR BODEGA"
     repmatriz.Frame1.Visible = False
     repmatriz.Show
Case 3        ' VENTAS ENTRE FECHAS por montos   colones
     
     S = "SELECT c_articulo,ccolor FROM MATRIZ WHERE c_bodega='AA' and cia='" + CiA + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Temp.EOF Then
        MsgBox "No se han procesado las ventas", vbCritical, "Reportes de Matriz"
     Else
        repmatriz.DTPicker1 = Temp!c_articulo
        repmatriz.DTPicker2 = Temp!ccolor
        repmatriz.DTPicker1.Enabled = False
        repmatriz.DTPicker2.Enabled = False
        repmatriz.Tag = 3
        repmatriz.Label4.Caption = "VENTAS X MONTOS"
        repmatriz.Show
     End If
Case 4        ' resumen de ventas por articulos
     S = "SELECT c_articulo,ccolor FROM MATRIZ WHERE c_bodega='AA' and cia='" + CiA + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Temp.EOF Then
        MsgBox "No se han procesado las ventas", vbCritical, "Reportes de Matriz"
     Else
        repmatriz.Frame6.Visible = True
        repmatriz.DTPicker1 = Temp!c_articulo
        repmatriz.DTPicker2 = Temp!ccolor
        repmatriz.DTPicker1.Enabled = False
        repmatriz.DTPicker2.Enabled = False
        Set Temp = DatOS.OpenRecordset("Select Tcambio from parametros where cia='" + CiA + "'")
        If Not Temp.EOF Then
           repmatriz.Text(0) = Temp!TCambio
        End If
        repmatriz.Tag = 4
        repmatriz.Label4.Caption = "Resumen de Ventas"
        repmatriz.Show
     End If
End Select
End Sub
Private Sub SubPop_Click(Index As Integer)
     If Actual.Name <> "Niveles" Then
          Dim Botones As Buttons
          Set Botones = Actual.Toolbar1.Buttons
          Call Actual.Toolbar1_ButtonClick(Botones(Index + 1))
     Else
          Actual.TreeView1.SelectedItem.Image = IIf(Actual.TreeView1.SelectedItem.Image = 1, 2, 1)
     End If
End Sub
Private Sub Timer1_Timer()
     StatusBar1.Panels(1) = ""
     StatusBar1.Panels(1).Picture = StatusBar1.Panels(2).Picture
     Timer1.Enabled = False
End Sub
Private Sub Timer3_Timer()
     If Not Timer4.Enabled Then
          If Marquee.Left + Marquee.Width < Picture1.Left Then
               If Timer4.Interval > 0 Then Timer4.Enabled = True
               Marquee.Left = (Picture1.Width + Picture1.Left) - 50
          Else
               Marquee.Left = Marquee.Left - 75
          End If
     End If
End Sub
Private Sub Timer4_Timer()
     Timer4.Enabled = False
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
     Select Case Button.Key
     Case "espacio1"
          GeneraMenus.Show
     Case "submenu11"
          Estaciones.Show
     Case "sale"
          Unload Me
     Case "submenu10"
          Articulos.Show
     Case "16"
          Parametros.Show
     Case "submenu20"
          Call SubMenu2_Click(0)
     Case "varios12"
          Monitoreo.Show
     Case "submenu31"
          Documentos.Show
     Case "submenu30"
          If PreCBoD = 1 Then
               Precios.Show
          Else
               ListaPrecios.Show
          End If
     Case "submenu22"
          'Call SetWindowPos(Facturas.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
          FactProd.Show
     Case "submenu21"
          Call SubMenu2_Click(1)
     Case "submenu213"
          Punto.Show
     Case "submenu27"
          For Each Form In Forms
               If LCase(Form.Caption) = "notas de crédito" Then
                    Call SetWindowPos(Form.hwnd, 0, 0, 0, 0, 0, &H1 + &H2)
                    Exit Sub
               End If
          Next
          Call forma.Init(7)
     Case "submenu115"
          Parametros.Show
     End Select
End Sub
Private Sub SubMEnu3_Click(Index As Integer)
     Select Case Index
     Case 0
          If PreCBoD = 1 Then
               Precios.Show
          Else
               ListaPrecios.Show
          End If
     Case 1
          Documentos.Show
     Case 2
          Existencias.Show 1
     Case 3
          TiposCambio.Show
     Case 5, 6
          Call Estadisticas.Init(Index)
          Estadisticas.Show
     Case 7
          Tomas.Show
     Case 8
          CostoInv.Show
     Case 9
          Movimientos.Show
     Case 11
          Etiquetas.Show
     Case 14
          Usuarios.Show
     Case 15
          Niveles.Show
     Case 16
          Bitacora.Show
     Case 18
          ImportArti.Show
     Case 19
          Define.Show
     Case 21
          Traspaso.Show
     Case 22
          Contables.Show
     Case 24
          Respaldos.Show
     Case 25
          Acerca.Show 1
     End Select
End Sub
Private Sub Timer2_Timer()
     Static Imagen%
     Imagen = Imagen + 1
     StatusBar1.Panels(1).Picture = ImageList2.ListImages(Imagen).Picture
     If Imagen >= ImageList2.ListImages.Count Then Imagen = 0
     w% = DoEvents
End Sub
Private Sub Varios1_Click(Index As Integer)
Select Case Index
Case 0            ' Cambio de precios
     CambiaPrecios.Show
Case 1
     Articulo1.Show
Case 2
     Apartados.Show
Case 3
     PagaApa.Show
Case 4
     CierreCaja.Show
Case 5
     Certificados.Show
Case 6
     BolAmeCla.Show 1
Case 11
     Congelados.Show
Case 12
     Barras.Show
End Select
End Sub

Private Sub Ventanas_Click(Index As Integer)
     Arrange Index
End Sub
Public Function ValidaCambio() As Boolean
     On Error GoTo Errores
     If TipoVerS = "00" Then
          StatusBar1.Panels(4).Visible = False
          Exit Function
     Else
          Dim Tabla As Recordset
          Dim S$
          S = "select * from tipocambio where fecha=#" + Format(Date, "m/d/yyyy") + "#"
          Set Tabla = DatOS.OpenRecordset(S)
          If Tabla.EOF Then
               MsgBox "Debe actualizar el tipo de cambio de hoy !", 64, ConVierte(Date, "F")
               Call TiposCambio.Carga(Date)
          Else
               TCambio = Tabla!TipoCambio
               StatusBar1.Panels(4) = Format(TCambio, "standard") + " "
               StatusBar1.Panels(4).ToolTipText = "Tipo de cambio del día : " + Format(TCambio, "standard")
          End If
          Set Tabla = Nothing
     End If
Errores:
     If err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "ValidaCambio"
     End If
     On Error GoTo 0
End Function
 
