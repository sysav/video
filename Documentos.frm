VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Documentos 
   Caption         =   "Documentos"
   ClientHeight    =   2880
   ClientLeft      =   1575
   ClientTop       =   1710
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Documentos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   10185
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   6150
      TabIndex        =   6
      Top             =   -15
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "busca"
            Object.ToolTipText     =   "Buscar un documento"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "para"
            Object.ToolTipText     =   "Detener el cargado de la lista"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "anula"
            Object.ToolTipText     =   "Anular el documento selecionado"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "autom"
            Object.ToolTipText     =   "Cargar la lista por omisión"
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "imprime"
            Object.ToolTipText     =   "Reimprimir el documento seleccionado"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "grandes"
            Object.ToolTipText     =   "Iconos grandes"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "small"
            Object.ToolTipText     =   "Iconos pequeños"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "lista"
            Object.ToolTipText     =   "Lista"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "reporte"
            Object.ToolTipText     =   "Reporte"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sale"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar Bar1 
      Height          =   285
      Left            =   3570
      TabIndex        =   9
      Top             =   2550
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   2280
      Top             =   1710
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   3
      Top             =   390
      Width           =   7815
   End
   Begin VB.PictureBox Splitter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   3120
      MouseIcon       =   "Documentos.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   1320
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   990
      Width           =   510
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1515
      Left            =   5160
      TabIndex        =   2
      Top             =   990
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   2672
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList3"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   1395
      Left            =   240
      TabIndex        =   1
      Top             =   810
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2461
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
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
   Begin ComctlLib.ListView ListView2 
      Height          =   1515
      Left            =   3630
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   2672
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList3"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   8
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
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "unidades"
         Object.Width           =   1411
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
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   2610
      TabIndex        =   10
      Top             =   45
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   37939
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   4515
      TabIndex        =   11
      Top             =   30
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   37939
   End
   Begin VB.Label Label3 
      Caption         =   "Al "
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4200
      TabIndex        =   13
      Top             =   105
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Del "
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2205
      TabIndex        =   12
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   60
      TabIndex        =   8
      Top             =   90
      Width           =   45
   End
   Begin VB.Label Etiqueta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   2340
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   2820
      TabIndex        =   4
      Top             =   2700
      Visible         =   0   'False
      Width           =   45
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   1500
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":0C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":0F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":127C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":1596
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":18B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":1BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":1EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":21FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":2518
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":2832
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":2B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":2E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":3180
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":37B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":3ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":3DE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   780
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":4102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":441C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":4736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":4A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":4D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":5084
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":539E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":56B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":59D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":5CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":6006
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":6320
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":663A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":6954
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":6C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":6F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":72A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":75BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":78D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":7BF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":7F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":8224
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":8336
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":8448
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":855A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":866C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":877E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":8A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":8DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":8EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":91DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":9360
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Documentos.frx":967A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempRec As Recordset
Dim S$
Private Const P_ECART = 30 'El ancho del separador
Private Y1 As Integer
Private Y2 As Integer
Private X1 As Integer
Private X2 As Integer
Private Width1 As Integer
Private Width2 As Integer
Private Height1 As Integer
Private Height2 As Integer
Private GlbfrmInSizeX As Long
Dim Nodo As Node
Dim Item As ListItem
Dim Campo$
Dim Itemo%

Private Sub Form_Activate()
     Call Menus(True, Me)
End Sub
Private Sub Form_Deactivate()
     Call Menus(False, Me)
End Sub
Private Sub Form_Load()
     Call Posiciona(Me, 0)
     On Error Resume Next
     TreeView1.Width = CSng(ReadKey("ArbolTreeWidth", "xy"))
     Toolbar1.Buttons("autom").Value = Val(ReadKey("Automatico"))
     On Error GoTo 0
     Y1 = Frame1.Top + Frame1.Height + 80
     Show
     DTPicker1 = Date
     DTPicker2 = Date
     Refresh
     GlbfrmInSizeX = &H7FFFFFFF
     Form_Resize
     Call Carga
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     If MousePointer = 11 Then
          Cancel = True
     End If
End Sub
Private Sub Form_Resize()
     On Error Resume Next
     Const B_ECART = 1
     Height1 = ScaleHeight - Y1 - B_ECART * 2
     X1 = B_ECART
     Width1 = TreeView1.Width
     X2 = X1 + TreeView1.Width + P_ECART - 1
     Width2 = ScaleWidth - X2 - B_ECART
     TreeView1.Move X1 - 1, Y1, Width1, Height1
     ListView1.Move X2, Y1, Width2 + 1, Height1
     Splitter.Move X1 + TreeView1.Width - 1, Y1, P_ECART, Height1
     Frame1.Width = Width - 100
     Toolbar1.Left = (Width - Toolbar1.Width) - 100
     Label1.Left = ListView1.Left
     Bar1.Left = ListView1.Left + 25
     Bar1.Width = ListView1.Width - 60
     Bar1.Top = Height - Bar1.Height - 420
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Menus(False, Me)
     Call Posiciona(Me, 1)
     Call SaveKey("ArbolTreeWidth", TreeView1.Width, "XY")
     Call SaveKey("Automatico", Toolbar1.Buttons("autom").Value)
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
     MousePointer = 11
     w% = DoEvents
     ListView1.SortOrder = IIf(ListView1.SortOrder = 0, 1, 0)
     ListView1.SortKey = ColumnHeader.Index - 1
     ListView1.Sorted = True
     MousePointer = 0
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     If MousePointer = 11 Then Exit Sub
     On Error Resume Next
     With Toolbar1
     Select Case Item.SmallIcon
     Case 7
          .Buttons("anula").Enabled = True
          .Buttons("imprime").Enabled = False
     Case 9, 13, 2, 3, 17
          .Buttons("anula").Enabled = True
          .Buttons("imprime").Enabled = True
     Case 10, 14, 18, 19
          .Buttons("anula").Enabled = False
          .Buttons("imprime").Enabled = True
     Case Else
          .Buttons("anula").Enabled = False
          .Buttons("imprime").Enabled = False
     End Select
     Inicio.SubPop(2).Enabled = .Buttons("anula").Enabled
     Inicio.SubPop(.Buttons("imprime").Index - 1).Enabled = .Buttons("imprime").Enabled
     End With
     On Error GoTo 0
End Sub
Private Sub Listview1_DblClick()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then
          Set Item = ListView1.SelectedItem
          Select Case Item.SmallIcon
          Case 2, 3 'Entradas
               S = "select desgent.c_articulo as Código,"
               S = S + "d_articulo as Descripción,Cantidad,"
               S = S + "desgent.p_compra as Costo,total_neto as Total "
               S = S + "from desgent left join articulos "
               S = S + "on articulos.cia=desgent.cia "
               S = S + "and articulos.c_articulo=desgent.c_articulo "
               S = S + "where desgent.cia='" + CiA + "' and n_entrada='" + Item.Text
               S = S + "' and bodega='" + Item.SubItems(2)
               S = S + "' order by d_articulo"
               Dim Pr As New DetalleFactura
               Call Pr.CargaLista(S, "Entrada No. " + Item.Text, Item.SmallIcon)
          Case 7 'Facturas
               S = "select desgfact.c_articulo as Código,"
               S = S + "d_articulo as Descripción,Cantidad,Unidades,Precio,"
               S = S + "total_brut as Subtotal,porc_imp as IV,"
               S = S + "Descuento,total_neto as Total "
               S = S + "from desgfact left join articulos "
               S = S + "on articulos.cia=desgfact.cia "
               S = S + "and articulos.c_articulo=desgfact.c_articulo "
               S = S + "where desgfact.cia='" + CiA
               S = S + "' and c_bodega='" + Item.SubItems(6) + "' and n_factura='" + Item.Text
               S = S + "'order by Desgfact.c_articulo"
               Dim Fr As New DetalleFactura
               Call Fr.CargaLista(S, "Factura No. " + Item.Text, Item.SmallIcon)
          Case 10 'Notas de credito
               S = "select desgcred.c_articulo as Código,"
               S = S + "d_articulo as Descripción,Cantidad,Unidades,Precio,"
               S = S + "total_brut as Subtotal,porc_imp as IV,"
               S = S + "Descuento,total_neto as Total "
               S = S + "from desgcred left join articulos "
               S = S + "on articulos.cia=desgcred.cia "
               S = S + "and articulos.c_articulo=desgcred.c_articulo "
               S = S + "where desgcred.cia='" + CiA + "' and n_factura='" + Item.Text
               S = S + "'order by d_articulo"
               Dim NC As New DetalleFactura
               Call NC.CargaLista(S, "Nota de Crédito No. " + Item.Text, Item.SmallIcon)
          Case 14 'Notas de debito
               S = "select desgdeb.c_articulo as Código,"
               S = S + "d_articulo as Descripción,Cantidad,Unidades,Precio,"
               S = S + "total_brut as Subtotal,porc_imp as IV,"
               S = S + "Descuento,total_neto as Total "
               S = S + "from desgdeb left join articulos "
               S = S + "on articulos.cia=desgdeb.cia "
               S = S + "and articulos.c_articulo=desgdeb.c_articulo "
               S = S + "where desgdeb.cia='" + CiA + "' and n_factura='" + Item.Text
               S = S + "'order by d_articulo"
               Dim ND As New DetalleFactura
               Call ND.CargaLista(S, "Nota de Débito No. " + Item.Text, Item.SmallIcon)
          Case 9 'Requisiciones
               S = "select desgsali.c_articulo as Código,"
               S = S + "d_articulo as Descripción,Cantidad,Unidades,"
               S = S + "desgsali.p_compra as Costo,total_neto as Total "
               S = S + "from desgsali left join articulos "
               S = S + "on articulos.cia=desgsali.cia "
               S = S + "and articulos.c_articulo=desgsali.c_articulo "
               S = S + "where desgsali.cia='" + CiA + "' and bodega='" + ListView1.SelectedItem.SubItems(1)
               S = S + "' and n_salida='" + Item.Text
               S = S + "'order by d_articulo"
               Dim Fre As New DetalleFactura
               Call Fre.CargaLista(S, "Requisición No. " + Item.Text, Item.SmallIcon)
          Case 13 'Proformas
               S = "select desgprof.c_articulo as Código,"
               S = S + "d_articulo as Descripción,Cantidad,Unidades,Precio,"
               S = S + "total_brut as Subtotal,porc_imp as IV,"
               S = S + "Descuento,total_neto as Total "
               S = S + "from desgprof left join articulos "
               S = S + "on articulos.cia=desgprof.cia "
               S = S + "and articulos.c_articulo=desgprof.c_articulo "
               S = S + "where desgprof.cia='" + CiA + "' and n_factura='" + Item.Text
               S = S + "'order by d_articulo"
               Dim Frf As New DetalleFactura
               Call Frf.CargaLista(S, "Proforma No. " + Item.Text, Item.SmallIcon)
          Case 17 'Ordenes de compra
               S = "select codart as Código,d_articulo as Descripción,"
               S = S + "Cantidad,Precio,Subtotal,Descuento,"
               S = S + "Impuesto,Total "
               S = S + "from desgord left join articulos "
               S = S + "on articulos.cia=desgord.cia "
               S = S + "and articulos.c_articulo=desgord.codart "
               S = S + "where desgord.cia='" + CiA
               S = S + "' and orden='" + Item.Text + "' order by d_articulo"
               Dim OrD As New DetalleFactura
               Call OrD.CargaLista(S, "Orden de Compra No. " + Item.Text, Item.SmallIcon)
          Case 18 'Traslados entre bodegas
               S = "select desgtras.c_Articulo as Código,d_articulo as Descripción,"
               S = S + "Cantidad,desgtras.P_compra as Costo,total_brut as Total "
               S = S + "from desgtras left join articulos "
               S = S + "on articulos.cia=desgtras.cia "
               S = S + "and articulos.c_articulo=desgtras.c_articulo "
               S = S + "where desgtras.cia='" + CiA
               S = S + "' and n_traslado='" + Item.Text + "' and c_bodega='" + Item.SubItems(1)
               S = S + "' order by d_articulo"
               Dim Tras As New DetalleFactura
               Call Tras.CargaLista(S, "Traslado entre Bodegas No. " + Item.Text, Item.SmallIcon)
          Case 19 'Devoluciones sobre compras
               S = "select desgdev.c_Articulo as Código,d_articulo as Descripción,"
               S = S + "Cantidad,desgdev.costo as Costo,total_brut as Total "
               S = S + "from desgdev left join articulos "
               S = S + "on articulos.cia=desgdev.cia "
               S = S + "and articulos.c_articulo=desgdev.c_articulo "
               S = S + "where desgdev.cia='" + CiA
               S = S + "' and n_devol='" + Item.Text + "' order by d_articulo"
               Dim dev As New DetalleFactura
               Call dev.CargaLista(S, "Devolución Sobre Compras No. " + Item.Text, Item.SmallIcon)
          End Select
     End If
End Sub
Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected And KeyCode = vbKeyDelete Then
          Call Toolbar1_ButtonClick(Toolbar1.Buttons("anula"))
     End If
End Sub
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then
          For I = 0 To Inicio.SubPop.Count - 1
               If Inicio.SubPop(I).Enabled Then Exit For
          Next
          PopupMenu Inicio.Popup, , x, y, Inicio.SubPop(I)
     End If
End Sub

Private Sub splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlbfrmInSizeX <> &H7FFFFFFF Then
        If CLng(x) <> GlbfrmInSizeX Then
            Splitter.Move Splitter.Left + x, Y1, P_ECART, ScaleHeight - Y1 - 2
            GlbfrmInSizeX = CLng(x)
        End If
    End If
End Sub
Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlbfrmInSizeX <> &H7FFFFFFF Then
        If CLng(x) <> GlbfrmInSizeX Then
            Splitter.Move Splitter.Left + x, Y1, P_ECART, ScaleHeight - Y1 - 2
        End If
        GlbfrmInSizeX = &H7FFFFFFF
        Splitter.BackColor = &H8000000F
        If Splitter.Left > 60 And Splitter.Left < (ScaleWidth - 60) Then
            TreeView1.Width = Splitter.Left - TreeView1.Left
        ElseIf Splitter.Left < 60 Then
            TreeView1.Width = 60
        Else
            TreeView1.Width = ScaleWidth - 60
        End If
        Form_Resize
    End If
End Sub
Private Sub splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Splitter.BackColor = &H808080
        GlbfrmInSizeX = CLng(x)
    Else
        If GlbfrmInSizeX <> &H7FFFFFFF Then
            splitter_MouseUp Button, Shift, x, y
        End If
        GlbfrmInSizeX = &H7FFFFFFF
    End If
End Sub
Private Sub Carga()
     Set Nodo = TreeView1.Nodes.Add(, , , "Entradas por compras", 2)
     Set Nodo = TreeView1.Nodes.Add(, , , "Entradas por ajuste", 3)
     Set Nodo = TreeView1.Nodes.Add(, , , "Facturas", 7)
     Set Nodo = TreeView1.Nodes.Add(, , , "Requisiciones", 9)
     Set Nodo = TreeView1.Nodes.Add(, , , "Facturas Proformas", 13)
     Set Nodo = TreeView1.Nodes.Add(, , , "Notas de Crédito", 10)
     Set Nodo = TreeView1.Nodes.Add(, , , "Notas de Débito", 14)
     Set Nodo = TreeView1.Nodes.Add(, , , "Ordenes de Compra", 17)
     Set Nodo = TreeView1.Nodes.Add(, , , "Traslados entre Bodegas", 18)
     Set Nodo = TreeView1.Nodes.Add(, , , "Devoluciones de Compra", 19)
     TreeView1.Visible = True
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
     Select Case Button.Key
     Case "grandes"
          ListView1.View = lvwIcon
     Case "small"
          ListView1.View = lvwSmallIcon
     Case "lista"
          ListView1.View = lvwList
     Case "reporte"
          ListView1.View = lvwReport
     Case "sale"
          Unload Me
     Case "busca"
          Dim NomTabla$
          Select Case TreeView1.SelectedItem.Image
          Case 2 'Entradas por compras
               NomTabla = "entradas"
          Case 3 'Entradas por ajuste
               NomTabla = "entradas"
          Case 7 'Facturas
               NomTabla = "facturas"
          Case 9 'Salidas por ajuste
               NomTabla = "salidas"
          Case 13 'Facturas Proformas
               NomTabla = "proformas"
          Case 19 'Devoluciones sobre Compras
               NomTabla = "devolucion"
          'case 10 Precios y Existencias
          'case 14 Existencias por Bodega
          End Select
          Call Busquedas.Carga(TreeView1.SelectedItem)
          Busquedas.Show 1
          TreeView1.SetFocus
     Case "anula"
          MousePointer = 11
          EspaCio.BeginTrans
          If Anula(ListView1.SelectedItem) Then
               EspaCio.CommitTrans
          Else
               EspaCio.Rollback
          End If
          MousePointer = 0
     Case "imprime"
          MousePointer = 11
          Select Case TreeView1.SelectedItem.Image
          Case 9 'Requisiciones
               bode = ListView1.SelectedItem.SubItems(2)
               Call ImprimeSalida(ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.Text, Report1)
          Case 10 'Notas de Credito
               Call ImprimeNotas(ListView1.SelectedItem.Text, Report1)
          Case 14 'Notas de Debito
               Call ImpreNotasDeb(ListView1.SelectedItem.Text, Report1)
          Case 17 'Ordenes de Compra
               Call ImprimeOrdenes(ListView1.SelectedItem, Report1)
          Case 2 'Imprime Entradas de Compras
              DescAdi.Show 1 'Descuento Adicional
              bode = ListView1.SelectedItem.SubItems(2)
              Call ImprimeEntrada1(ListView1.SelectedItem, Report1)
          Case 3 'Imprime Entradas de Ajuste
               bode = ListView1.SelectedItem.SubItems(2)
               monImp = 1
               DesAdi = 0
               Call ImprimeEntrada1(ListView1.SelectedItem, Report1)
          Case 18 'Imprime Traslados
               bode = ListView1.SelectedItem.SubItems(1)
               bode2 = ListView1.SelectedItem.SubItems(4)
               Call Imprimetraslado(ListView1.SelectedItem, bode, bode2, Report1)
          Case 19 'Imprime Devoluciones sobre compras
               bode = ListView1.SelectedItem.SubItems(2)
               Call Imprimedevolucion(ListView1.SelectedItem, Report1)
          Case 13 'Facturas Proformas
               Report1.DataFiles(0) = DatOS.Name
               Report1.DataFiles(1) = DatOS.Name
               Report1.DataFiles(2) = DatOS.Name
               Report1.DataFiles(3) = RutaCXC
               If App.Comments = "Dafesa" Then
                    Report1.ReportFileName = DirTrA + "Reportes\profdafe.rpt"
               Else
                    Report1.ReportFileName = DirTrA + "Reportes\profhov.rpt"
               End If
               S = "{proformas.cia}='" + CiA
               S = S + "' and {proformas.n_factura}='" + ListView1.SelectedItem.Text + "'"
               Report1.SelectionFormula = S
               Report1.WindowTitle = "Proforma No. " + ListView1.SelectedItem
               Report1.Action = 1
          End Select
          MousePointer = 0
     End Select
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
     If MousePointer = 11 Then Exit Sub
     Toolbar1.Buttons("anula").Enabled = False
     Toolbar1.Buttons("imprime").Enabled = False
     MousePointer = 11
     w% = DoEvents
     Select Case Node.Image
     Case 2, 3 'Despliega las Entradas
          Label2.Caption = "Entradas " + IIf(Node.Image = 3, "por ajuste", "por compra")
          S = "select n_entrada as Número,FACT_COMP as Factura,c_bodega as Bodega,f_entrada as Fecha,"
          S = S + "monto_real as Total,nombre as Proveedor "
          S = S + "from entradas left join proveedores "
          S = S + "on proveedores.compania=entradas.cia "
          S = S + "and proveedores.codigo=entradas.c_prove "
          S = S + " where entradas.tipo=" + Trim(Node.Image - 2)
          S = S + " and entradas.cia='" + CiA + "'"
          S = S + " and f_entrada>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_entrada<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by f_entrada,n_entrada"
          Etiqueta.Caption = "Número de Entrada :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 0)
          End If
     Case 7 'Despliega las Facturas
          Label2.Caption = "Facturas"
          S = "select N_factura as Número,f_factura as Fecha,"
          S = S + "Financiamiento as Financ,monto_real as Total,Nombre as Cliente,"
          S = S + "iif(Tipofact=0,'CO',iif(tipofact=1,'CR','AP')) as Tipo,c_bodega as Bodega,estado"
          S = S + " from facturas left join clientes "
          S = S + "on clientes.cia=facturas.cia "
          S = S + "and clientes.codigo=facturas.c_cliente "
          S = S + "where facturas.cia='" + CiA + "'"
          S = S + " and f_factura>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_factura<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by facturas.cia,c_bodega,n_factura"
          Etiqueta.Caption = "Número de Factura :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 4)
          End If
     Case 10 'Notas de credito
          Label2.Caption = "Notas de Crédito"
          S = "select N_factura as Número,f_factura as Fecha,"
          S = S + "monto_real as Total,Nombre as Cliente,"
          S = S + "estado from notascred left join clientes "
          S = S + "on clientes.cia=notascred.cia "
          S = S + "and clientes.codigo=notascred.c_cliente "
          S = S + "where notascred.cia='" + CiA + "'"
          S = S + " and f_factura>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_factura<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by f_factura,n_factura"
          Etiqueta.Caption = "Número de Nota de Crédito :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 1)
          End If
     Case 14 'Notas de debito
          Label2.Caption = "Notas de Débito"
          S = "select N_factura as Número,f_factura as Fecha,"
          S = S + "monto_real as Total,Nombre as Cliente,"
          S = S + "estado from notasdeb left join clientes "
          S = S + "on clientes.cia=notasdeb.cia "
          S = S + "and clientes.codigo=notasdeb.c_cliente "
          S = S + "where notasdeb.cia='" + CiA + "'"
          S = S + " and f_factura>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_factura<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by f_factura,n_factura"
          Etiqueta.Caption = "Número de Nota de Débito :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 1)
          End If
     Case 9 'Salidas
          Label2.Caption = "Requisiciones"
          S = "select N_salida as Número,c_bodega as Bodega,f_salida as Fecha,"
          S = S + "monto as SubTotal,monto_real as Total "
          S = S + "from salidas where cia='" + CiA + "'"
          S = S + " and f_salida>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_salida<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by c_bodega,n_salida"
          Etiqueta.Caption = "Número de Requisición :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
                Call FillList(S, Node.Image, 0)
          End If
     Case 13 'Proformas
          Label2.Caption = "Proformas"
          S = "select N_factura as Número,f_factura as Fecha,"
          S = S + "Plazo,Monto,Descuento,Impuesto,monto_real as Total,estado "
          S = S + "from proformas where cia='" + CiA + "'"
          S = S + " and f_factura>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_factura<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by f_factura,n_factura"
          Etiqueta.Caption = "Número de Proforma :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 1)
          End If
     Case 17 'Ordenes de Compra
          Label2.Caption = "Ordenes de Compra"
          S = "select Numero as Número,Proveedor,"
          S = S + "Fecha,Plazo,Total "
          If App.Comments = "3M" Then
               S = S + ",TCambio "
          End If
          S = S + "from ordenes where cia='" + CiA + "'"
          S = S + " and fecha>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and fecha<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by fecha,numero"
          Etiqueta.Caption = "Número de Orden :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 0)
          End If
     Case 18 'Traslados entre bodegas
          Label2.Caption = "Traslados entre Bodegas"
          S = "select N_traslado as Número,c_Bodega as bodega,f_traslado as Fecha,"
          S = S + "c_bodega as Fuente,c_bodedest as Destino,Monto "
          S = S + "from traslados where cia='" + CiA + "'"
          S = S + " and f_traslado>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_traslado<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by f_traslado,n_traslado"
          Etiqueta.Caption = "Número de Traslado :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 0)
          End If
     Case 19 'Despliegue de Devoluciones
          Label2.Caption = "Devoluciones de Compra"
          S = "select N_DEVOL as Número,F_DEVOL as Fecha,Impuesto,Descuento,Monto,Estado "
          S = S + "from devolucion where cia='" + CiA + "'"
          S = S + " and f_devol>=#" + Format(DTPicker1, "m/d/yyyy") + "#"
          S = S + " and f_devol<=#" + Format(DTPicker2, "m/d/yyyy") + "#"
          S = S + " order by f_devol,n_devol"
          Etiqueta.Caption = "Número de Devolución :"
          Toolbar1.Buttons("busca").Enabled = True
          If Toolbar1.Buttons("autom").Value = 1 Then
               Call FillList(S, Node.Image, 1)
          End If
     Case Else
          Label2.Caption = ""
          ListView1.ListItems.Clear
          Toolbar1.Buttons("busca").Enabled = False
          Toolbar1.Buttons("imprime").Enabled = False
          For I = 1 To ListView1.ColumnHeaders.Count
               ListView1.ColumnHeaders(I).Text = Space(1)
          Next
     End Select
     For I = 1 To Toolbar1.Buttons.Count
          Inicio.SubPop(I - 1).Enabled = Toolbar1.Buttons(I).Enabled
     Next
     MousePointer = 0
End Sub
Public Function FillList(SQL$, Imagen%, Viene%) As Boolean
     On Error GoTo Errores
     Toolbar1.Buttons("para").Enabled = True
     ListView2.Height = ListView1.Height
     ListView2.Width = ListView1.Width
     ListView2.Top = ListView1.Top
     ListView2.Left = ListView1.Left
     ListView2.Visible = True
     ListView1.Visible = False
     Call Barra("Un momento ... Cargando registros", 2)
     Dim Columnas As ColumnHeaders
     Set Columnas = ListView1.ColumnHeaders
     ListView1.ListItems.Clear
     Set TempRec = DatOS.OpenRecordset(SQL)
     Bar1.Min = 0
     If Not TempRec.EOF Then
          Bar1.Visible = True
          TempRec.MoveLast
          Bar1.Max = TempRec.RecordCount
          TempRec.MoveFirst
     End If
     Dim CTabla%
     Dim CColumna%
     If Viene Then
          CTabla = TempRec.Fields.Count - 1
     Else
          CTabla = TempRec.Fields.Count
     End If
     CColumna = Columnas.Count
     If CTabla > CColumna Then
          Do Until CColumna = CTabla
               Set Columna = Columnas.Add
               CColumna = Columnas.Count
          Loop
     ElseIf CColumna > CTabla Then
          Do Until CColumna = CTabla
               Columnas.Remove Columnas.Count
               CColumna = Columnas.Count
          Loop
     End If
     For I = 0 To CTabla - 1
          If LCase(TempRec(I).Name) <> "estado" Then
               Columnas(I + 1).Text = TempRec(I).Name
               Label1 = Columnas(I + 1).Text
               If I = 0 Then
                    Columnas(I + 1).Width = Label1.Width + 500
               Else
                    Columnas(I + 1).Width = Label1.Width
               End If
               Select Case TempRec.Fields(I).Properties(3)
               Case dbText
                    Columnas(I + 1).Alignment = lvwColumnLeft
               Case dbLong, dbInteger, dbSingle, dbDouble, dbCurrency
                    Columnas(I + 1).Alignment = lvwColumnRight
               Case dbDate
                    Columnas(I + 1).Alignment = lvwColumnCenter
               Case Else
                    Columnas(I + 1).Alignment = lvwColumnLeft
               End Select
          End If
          w% = DoEvents
     Next
     Dim PrevImag%
     PrevImag = Imagen
     Do Until TempRec.EOF
          If Viene > 0 Then
               Imagen = IIf(TempRec!Estado = 1, 4, PrevImag)
          End If
          If Imagen = 17 Or Imagen = 4 Then 'Carga la Orden en Anulada
               'MsgBox TempRec!Numero
               If TempRec!Total = 0 Then
                    Imagen = 4
               Else
                    Imagen = 17
               End If
          End If
          Label1.Caption = Formato(TempRec(0))
          If Label1.Width > Columnas(1).Width Then
               Columnas(1).Width = Label1.Width
          End If
          Set Item = ListView1.ListItems.Add(, , Label1, Imagen, Imagen)
          For I = 1 To CTabla - 1
               Label1.Caption = Formato(TempRec(I))
               Item.SubItems(I) = Label1.Caption
               If Label1.Width > Columnas(I + 1).Width Then
                    Columnas(I + 1).Width = Label1.Width
               End If
          Next
          w% = DoEvents
          If Toolbar1.Buttons("para").Value = 1 Then
               Toolbar1.Buttons("para").Value = 0
               Exit Do
          Else
               Bar1.Value = Bar1.Value + 1
               TempRec.MoveNext
          End If
     Loop
     Dim Cant&
     Cant = ListView1.ListItems.Count
     S = Cant & " documento"
     If Cant <> 1 Then S = S + "s"
     S = S + " encontrado"
     If Cant <> 1 Then S = S + "s"
     Call Barra(S, 1)
     ListView2.Visible = False
     ListView1.Visible = True
     FillList = True
Errores:
     If err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number) + Chr(13) + SQL, 16, "FillList"
     End If
     On Error GoTo 0
     Call Barra("", 2)
     Bar1.Value = 0
     Bar1.Visible = False
     Toolbar1.Buttons("para").Enabled = False
End Function
Private Function Anula(Item As ListItem) As Boolean
     On Error GoTo Errores
     Dim Temp As Recordset
     Dim CliBod As Recordset
     Select Case Item.SmallIcon
     Case 7 'Facturas
          If MsgBox("Desea anular la factura seleccionada ?", 36, Item.Text) = 6 Then
               S = "select c_cliente,c_bodega,f_factura from facturas "
               S = S + "where cia='" + CiA + "' and c_bodega='" + Item.SubItems(6)
               S = S + "' and n_factura='" + Item.Text + "'"
               Set CliBod = DatOS.OpenRecordset(S)

               S = "select * from desgfact where cia='" + CiA
               S = S + "' and c_bodega='" + Item.SubItems(6) + "' and n_factura='"
               S = S + Item.Text + "'"
               Set Temp = DatOS.OpenRecordset(S)
               Do Until Temp.EOF
                    S = "update existencias set existencia=existencia+" & Temp!Cantidad
                    S = S + " where cia='" + CiA + "' and c_articulo='" + Temp!c_articulo
                    S = S + "' and c_bodega='" + Item.SubItems(6) + "'"
                    If UsaLotes = 1 Then
                         S = S + " and lote='" + Nulo(Temp!Lote) + "'"
                    End If
                    DatOS.Execute S, 128
                    S = "update estadisticas set unidades=unidades-"
                    S = S & Temp!Cantidad * Temp!Unidades
                    S = S + ",monto=monto-" & Temp!total_neto
                    S = S + " where cia='" + CiA + "' and codart='" + Temp!c_articulo
                    S = S + "' and codbod='" + CliBod!c_bodega
                    S = S + "' and cliente='" + CliBod!c_cliente
                    S = S + "' and fecha='" + Format(CliBod!f_factura, "yyyymm") + "'"
                    DatOS.Execute S, 128
                    w% = DoEvents
                    Temp.MoveNext
               Loop
               S = "update facturas set estado=1,"
               S = S + "monto=0,descuento=0,impuesto=0,monto_real=0,"
               S = S + "flete=0,costo=0,pago_1=0,pago_2=0,pago_3=0 "
               S = S + "where cia='" + CiA + "' and c_bodega='" + Item.SubItems(6)
               S = S + "' and n_factura='" + Item.Text + "'"
               DatOS.Execute S, 128
               S = "update desgfact set cantidad=0,total_brut=0,porc_imp=0,"
               S = S + "descuento=0,total_neto=0,precio=0 "
               S = S + "where cia='" + CiA + "' and c_bodega='" + Item.SubItems(6)
               S = S + "' and n_factura='" + Item.Text + "'"
               DatOS.Execute S, 128
               S = "delete from apartados "
               S = S + "where codclie='" + CliBod!c_cliente
               S = S + "' and n_factura='" + Item.Text + "'"
               DatOS.Execute S, 128
               
               Anula = True
               Item.Icon = 4
               Item.SmallIcon = 4
               Set CliBod = Nothing
               Set Temp = Nothing
               Call GBitacora(5, "Factura No. " + Item.Text)
          End If
     Case 13         'ProformAS
          If MsgBox("Desea anular la proforma seleccionada ?", 36, Item.Text) = 6 Then
             S = "delete from proformas "
             S = S + "where cia='" + CiA + "' and n_factura='" + Item.Text + "'"
             DatOS.Execute S, 128
             Anula = True
             Item.SmallIcon = 12
             Call GBitacora(5, "Proforma No. " + Item.Text)
          End If
     Case 2, 3 'Entradas
          If MsgBox("Desea reversar la entrada seleccionada ?", 36, Item.Text) = 6 Then
               S = "select * from desgent where cia='" + CiA
               S = S + "' and bodega='" + Item.SubItems(2)
               S = S + "' and n_entrada='" + Item.Text + "'"
               Set Temp = DatOS.OpenRecordset(S)
               Do Until Temp.EOF
                    S = "update existencias set existencia=existencia-" & Temp!Cantidad
                    S = S + " where cia='" + CiA
                    S = S + "' and c_bodega='" + Item.SubItems(2)
                    S = S + "' and c_articulo='" + Temp!c_articulo + "'"
                    If UsaLotes = 1 Then
                         S = S + " and lote='" + Temp!Lote + "'"
                    End If
                    DatOS.Execute S, 128
                    w% = DoEvents
                    Temp.MoveNext
               Loop
               S = "delete from entradas where cia='" + CiA
               S = S + "' and n_entrada='" + Item.Text
               S = S + "' and c_bodega='" + Item.SubItems(2) + "'"
               DatOS.Execute S, 128
               S = "delete from desgent where cia='" + CiA + "' and n_entrada='" + Item.Text
               S = S + "' and bodega='" + Item.SubItems(2) + "'"
               DatOS.Execute S, 128
               ListView1.ListItems.Remove Item.Index
               Call GBitacora(5, "Entrada No. " + Item.Text)
               Anula = True
               Set Temp = Nothing
          End If
     Case 9 'Salidas
          If MsgBox("Desea reversar la requisición seleccionada ?", 36, Item.Text) = 6 Then
               S = "select cantidad,lote,c_articulo "
               S = S + "from desgsali where n_salida='" + Item.Text
               S = S + "' and bodega='" & Item.SubItems(1) & "'"
               Set Temp = DatOS.OpenRecordset(S)
               Do Until Temp.EOF
                    S = "update existencias set existencia=existencia+"
                    S = S & Temp!Cantidad
                    S = S + " where c_articulo='" + Temp!c_articulo
                    S = S + "' and c_bodega='" + Item.SubItems(1) + "'"
                    If UsaLotes = 1 Then
                         S = S + " and lote='" + Temp!Lote + "'"
                    End If
                    DatOS.Execute S, 128
                    w% = DoEvents
                    Temp.MoveNext
               Loop
               S = "delete from salidas where n_salida='" + Item.Text
               S = S + "' and c_bodega='" + Item.SubItems(1) + "'"
               DatOS.Execute S, 128
               S = "delete from desgsali where n_salida='" + Item.Text
               S = S + "' and bodega='" + Item.SubItems(1) + "'"
               DatOS.Execute S, 128
               ListView1.ListItems.Remove Item.Index
               Call GBitacora(5, "Requisición No. " + Item.Text)
               Anula = True
               Set Temp = Nothing
          End If
     Case 17 'Ordenes de compra
          S = "delete from desgord"
          S = S + " where cia='" + CiA + "' and orden='" + Item.Text + "'"
          DatOS.Execute S, 128
          S = "update ordenes set total=0,"
          S = S + "subtotal=0,descuento=0,impuesto=0,TCAMBIO=0"
          S = S + " where cia='" + CiA + "' and numero='" + Item.Text + "'"
          DatOS.Execute S, 128
          Anula = True
          Item.Icon = 4
          Item.SmallIcon = 4
     End Select
     If ListView1.ListItems.Count > 0 Then
          ListView1.SelectedItem.Selected = True
          ListView1.SetFocus
     End If
Errores:
     If err.Number <> 0 Then
          s1 = err.Description + Str(err.Number)
          s1 = s1 + Chr(13) + S
          MsgBox s1, 16, "Anula"
     End If
End Function
