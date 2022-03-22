VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DetOtros 
   Caption         =   "Detalle de Productos Facturados"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   840
      Left            =   7950
      Picture         =   "DetOtros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5940
      Width           =   1035
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5220
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9208
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cant"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DETALLE DE OTROS PRODUCTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1350
      TabIndex        =   1
      Top             =   165
      Width           =   5670
   End
End
Attribute VB_Name = "DetOtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DetOtros.Visible = False
End Sub

Private Sub Form_Load()
    Centra Me
End Sub
