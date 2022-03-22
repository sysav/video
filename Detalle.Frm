VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Detalle 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Leyenda"
   ClientHeight    =   4980
   ClientLeft      =   2370
   ClientTop       =   3015
   ClientWidth     =   1965
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Detalle.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4980
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView ListView2 
      Height          =   4845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   8546
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   3870
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":0442
            Key             =   ""
            Object.Tag             =   "Entrada al Sistema"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":075C
            Key             =   ""
            Object.Tag             =   "Inclusión"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":0A76
            Key             =   ""
            Object.Tag             =   "Modificación"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":0D90
            Key             =   ""
            Object.Tag             =   "Borrado"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":10AA
            Key             =   ""
            Object.Tag             =   "Impresión"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":13C4
            Key             =   ""
            Object.Tag             =   "Anulación"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":16DE
            Key             =   ""
            Object.Tag             =   "Proceso"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Detalle.frx":19F8
            Key             =   ""
            Object.Tag             =   "Salida del Sistema"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
     Call Posiciona(Me, 1)
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Load()
     SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H1 + &H2
     Call Posiciona(Me, 0)
End Sub
Public Sub Carga(Form As Form, Optional Iconos As ImageList)
     Dim Item As ListItem
     Dim ListaI As ImageList
     Select Case LCase(Form.Name)
     Case "bitacora"
          Set ListaI = ImageList2
     End Select
     For I = 1 To ListaI.ListImages.Count
          Set Item = ListView2.ListItems.Add(, , ListaI.ListImages(I).Tag, I)
     Next
     Show
     Refresh
End Sub
