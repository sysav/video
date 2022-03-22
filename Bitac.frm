VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{B4409115-5405-11D3-943D-0080AD4162AE}#1.0#0"; "ECOMBO.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Bitacora 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitácora del Sistema"
   ClientHeight    =   5985
   ClientLeft      =   1800
   ClientTop       =   1605
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bitac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   10215
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Left            =   1200
      TabIndex        =   14
      Top             =   60
      Width           =   4110
      _ExtentX        =   7250
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3630
      TabIndex        =   12
      Top             =   450
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23396353
      CurrentDate     =   36670
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1200
      TabIndex        =   13
      Top             =   450
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      _Version        =   393216
      Format          =   23396353
      CurrentDate     =   36670
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Refrescar"
      Enabled         =   0   'False
      Height          =   570
      Left            =   5121
      Picture         =   "Bitac.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5370
      Width           =   1155
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Buscar en todo el texto"
      Height          =   255
      Left            =   5430
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   4155
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   1185
      Left            =   6090
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3390
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Iconos"
      Height          =   570
      Left            =   3936
      Picture         =   "Bitac.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5370
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   570
      Left            =   6321
      Picture         =   "Bitac.frx":093E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5370
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Desplegar"
      Height          =   570
      Left            =   2739
      Picture         =   "Bitac.frx":0C80
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5370
      Width           =   1155
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4005
      Left            =   150
      TabIndex        =   0
      Top             =   1290
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   7064
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nombre"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Día"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Hora"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   870
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde :"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   570
      TabIndex        =   5
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta :"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3030
      TabIndex        =   4
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario :"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   420
      TabIndex        =   3
      Top             =   105
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":0D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":109C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":13B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":16D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":19EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":1D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":201E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bitac.frx":2338
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fechas As Recordset
Dim Item As ListItem
Dim Conta As Long
Private Sub Command1_Click()
     On Error GoTo Errores
     Call Barra("Un momento por favor ...")
     MousePointer = 11
     Conta = 0
     ListView1.ListItems.Clear
     s = "select bitacora.descripcion,nombre,bitacora.fecha,cia,"
     s = s + "bitacora.hora,tipo,bitacora.login "
     s = s + "from bitacora left join usuarios "
     s = s + "on usuarios.login=bitacora.login "
     s = s + " where bitacora.sistema='" + NomSis + "'"
     Dim Si%
     If ECombo1.ListIndex > 0 Then
          s = s + " and bitacora.login='" + ECombo1.Indice(1) + "'"
     End If
     s = s + " and bitacora.fecha >= #" + Format(DTPicker1, "m/d/yyyy")
     s = s + "# and bitacora.fecha <= #"
     s = s + Format(DTPicker2, "m/d/yyyy") + "#"
     If Text2 <> "" Then
          s = s + " and bitacora.descripcion like '"
          If Check1 Then s = s + "*"
          s = s + Text2 + "*' "
     End If
     s = s + " order by nombre,bitacora.fecha desc,bitacora.hora"
     Set Fechas = DatOS.OpenRecordset(s)
     ListView1.Visible = False
     Dim Llave$
     Do Until Fechas.EOF
          Llave = "*" + Fechas!LoGiN + Format(Fechas!Fecha, "dd/mm/yyyy") + Fechas!Hora
          Set Item = ListView1.ListItems.Add(, Llave, Nulo(Fechas!Nombre))
          Item.Tag = Fechas!CiA
          Item.SmallIcon = Fechas!Tipo + 1
          Item.SubItems(1) = Format(Fechas!Fecha, "dd/mm/yyyy")
          Item.SubItems(2) = Fechas!Hora
          Item.SubItems(3) = Fechas!Descripcion
          Fechas.MoveNext
          w% = DoEvents
     Loop
     Command4.Enabled = True
     Call Barra(ListView1.ListItems.Count & " registros encontrados", 1)
Errores:
     If err.Number = 35602 Then
          Resume Next
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number)
     End If
     On Error GoTo 0
     ListView1.Visible = True
     MousePointer = 0
End Sub
Private Sub Command2_Click()
     Call Detalle.Carga(Me)
End Sub
Private Sub Command3_Click()
     Unload Me
End Sub
Private Sub Command4_Click()
     Call Barra("Un momento por favor ...")
     MousePointer = 11
     Fechas.Requery
     ListView1.ListItems.Clear
     ListView1.Visible = False
     Dim Llave$
     Do Until Fechas.EOF
          Llave = "*" + Fechas!LoGiN + Format(Fechas!Fecha, "dd/mm/yyyy") + Fechas!Hora
          Set Item = ListView1.ListItems.Add(, Llave, Nulo(Fechas!Nombre))
          Item.SmallIcon = Fechas!Tipo + 1
          Item.SubItems(1) = Format(Fechas!Fecha, "dd/mm/yyyy")
          Item.SubItems(2) = Fechas!Hora
          Item.SubItems(3) = Fechas!Descripcion
          Fechas.MoveNext
          w% = DoEvents
     Loop
     ListView1.Visible = True
     MousePointer = 0
     Barra ""
End Sub

Private Sub Form_Load()
     Call Posiciona(Me, 0)
     Dim Temp As Recordset
     s = "select login,nombre from usuarios order by nombre"
     Set Temp = DatOS.OpenRecordset(s)
     ECombo1.AddItem "Todos"
     Do Until Temp.EOF
          ECombo1.AddItem Temp!Nombre, Temp!LoGiN
          Temp.MoveNext
     Loop
     ECombo1.ListIndex = 0
     DTPicker1 = Format(Date, "dd/mm/yyyy")
     DTPicker2 = Date
     Show
     Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Posiciona(Me, 1)
     Unload Detalle
End Sub
Private Sub ListView1_Click()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then Text1.Visible = False
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
     MousePointer = 11
     ListView1.SortOrder = IIf(ListView1.SortOrder = 0, 1, 0)
     ListView1.SortKey = ColumnHeader.Index - 1
     ListView1.Sorted = True
     MousePointer = 0
End Sub
Private Sub Listview1_DblClick()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then Text1.Visible = True
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     Text1.Visible = False
     Text1 = TipoAccion(Item.SmallIcon - 1) + Item.SubItems(3)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case KeyAscii
     Case 8, 47 To 57
          Command1.Enabled = True
     Case Else
          KeyAscii = 0
     End Select
End Sub
Private Function TipoAccion(Modo%) As String
     Select Case Modo
     Case 1
          TipoAccion = "Insertar : "
     Case 2
          TipoAccion = "Actualización : "
     Case 3
          TipoAccion = "Eliminar : "
     Case 4
          TipoAccion = "Impresión : "
     Case 5
          TipoAccion = "Anular : "
     Case 6
          TipoAccion = "Proceso : "
     End Select
End Function
'Private Sub BuscaTexto()
'     On Error GoTo Errores
'     MousePointer = 11
'     Dim Fechas As Recordset
'     S = "select fecha,hora,login "
'     S = S + "from bitacora "
'     S = S + "where descripcion like '"
'     If Check1 Then S = S + "*"
'     S = S + Text2.Text + "'* "
'     If ECombo1.ListIndex > 0 Then S = S + " and login='" + ECombo1.Indice(1) + "'"
'     S = S + " and bitacora.descripcion"
'     If Text(0).Text <> "" And Text(1).Text <> "" Then
'          S = S + " and fecha >= #" + Format(Text(0).Text, "m/d/yyyy")
'          S = S + "# and fecha <= #" + Format(Text(1).Text, "m/d/yyyy") + "#"
'     End If
'     Set Fechas = DatOS.OpenRecordset(S)
'     If Not Fechas.EOF Then
'          Dim Llave$
'          Llave = "*" + Fechas!LoGiN + Format(Fechas!Fecha, "dd/mm/yyyy") + Fechas!Hora
'          Set Item = ListView1.ListItems(Llave)
'          Item.EnsureVisible
'          Item.Selected = True
'     End If
'     MousePointer = 0
'     Call Barra(Trim(ListView1.ListItems.Count) + " registros encontrados", 1)
'Errores:
'     If Err.Number = 3075 Then
'          MsgBox "Fecha Inválida !", 16, "Error"
'          Text(0).SelStart = 1
'          Text(0).SelLength = Len(Text(0).Text)
'     ElseIf Err.Number > 0 Then
'          MsgBox Err.Description + Str(Err.Number)
'     End If
'     On Error GoTo 0
'End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
     Case 10, 13, 39
          KeyAscii = 0
     End Select
End Sub
