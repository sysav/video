VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Existencias 
   Caption         =   "Consulta de Existencias y Precios"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Existencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   11040
      TabIndex        =   17
      Top             =   120
      Width           =   285
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   8430
      TabIndex        =   16
      Top             =   120
      Width           =   285
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   1
      Left            =   9810
      TabIndex        =   14
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   0
      Left            =   7200
      TabIndex        =   12
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   3750
      TabIndex        =   11
      Top             =   60
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   90
      Width           =   285
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4740
      TabIndex        =   8
      Top             =   30
      Width           =   1455
   End
   Begin VB.PictureBox Splitter2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   4080
      MouseIcon       =   "Existencias.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   1380
      ScaleWidth      =   450
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   510
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
      Height          =   1440
      Left            =   4080
      MouseIcon       =   "Existencias.frx":0614
      MousePointer    =   99  'Custom
      ScaleHeight     =   1380
      ScaleWidth      =   450
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   510
      Left            =   270
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9060
      Visible         =   0   'False
      Width           =   795
   End
   Begin ComctlLib.ListView ListView3 
      Height          =   2085
      Left            =   4740
      TabIndex        =   3
      ToolTipText     =   "Lista de precios"
      Top             =   3720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3678
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   16777215
      BackColor       =   16744576
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P. Local"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P. Local (IV)"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P. Extranjero"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Precio Extranjero (IV)"
         Object.Width           =   1764
      EndProperty
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   2805
      Left            =   4770
      TabIndex        =   2
      ToolTipText     =   "Lista de existencias por bodega y lote"
      Top             =   840
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   16711680
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
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
         Text            =   "Bodega"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Existencia"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lote"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4965
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "La lista de artículos"
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8758
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
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Columna 3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   405
      Width           =   8535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   405
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Fabricante"
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   15
      Top             =   60
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Alterno"
      Height          =   255
      Index           =   0
      Left            =   6570
      TabIndex        =   13
      Top             =   60
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   4050
      TabIndex        =   9
      Top             =   60
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion :"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   1050
   End
End
Attribute VB_Name = "Existencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextBox As TextBox
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
Private GlbfrmInSizeY As Long
Dim Item As ListItem
Dim Temp As Recordset
     
Private Sub CargaExis(Articulo$)
     MousePointer = 11
     Dim Total@
     ListView2.ListItems.Clear
     Dim Tabla As Recordset
     S = "select existencias.c_bodega,d_bodega,existencia,lote "
     S = S + "from bodegas left join existencias "
     S = S + "on existencias.cia=bodegas.cia "
     S = S + "and existencias.c_bodega=bodegas.c_bodega "
     S = S + "where existencias.cia='" + CiA
     S = S + "' and existencias.c_articulo='" + Articulo
     S = S + "' order by existencias.c_bodega,lote"
     Set Tabla = DatOS.OpenRecordset(S, dbOpenSnapshot)
     Do Until Tabla.EOF
          Set Item = ListView2.ListItems.Add
          Item.Text = Tabla!c_bodega
          Item.SubItems(1) = Tabla!d_bodega
          Item.SubItems(2) = FormatNumber(Doble(Tabla!Existencia))
          Item.SubItems(3) = Nulo(Tabla!Lote)
          Total = Total + Nulo(Tabla!Existencia)
          Tabla.MoveNext
          w% = DoEvents
     Loop
     Set Item = ListView2.ListItems.Add
     If Total > 0 Then
          Item.SubItems(1) = "Total: "
          Item.SubItems(2) = Format(Total, "standard")
     Else
          Item.SubItems(1) = "No tiene existencias !"
     End If
     MousePointer = 0
End Sub
Private Sub Command1_Click()
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     ListView3.ListItems.Clear
     S = "select c_articulo,d_articulo from articulos "
     S = S + "where cia='" + CiA + "' and c_articulo like '*" + Text2.Text + "*' order by d_articulo "
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add(, , Temp!c_articulo)
          Item.SubItems(1) = Temp!d_articulo
          List1.AddItem Temp!d_articulo
          Temp.MoveNext
     Loop
     ListView1.Visible = True
     'ListView1.Width = CSng(ReadKey("ArbolTreeWidth", "xy"))
     'ListView2.Height = CSng(ReadKey("ListaHeight", "xy"))
     'Y1 = Text1.Top + Text1.Height + 80
     'GlbfrmInSizeX = &H7FFFFFFF
     'GlbfrmInSizeY = &H7FFFFFFF
     ListView3.ColumnHeaders(2).Text = MoNLoC
     ListView3.ColumnHeaders(3).Text = MoNLoC + " IV"
     ListView3.ColumnHeaders(4).Text = MoNExT
     ListView3.ColumnHeaders(5).Text = MoNExT + " IV"
     If UsaLotes = 0 Then
          ListView2.ColumnHeaders(4).Width = 0
     End If
End Sub

Private Sub Command2_Click()
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     ListView3.ListItems.Clear
     S = "select c_articulo,d_articulo from articulos "
     S = S + "where cia='" + CiA + "' and d_articulo like '*" + Text1 + "*' order by d_articulo "
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add(, , Temp!c_articulo)
          Item.SubItems(1) = Temp!d_articulo
          List1.AddItem Temp!d_articulo
          Temp.MoveNext
     Loop
     ListView1.Visible = True
     'ListView1.Width = CSng(ReadKey("ArbolTreeWidth", "xy"))
     'ListView2.Height = CSng(ReadKey("ListaHeight", "xy"))
     'Y1 = Text1.Top + Text1.Height + 80
     'GlbfrmInSizeX = &H7FFFFFFF
     'GlbfrmInSizeY = &H7FFFFFFF
     ListView3.ColumnHeaders(2).Text = MoNLoC
     ListView3.ColumnHeaders(3).Text = MoNLoC + " IV"
     ListView3.ColumnHeaders(4).Text = MoNExT
     ListView3.ColumnHeaders(5).Text = MoNExT + " IV"
     If UsaLotes = 0 Then
          ListView2.ColumnHeaders(4).Width = 0
     End If
     'Call Form_Resize
End Sub

Private Sub Command3_Click()
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     ListView3.ListItems.Clear
     S = "select c_articulo,alterno,d_articulo from articulos "
     S = S + "where cia='" + CiA + "' and alterno like '*" + Text3(0) + "*' order by d_articulo "
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add(, , Temp!c_articulo)
          Item.SubItems(1) = Temp!d_articulo
          List1.AddItem Temp!d_articulo
          Item.SubItems(2) = Temp!Alterno
          List1.AddItem Temp!Alterno
          Temp.MoveNext
     Loop
     ListView1.Visible = True
     'ListView1.Width = CSng(ReadKey("ArbolTreeWidth", "xy"))
     'ListView2.Height = CSng(ReadKey("ListaHeight", "xy"))
     'Y1 = Text1.Top + Text1.Height + 80
     'GlbfrmInSizeX = &H7FFFFFFF
     'GlbfrmInSizeY = &H7FFFFFFF
     ListView3.ColumnHeaders(2).Text = MoNLoC
     ListView3.ColumnHeaders(3).Text = MoNLoC + " IV"
     ListView3.ColumnHeaders(4).Text = MoNExT
     ListView3.ColumnHeaders(5).Text = MoNExT + " IV"
     If UsaLotes = 0 Then
          ListView2.ColumnHeaders(4).Width = 0
     End If
     
End Sub

Private Sub Command4_Click()
     ListView1.ListItems.Clear
     ListView2.ListItems.Clear
     ListView3.ListItems.Clear
     S = "select c_articulo,d_articulo from articulos "
     S = S + "where cia='" + CiA + "' and ubicacion like '*" + Text3(1) + "*' order by d_articulo "
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add(, , Temp!c_articulo)
          Item.SubItems(1) = Temp!d_articulo
          List1.AddItem Temp!d_articulo
          Temp.MoveNext
     Loop
     ListView1.Visible = True
     'ListView1.Width = CSng(ReadKey("ArbolTreeWidth", "xy"))
     'ListView2.Height = CSng(ReadKey("ListaHeight", "xy"))
     'Y1 = Text1.Top + Text1.Height + 80
     'GlbfrmInSizeX = &H7FFFFFFF
     'GlbfrmInSizeY = &H7FFFFFFF
     ListView3.ColumnHeaders(2).Text = MoNLoC
     ListView3.ColumnHeaders(3).Text = MoNLoC + " IV"
     ListView3.ColumnHeaders(4).Text = MoNExT
     ListView3.ColumnHeaders(5).Text = MoNExT + " IV"
     If UsaLotes = 0 Then
          ListView2.ColumnHeaders(4).Width = 0
     End If

End Sub

Private Sub Form_Load()
'Call Posiciona(Me, 0)
Centra Me
     S = "select c_articulo,d_articulo from articulos "
     S = S + "where cia='" + CiA + "' and 1=2  order by d_articulo "
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add(, , Temp!c_articulo)
          Item.SubItems(1) = Temp!d_articulo
          List1.AddItem Temp!d_articulo
          Temp.MoveNext
     Loop
     ListView1.Visible = True
     'ListView1.Width = CSng(ReadKey("ArbolTreeWidth", "xy"))
     'ListView2.Height = CSng(ReadKey("ListaHeight", "xy"))
     'Y1 = Text1.Top + Text1.Height + 80
     'GlbfrmInSizeX = &H7FFFFFFF
     'GlbfrmInSizeY = &H7FFFFFFF
     'ListView3.ColumnHeaders(2).Text = MoNLoC
     'ListView3.ColumnHeaders(3).Text = MoNLoC + " IV"
     'ListView3.ColumnHeaders(4).Text = MoNExT
     'ListView3.ColumnHeaders(5).Text = MoNExT + " IV"
     'If UsaLotes = 0 Then
     '     ListView2.ColumnHeaders(4).Width = 0
     'End If
     'Call Form_Resize

End Sub
Private Sub Form_Unload(Cancel As Integer)
     'Call SaveKey("ArbolTreeWidth", ListView1.Width, "xy")
     'Call SaveKey("ListaHeight", ListView2.Height, "xy")
     'Call Posiciona(Me, 1)
End Sub

'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'     ListView1.SortKey = ColumnHeader.Index - 1
'     ListView1.SortOrder = IIf(ListView1.SortOrder = 0, 1, 0)
'     ListView1.Sorted = True
'End Sub
Private Sub Listview1_DblClick()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected And Not (TextBox Is Nothing) Then
          TextBox = ListView1.SelectedItem.Text
          Unload Me
     End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     Label4 = Item.Text
     Label5 = Item.SubItems(1)
     Call CargaExis(Item.Text)
     Call CargaPrecios(Item.Text)
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then Listview1_DblClick
End Sub
Private Sub Text1_Change()
     Dim Res&
     Dim C$
     'C = Text1
     'Res = SendMessage(List1.hwnd, LB_FINDSTRING, -1, C)
     'If Res > -1 Then
     '     Set Item = ListView1.ListItems(Res + 1)
     '     Item.Selected = True
     '     Item.EnsureVisible
     '     Call CargaPrecios(Item.Text)
     '     Call CargaExis(Item.Text)
     'Else
     '     ListView1.SelectedItem.Selected = False
     '     ListView2.ListItems.Clear
     '     ListView3.ListItems.Clear
     'End If
End Sub

Private Sub CargaPrecios(Articulo$)
     MousePointer = 11
     ListView3.ListItems.Clear
     Dim Tabla As Recordset
     Dim PrecLoc#
     Dim PrecExt#
     If PreCBoD = 0 Then
          S = "select descripcion,monto from listaprecios where cia='" + CiA
          S = S + "' and codart='" + Articulo + "' order by codigo"
     Else
          S = "select monto,unidades.descripcion "
          S = S + "from precios left join unidades "
          S = S + "on unidades.cia=precios.cia "
          S = S + "and unidades.codart=precios.codart "
          S = S + "and unidades.unidades=precios.numuni "
          S = S + "where precios.cia='" + CiA
          S = S + "' and precios.codart='" + Articulo + "'"
     End If
     Set Tabla = DatOS.OpenRecordset(S, dbOpenSnapshot)
     If Not Tabla.EOF Then
          Do Until Tabla.EOF
               Set Item = ListView3.ListItems.Add
               Item.Text = Tabla!Descripcion
               PrecLoc = Tabla!monto
               PrecExt = cero(Tabla!monto, TCambio)  'precio $ de referencia 31/7/00
               Item.SubItems(1) = FormatNumber(PrecLoc, DeCiMaleS)
               Item.SubItems(2) = FormatNumber(PrecLoc * (1 + IV), DeCiMaleS)
               Item.SubItems(3) = FormatNumber(PrecExt, DeCiMaleS)
               Item.SubItems(4) = FormatNumber(PrecExt * (1 + IV), DeCiMaleS)
               List1.AddItem Tabla!Descripcion
               Tabla.MoveNext
               w% = DoEvents
          Loop
     Else
          Set Item = ListView3.ListItems.Add
          Item.Text = "Precio no Definido!"
     End If
     MousePointer = 0
End Sub
Public Sub Carga(Texto As TextBox)
     'Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 1)
     Set TextBox = Texto
     Call Posiciona(Me, 0)
     Show 1
     Refresh
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then Unload Me
End Sub
