VERSION 5.00
Object = "{B4409115-5405-11D3-943D-0080AD4162AE}#1.0#0"; "ECOMBO.OCX"
Begin VB.Form DetallePrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Precios"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DetallePrecio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Left            =   1260
      TabIndex        =   1
      Top             =   420
      Width           =   3795
      _ExtentX        =   6694
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
      Index           =   1
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   705
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   330
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   2618
      Picture         =   "DetallePrecio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   1388
      Picture         =   "DetallePrecio.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   0
      Left            =   3570
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   780
      Width           =   1485
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1980
      TabIndex        =   10
      Top             =   60
      Width           =   3075
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Bodega :"
      Height          =   225
      Left            =   510
      TabIndex        =   9
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1260
      TabIndex        =   5
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unidad :"
      Height          =   225
      Left            =   540
      TabIndex        =   7
      Top             =   870
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Precio :"
      Height          =   225
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   585
   End
End
Attribute VB_Name = "DetallePrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp As Recordset
Dim Bodega$
Dim Lista As ListView
Dim S$
Dim Item As ListItem
Private Sub Command1_Click()
     If Val(Tag) = 0 Then
          If Inserta Then
               Text(0).Text = ""
               Text(0).SetFocus
          End If
     ElseIf Val(Tag) = 1 Then
          If Modifica Then Unload Me
     End If
End Sub
Private Function Inserta() As Boolean
     S = "insert into Precios(codart,monto,numuni,codbod,default) "
     S = S + " values ('" + Precios.Text1.Text + "'," 'El codigo del articulo
     S = S + Trim(Doble(Text(0).Text)) + "," 'El monto
     S = S + Label2.Caption + ",'" 'Las unidades
     S = S + Text(1).Text + "',"  'El codigo de la bodega
     Dim Char As String * 1
     Char = "1"
     For I = 1 To Precios.ListView1.ListItems.Count
          If Trim(Precios.ListView1.ListItems(I).Tag) = Text(1).Text Then
               Char = "0"
               Exit For
          End If
     Next
     S = S + Char + ")"
     Inserta = Procesa(S, 1)
     If Inserta Then
          S = "Precio del artículo: " + Precios.Text1.Text
          S = S + ", Bodega: " + Text(1).Text
          S = S + ", Unidad: " + Label2.Caption
          Call GBitacora(1, S)
     End If
End Function
Private Function Procesa(SQL As String, Modo As Integer) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     If Modo = 1 Then
          Set Item = Lista.ListItems.Add(, , , , 6)
     Else
          Set Item = Lista.SelectedItem
     End If
     Item.Tag = Text(1).Text
     Item.Text = Label5.Caption
     Item.SubItems(1) = ECombo1.Text
     Item.SubItems(2) = Label2.Caption
     Item.SubItems(3) = FormatNumber(Text(0).Text, DeCiMaleS)
     Procesa = True
Errores:
     If err.Number = 3022 Then
          MsgBox "Crea una entrada duplicada !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number = 3315 Then
          MsgBox "Al menos uno de los campos vacíos es requerido !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number = 3464 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Procesa"
     End If
     On Error GoTo 0
End Function
Private Function Modifica() As Boolean
     S = "update Precios set monto=" + Trim(Doble(Text(0).Text))
     S = S + " where codart='" + Precios.Text1.Text
     S = S + "' and codbod='" + Text(1).Text
     S = S + "' and numuni=" + Label2.Caption
     S = S + " and monto=" + Trim(Doble(Text(0).Tag))
     Modifica = Procesa(S, 2)
     If Modifica Then
          S = "Precio del artículo: " + Precios.Text1.Text
          S = S + ", Bodega: " + Text(1).Text
          S = S + ", Unidad: " + Label2.Caption
          S = S + ", Monto Anterior: " + Text(0).Tag
          S = S + ", Monto Nuevo: " + Text(0).Text
          Call GBitacora(2, S)
     End If
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub ECombo1_Click()
     Label2.Caption = IIf(ECombo1.ListIndex > -1, ECombo1.Indice(1), "")
     Call Text_Change(0)
End Sub
Private Sub Form_Load()
     S = "select * from unidades where cia='" + CiA + "' and codart='"
     S = S + Precios.Text1.Text + "'"
     Set Temp = DatOS.OpenRecordset(S)
     Do Until Temp.EOF
          ECombo1.AddItem Temp!Descripcion, Temp!Unidades
          Temp.MoveNext
          w% = DoEvents
     Loop
     S = "select c_bodega,d_bodega from bodegas "
     S = S + "where cia='" + CiA + "' order by d_bodega"
     Set Temp = DatOS.OpenRecordset(S)
     Set Lista = Precios.ListView1
End Sub
Public Sub DatosPrecio(Item As ListItem)
     Tag = 1
     Text(1).Text = Item.Tag
     For I = 0 To ECombo1.ListCount - 1
          If Trim(ECombo1.List(I, 1)) = Trim(Item.SubItems(2)) Then
               ECombo1.ListIndex = I
               Exit For
          End If
     Next
     Text(0).Text = Item.SubItems(3)
     Text(0).Tag = Item.SubItems(3)
     Text(1).Enabled = False
End Sub
Private Sub Text_Change(Index As Integer)
     If Text(0).Text <> "" And ECombo1.ListIndex > -1 And Label5.Caption <> "" Then
          Command1.Enabled = True
     Else
          Command1.Enabled = False
     End If
     If Index = 1 Then
          Label5.Caption = ""
          S = "select c_bodega,d_bodega from bodegas "
          S = S + "where cia='" + CiA + "'"
          S = S + " and c_bodega='" + Text(1) + "'"
          Set Temp = DatOS.OpenRecordset(S)
          If Not Temp.EOF Then Label5.Caption = Temp!d_bodega
     End If
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 0 Then
          Select Case KeyAscii
          Case 8, 46, 47 To 58
          Case Else
               KeyAscii = 0
          End Select
     End If
End Sub
Private Sub Limpia()
     ECombo1.ListIndex = -1
     ECombo1.SetFocus
     Text(0).Text = ""
     Command1.Enabled = False
End Sub
