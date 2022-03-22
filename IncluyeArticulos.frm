VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form AddArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Artículos"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IncluyeArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6705
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      TabIndex        =   0
      Top             =   90
      Width           =   2025
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   3
      Left            =   1230
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Width           =   825
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   345
      Left            =   2250
      TabIndex        =   10
      Top             =   870
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   609
      _Version        =   327681
      BuddyControl    =   "Text(1)"
      BuddyDispid     =   196610
      BuddyIndex      =   1
      OrigLeft        =   2220
      OrigTop         =   870
      OrigRight       =   2415
      OrigBottom      =   1185
      Max             =   1000000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   2
      Left            =   4470
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1350
      Width           =   825
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   0
      Left            =   4470
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   870
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1710
      Picture         =   "IncluyeArticulos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   2970
      Picture         =   "IncluyeArticulos.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   1
      Left            =   1230
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   870
      Width           =   990
   End
   Begin VB.Label Label18 
      Caption         =   "Costo ¢:"
      Height          =   255
      Left            =   3540
      TabIndex        =   14
      Top             =   930
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descuento :"
      Height          =   225
      Left            =   180
      TabIndex        =   13
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   510
      Width           =   1050
   End
   Begin VB.Label Descripcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1230
      TabIndex        =   11
      Top             =   450
      Width           =   5025
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "I.V. (%) :"
      Height          =   225
      Left            =   3630
      TabIndex        =   9
      Top             =   1380
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cantidad :"
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
      Left            =   210
      TabIndex        =   8
      Top             =   870
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Artículo :"
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
      Left            =   285
      TabIndex        =   7
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "AddArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Padre As Form
Dim Cant@
Dim Estado%
Dim SQLStr$
Dim Como As Boolean
Dim Articulo As Recordset
Dim Item As ListItem
Dim S$
Private Sub Command1_Click()
     On Error GoTo Errores
     Dim Monto@
     Dim Impuesto@
     Dim Descuento@
     Dim Item2 As ListItem
     Dim Llave$
     Dim Costo@
     Dim CostoFijo@
     Dim CT@
     Dim Cambio@
     Dim Cantidad@
     Set Item = Nothing
     Costo = Doble(Text(0))
     CostoFijo = Costo
     'Rebaja el descuento si hubiera
     Costo = Costo - (Costo * (Doble(Text(3)) / 100))
     If Costo = 0 Then
          S = "El costo debe ser mayor a cero !"
          MsgBox S, 16, "Error en el costo"
          Text(0).SetFocus
          Exit Sub
     End If
     Cantidad = Doble(Text(1))
     If Not ValidaExis(Text1, Doble(Text(1)), Doble(Text(4))) Then
          Selecciona Text(1)
          Exit Sub
     End If
     If Estado < 2 And Doble(Text(1)) = 0 Then err.Raise 9999
     If Estado = 2 And Doble(Text(1)) > Cant Then err.Raise 9998
     Llave = "*" + Text1
     If Not Como Then
          Set Item2 = Padre.ListView2.ListItems.Add(, Llave, , , 5)
          Set Item = Padre.ListView1.ListItems.Add
     Else
          Set Item2 = Padre.ListView2.SelectedItem
          Item2.Key = Llave
          Set Item = Padre.ListView1.ListItems(Item2.Index)
     End If
     If CosTeO = 2 Then     'Costo promedio
          Dim Temp As Recordset
          S = "select sum(existencia) from existencias "
          S = S + "where cia='" + CiA + "' and c_bodega='" + Padre.Text(4)
          S = S + "' and c_articulo='" + Text1 + "'"
          Set Temp = DatOS.OpenRecordset(S)
          If Not IsNull(Temp(0)) Then CT = Temp(0)
          Costo = (CT * Articulo!p_compra)
          Costo = Costo + (Cantidad * CostoFijo)
          Costo = Costo / IIf(CT + Cantidad <> 0, CT + Cantidad, 1)
     End If
     Monto = CostoFijo * Cantidad     ' ojo  costo*cantidad estaba !! 10/08/00
     Descuento = Monto * (Doble(Text(3)) / 100)
     Impuesto = (Doble(Text(2)) / 100) * (Monto - Descuento)
     Item2.Tag = Costo ' costo promedio  so far
     Item2.Text = Text1 'El codigo
     Item2.SubItems(1) = Descripcion.Caption 'La descripcion
     Item2.SubItems(2) = ""         'El lote
     Item2.SubItems(3) = FormatNumber(CostoFijo, DeCiMaleS) 'El costo
     Item2.SubItems(4) = FormatNumber(Costo, DeCiMaleS) 'El nuevo costo
     Item2.SubItems(5) = Cantidad 'La cantidad
     Item2.SubItems(6) = Doble(Text(3)) 'El % de descuento
     Item2.SubItems(7) = Doble(Text(2)) 'El % de impuesto
     Item2.SubItems(8) = FormatNumber(Monto) 'El total bruto
     
     'Item2.SubItems(9) = Text(7)     ' FOB $
     'Item2.SubItems(10) = Text(12)     ' pRECIO 1 cOLONES COLONES
     'Item2.SubItems(11) = Text(11)    ' PRECIO 2 DOLARES $$$$$
     'Item2.SubItems(12) = Text(9)     ' FACTOR 1
     'Item2.SubItems(13) = Text(10)     'FACTOR 2
     'Item2.SubItems(14) = Text(13)     ' ETIQUETAS
     
     Item2.Selected = True
     Item2.EnsureVisible
     'La lista invisible
     'Item.Tag = Doble(Text(7)) 'El costo extranjero
     Item.Text = Impuesto 'El impuesto en colones
     Item.SubItems(1) = Descuento 'El descuento en colones
     'Item.SubItems(2) = Doble(Text(4)) 'El nuevo maximo
     'Item.SubItems(3) = Doble(Text(5)) 'El nuevo utilidad
     Item.SubItems(4) = Doble(Text(0).Tag) 'El costo anterior
     Call Padre.CalculaSaldo
     If Not Como Then
          Call Limpia
     Else
          Unload Me
     End If
Errores:
     If err.Number = 35601 Then
          Resume Next
     ElseIf err.Number = 35602 Then
          MsgBox "Ya incluyó ese artículo!", 16, "Error"
          Text1.SetFocus
     ElseIf err.Number = 9999 Then
          MsgBox "La cantidad debe ser mayor a cero !", 16, "Error"
          Selecciona Text(1)
     ElseIf err.Number = 9998 Then
          MsgBox "La cantidad debe ser menor a la cantidad entrante !", 16, "Error"
          Selecciona Text(1)
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Command1_Click"
     End If
     On Error GoTo 0
     Como = False
End Sub
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Command3_Click()
     ListView1.SelectedItem.SubItems(1) = Format(Doble(Text2.Text), "standard")
     ListView1.SelectedItem.SubItems(2) = Format(Doble(Padre.Costos(1).Text), "standard")
     Text2.Text = ""
     ListView1.SetFocus
End Sub
Private Sub Descripcion_Change()
     Call Valida
End Sub
Private Sub Form_Load()
     Call Posiciona(Me, 0)
     If CosTeO = 0 Then
          Caption = Caption + " (Costo Simple)"
     ElseIf CosTeO = 1 Then
          Caption = Caption + " (Costeo de Importaciones)"
     ElseIf CosTeO = 2 Then
          Caption = Caption + " (Costo Promedio)"
     End If
     Set Articulo = Nothing
     S = "select * from articulos where cia='" + CiA + "' order by d_articulo"
     Set Articulo = DatOS.OpenRecordset(S, dbOpenSnapshot)
     Dim Temp As Recordset
     S = "Select * from parametros where cia='" + CiA + "'"
     Set Temp = DatOS.OpenRecordset(S)
     'If Not Temp.EOF Then
     '     Text(8) = Temp!TCambIO
     '     Text(9) = Temp!fact1
     '     Text(10) = Temp!fact2
     'End If
     Temp.Close
     Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Posiciona(Me, 1)
     Como = False
End Sub
Private Sub Label10_Change()
     Call Valida
End Sub

Private Sub Text_Change(Index As Integer)
Select Case Index
'     Case 1      ' la cantidad
'          Text(13) = Text(1)     ' las etiquetas
'     Case 7, 8, 9, 10 ' Cambia el FOB $ costo extranjero  o 8 el TC, f1,f2
'          Text(0) = Doble(Text(7)) * Doble(Text(8)) * Doble(Text(9))  ' fob * tc * f1
'          If Doble(Text(0)) > 0 Then
'             Text(15) = Doble(Text(7)) * Doble(Text(9))  ' fob * f1  costo $ +f1
'             Text(16) = Doble(Text(7)) * Doble(Text(8))   ' fob * tc  osea fob colones
'             Calcula_Precio
'          End If
'     Case 0, 9, 10      ' costo colones,tc, f1, f2
'     If Doble(Text(0)) > 0 Then       ' costo en colones calculado
'          Calcula_Precio
'     End If
End Select
Call Valida
End Sub
Private Sub Valida()
     Command1.Enabled = False
     If Text(0) = "" Then Exit Sub
     If Text(1) = "" Then Exit Sub
     If Descripcion.Caption = "" Then Exit Sub
     Command1.Enabled = True
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case KeyAscii
     Case 8, 46, 47 To 58
     Case Else
          KeyAscii = 0
     End Select
     If Index = 1 And KeyAscii = 13 And Command1.Enabled Then Command1_Click
End Sub
Public Sub Carga(forma As Form, Nueva%, Optional Item As ListItem, _
     Optional Tipo%)
     On Error GoTo Errores
     Set Padre = forma
     Text(3) = Padre.Text(5)
     If Nueva = 1 Then
          Dim Llave$
          Dim Item4 As ListItem
          Dim Costo!
          Como = True
          Estado = Tipo
          Cant = Doble(Item.SubItems(5))
          Costo = Doble(Item.SubItems(3))
          Text1 = Item.Text 'Codigo
          Text(1) = Cant 'Cantidad
          Text(2) = Item.SubItems(7) 'IV
          Text(3) = Item.SubItems(6) 'Descuento
          'Text(4) = Padre.ListView1.ListItems(Item.Index).SubItems(2) 'Maximo
          'Text(5) = Padre.ListView1.ListItems(Item.Index).SubItems(3) 'Utilidad
          Text(0) = Costo
          Text(0).Tag = Padre.ListView2.ListItems(Item.Index).SubItems(4)
          'Text(7) = Padre.ListView2.ListItems(Item.Index).SubItems(9)   'FOB $
          'Text(13) = Padre.ListView2.ListItems(Item.Index).SubItems(14) 'ETIQUETAS
     End If
Errores:
     If err.Number = 35601 Then
          Set Item4 = Nothing
          Resume Next
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Carga"
     End If
     On Error GoTo 0
End Sub
Private Sub Limpia()
      
     Text1 = ""
     Text2 = ""
     Text3 = ""
     Text1.SetFocus
End Sub
Private Function ValidaExis(Codigo$, Cantidad#, Maximo#) As Boolean
     Dim Actuales&
     Dim Temp As Recordset
     S = "select sum(existencia) from existencias where cia='" + CiA + "' and c_articulo='"
     S = S + Codigo + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If IsNull(Temp(0)) Then
          Actuales = 0
     Else
          Actuales = Temp(0)
     End If
     'If (Actuales + Cantidad) > Maximo Then
     '     S = "La cantidad entrante más la existencia en bodegas "
     '     S = S + Chr(13) + "es mayor al máximo del articulo, desea continuar ?"
     '     If MsgBox(S, 36, "Entradas") = 6 Then ValidaExis = True
     'Else
          ValidaExis = True
     'End If
End Function
Private Sub Text_LostFocus(Index As Integer)
Select Case Index
     Case 12    ' PRECIO COLONES A PIE
     Precio_Apie
End Select
End Sub

Private Sub Text1_Change()
     Descripcion.Caption = ""
     Text(0) = ""
     Text(2) = ""
     'Text(4) = ""
     'Text(5) = ""
     'Text(6) = ""
     'Text(7) = ""
     'Text(11) = ""
     'Text(12) = ""
     'Text(14) = ""
     Articulo.FindFirst "c_Articulo='" + Text1 + "'"
     If Not Articulo.NoMatch Then
          Descripcion.Caption = Articulo!d_articulo
          Text(0) = FormatNumber(Articulo!p_compra, DeCiMaleS)
          Text(0).Tag = Articulo!p_compra
          Text(2) = IV * 100
          'Text(5) = Articulo!porc_util
          'Text(4) = FormatNumber(Articulo!Maximo)
          'Text(6) = FormatNumber(Articulo!CostoAnt, DeCiMaleS)
          'Text(7) = FormatNumber(Articulo!CostoEx, DeCiMaleS)
     End If
End Sub
Private Sub Text1_GotFocus()
     S = "Presione F3 para buscar por lista - F5 para refrescar artículos"
     Call Barra(S)
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF3 Then
          S = "select c_articulo as Código,d_articulo as Descripción "
          S = S + "from articulos where cia='" + CiA + "' order by d_articulo"
          Call Lista.Carga(Text1, S, "Artículos")
          Lista.Show 1
     ElseIf KeyCode = vbKeyF5 Then
          MousePointer = 11
          Barra "Refrescando lista de artículos ..."
          Articulo.Requery
          MousePointer = 0
          Barra ""
     End If
End Sub
Private Sub Text1_LostFocus()
     Barra ""
End Sub
Private Sub Text2_Change()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If Text2.Text <> "" And ListView1.SelectedItem.Selected Then
          Command3.Enabled = True
     Else
          Command3.Enabled = False
     End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
     Case 13
          If Text2.Text <> "" Then Command3_Click
     Case 8, 46, 48 To 57
     Case Else
          KeyAscii = 0
     End Select
End Sub
Private Sub Calcula_Precio()    'text(7) fob $
Text(11) = Format(Doble(Text(7)) * Doble(Text(10)), "###,##0.00")
Text(12) = Format(Doble(Text(11)) * Doble(Text(8)), "###,##0.00") ' * TC
Text(14) = Format(Round((Doble(Text(12)) * (1 + IV)) / 2, 0) * 2, "###,##0.00") 'ivi
Text(5) = Format((cero(Doble(Text(11)), Doble(Text(0))) - 1) * 100, "###0.00")   'porcutil s/fob+f1
End Sub
Private Sub Precio_Apie()
Text(11) = Format(cero(Doble(Text(12)), Doble(Text(8))), "###,##0.00") ' Precio en dolares patras
Text(14) = Format(Round((Doble(Text(12)) * (1 + IV)) / 2, 0) * 2, "###,##0.00") 'ivi
Text(5) = Format((cero(Doble(Text(11)), Doble(Text(7))) - 1) * 100, "###0.00")   ' porcutil
End Sub

