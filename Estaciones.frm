VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Estaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estaciones"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   6945
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   5370
      TabIndex        =   1
      Top             =   0
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "inserta"
            Object.ToolTipText     =   "Incluir una zona"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modifica"
            Object.ToolTipText     =   "Modificar el registro seleccionado"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "borra"
            Object.ToolTipText     =   "Borrar el registro seleccionado"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sale"
            Object.ToolTipText     =   "Cerrar esta ventana"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   5685
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   10028
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripción"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Precio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "MediaHora"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Estaciones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Estaciones.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Estaciones.frx":042C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Estaciones.frx":0746
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Estaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Item As ListItem
Dim S$
Private Sub Borra(NItem As ListItem)
     On Error GoTo Errores
     If MsgBox("Desea borrar el registro ?", 36, NItem.SubItems(1)) = 6 Then
          S = "delete from estaciones where cia='" + CiA + "' and Codigo='"
          S = S + Trim(NItem.Text) + "'"
          DatOS.Execute S, 128
          Call GBitacora(3, "Estación: " + NItem.Text + " " + NItem.SubItems(1))
          ListView1.ListItems.Remove NItem.Index
          If ListView1.ListItems.Count > 0 Then
               ListView1.SelectedItem.Selected = True
          End If
     End If
     ListView1.SetFocus
     Toolbar1.Buttons(2).Enabled = False
     Toolbar1.Buttons(3).Enabled = False
Errores:
     If err.Number = 3200 Then
          S = "Existen registros asociados a esta zona,"
          S = S + Chr(13) + "    No es posible su eliminación."
          MsgBox S, 48, "Advertencia"
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Borra"
     End If
     On Error GoTo 0
End Sub
Private Sub Form_Activate()
     Call Menus(True, Me)
End Sub
Private Sub Form_Deactivate()
     Call Menus(False, Me)
End Sub
Private Sub Form_Load()
     Top = (Inicio.Height - Height) / 2
     Left = (Inicio.Width - Width) / 2
     Show
     Refresh
     Dim Temp As Recordset
     S = "select * from estaciones where cia='" + CiA + "' order by codigo"
     Set Temp = DatOS.OpenRecordset(S)
     With Temp
     Do Until .EOF
          Set Item = ListView1.ListItems.Add()
          Item.Text = !Codigo
          Item.SubItems(1) = !Descripcion
          Item.SubItems(2) = !PreciO
          Item.SubItems(3) = !mediahora
          .MoveNext
          w% = DoEvents
     Loop
     End With
     Call Barra(Trim(ListView1.ListItems.Count) + " registros encontrados", 1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Menus(False, Me)
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
     ListView1.Sorted = True
     ListView1.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub Listview1_DblClick()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then
          Call DetalleEsta.Carga(ListView1.SelectedItem)
          DetalleEsta.Tag = 2
          DetalleEsta.Show 1
          Toolbar1.Buttons(2).Enabled = False
          Toolbar1.Buttons(3).Enabled = False
     End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     Toolbar1.Buttons(2).Enabled = True
     Toolbar1.Buttons(3).Enabled = True
End Sub
Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If KeyCode = vbKeyDelete And ListView1.SelectedItem.Selected Then
          Call Borra(ListView1.SelectedItem)
     End If
End Sub
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then
          PopupMenu Inicio.Popup, , x, y, Inicio.SubPop(0)
     End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
     Select Case Button.Key
     Case "inserta"
          If Not Licencia(ListView1, 2) Then Exit Sub
          DetalleEsta.Tag = 1
          DetalleEsta.Show 1
          Toolbar1.Buttons(2).Enabled = False
          Toolbar1.Buttons(3).Enabled = False
     Case "modifica"
          Listview1_DblClick
     Case "borra"
          Call Borra(ListView1.SelectedItem)
     Case "sale"
          Unload Me
     End Select
End Sub
