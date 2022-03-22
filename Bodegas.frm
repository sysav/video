VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Bodegas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bodegas"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bodegas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   6090
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   4110
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "inserta"
            Object.ToolTipText     =   "Incluir un registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modifica"
            Object.ToolTipText     =   "Modificar el registro seleccionado"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "borra"
            Object.ToolTipText     =   "Borrar el registro seleccionado"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "imprime"
            Object.ToolTipText     =   "Imprimir la lista de bodegas"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sale"
            Object.ToolTipText     =   "Cerrar esta ventana"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   210
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2505
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
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
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Principal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Principal Entrada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Otra Cia"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bodegas.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bodegas.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bodegas.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bodegas.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Bodegas.frx":095A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Bodegas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Item As ListItem
Dim S$
Private Sub Borra(NItem As ListItem)
     On Error GoTo Errores
     If MsgBox("Desea borrar el registro ?", 36, NItem.SubItems(1)) = 6 Then
          S = "delete from bodegas where cia='" + CiA + "' and c_bodega='"
          S = S + NItem.Text + "'"
          DatOS.Execute S, 128
          Call GBitacora(3, "Bodega : " + NItem.Text + " " + NItem.SubItems(1))
          ListView1.ListItems.Remove NItem.Index
          If ListView1.ListItems.Count > 0 Then
               ListView1.SelectedItem.Selected = True
               ListView1.SetFocus
          End If
     End If
     Call Limpia(False)
Errores:
     If err.Number = 3200 Then
          S = "Existen registros asociados a este registro,"
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
     Centra Me
     Show
     Refresh
     Call Carga("select * from bodegas where cia='" + CiA + "' order by c_bodega")
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Menus(False, Me)
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
     ListView1.SortKey = ColumnHeader.Index - 1
     ListView1.Sorted = True
End Sub
Private Sub Listview1_DblClick()
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If ListView1.SelectedItem.Selected Then
          Call Load(DetalleBodega)
          Call DetalleBodega.DatosCliente(ListView1.SelectedItem)
          DetalleBodega.Tag = 2
          DetalleBodega.Show 1
          Call Limpia(False)
     End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
     Call Limpia(True)
End Sub
Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
     If ListView1.SelectedItem Is Nothing Then Exit Sub
     If KeyCode = vbKeyDelete Then
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
          DetalleBodega.Tag = 1
          DetalleBodega.Show 1
          Call Limpia(False)
     Case "modifica"
          Listview1_DblClick
     Case "borra"
          Call Borra(ListView1.SelectedItem)
     Case "sale"
          Unload Me
     Case "imprime"
          Report1.ReportFileName = DirTrA + "reportes\bodegas.rpt"
          Report1.WindowTitle = "Lista de Bodegas"
          Report1.DataFiles(0) = DatOS.Name
          Report1.Formulas(0) = "comodin='Compañía : " + Inicio.StatusBar1.Panels(2) + "'"
          Report1.Action = 1
     End Select
End Sub
Private Sub Limpia(Modo As Boolean)
     Toolbar1.Buttons(2).Enabled = Modo
     Toolbar1.Buttons(3).Enabled = Modo
End Sub
Public Function Carga(SQL As String) As Integer
     Dim Temp As Recordset
     Set Temp = DatOS.OpenRecordset(SQL)
     ListView1.ListItems.Clear
     Do Until Temp.EOF
          Set Item = ListView1.ListItems.Add '(, , , , 5)
          Item.Text = Trim(Temp!c_bodega)
          Item.SubItems(1) = Nulo(Temp!d_bodega)
          Item.SubItems(2) = IIf(Temp!Default = 0, "No", "Si")
          Temp.MoveNext
          w% = DoEvents
     Loop
     Call Barra(Trim(ListView1.ListItems.Count) + " registros encontrados", 1)
End Function
