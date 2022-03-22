VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Estadisticas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Estadisticas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   8160
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   1500
      MaxLength       =   15
      TabIndex        =   2
      Top             =   390
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   1500
      MaxLength       =   15
      TabIndex        =   3
      Top             =   750
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Default         =   -1  'True
      Height          =   345
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   405
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   3210
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   1545
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   1695
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4425
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1110
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   7805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Text            =   "Año"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Mes"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cantidad"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Artículo :"
      Height          =   225
      Left            =   720
      TabIndex        =   10
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3210
      TabIndex        =   9
      Top             =   390
      Width           =   3525
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      Height          =   225
      Left            =   810
      TabIndex        =   8
      Top             =   450
      Width           =   645
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6900
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Estadisticas.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3210
      TabIndex        =   7
      Top             =   750
      Width           =   3525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mostrar desde :"
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   90
      Width           =   1245
   End
End
Attribute VB_Name = "Estadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Indice%
Dim Articulo As Recordset
Dim otro As Recordset
Dim Item As ListItem
Dim S$
Private Sub Command1_Click()
     If Text(0) = "" Then
          MsgBox "El artículo es requerido !", 48, "Estadisticas"
          Selecciona Text(0)
          Exit Sub
     End If
     Dim MeS$
     Dim Year$
     Year = Format(Date, "yyyy")
     MeS = Format(Date, "mm")
     ListView1.ListItems.Clear
     If Indice = 6 Then 'Compras
          S = "select fecha,unidades,monto from estcompras where "
          If Text(1) <> "" Then S = S + "codprov='" + Text(1) + "' and "
          S = S + "codart='" + Text(0)
          S = S + "' and fecha >='" + Combo2.Text
          S = S + Format(Combo1.ListIndex + 1, "00")
          S = S + "' and fecha <='" + Year + MeS
          S = S + "' order by fecha"
     ElseIf Indice = 5 Then 'Ventas
          S = "select fecha,unidades,monto from estadisticas where "
          If Text(1) <> "" Then S = S + "cliente='" + Text(1) + "' and "
          S = S + "codart='" + Text(0)
          S = S + "' and fecha >='" + Combo2.Text + Format(Combo1.ListIndex + 1, "00")
          S = S + "' and fecha <='" + Year + MeS + "'"
          S = S + " order by fecha"
     End If
     Dim Temp As Recordset
     Set Temp = DatOS.OpenRecordset(S)
     If Temp.EOF Then
          S = "No se encontró ninguna estadistica en el rango seleccionado,"
          S = S + Chr(13) + "             Desea buscar un año anterior ?"
          If MsgBox(S, 36, "Estadisticas") = 6 Then
               Combo2.ListIndex = IIf(Combo2.ListIndex + 1 < Combo2.ListCount, Combo2.ListIndex + 1, 0)
               Command1_Click
          End If
     Else
          Do Until Temp.EOF
               Call AgregaItem(Temp)
               Temp.MoveNext
               w% = DoEvents
          Loop
     End If
End Sub
Private Sub Form_Load()
     Call Posiciona(Me, 0)
     Show
     Refresh
     Dim I%
     For I = 1 To 12
          Combo1.AddItem Meses(I)
     Next
     Dim Year%
     Year = Format(Date, "yyyy")
     For I = Year To Year - 10 Step -1
          Combo2.AddItem I
     Next
     S = "select d_articulo,c_articulo from articulos where cia='" + CiA + "'"
     Set Articulo = DatOS.OpenRecordset(S, dbOpenSnapshot)
     Combo1.ListIndex = Val(Format(Date, "mm")) - 1
     Combo2.ListIndex = 0
End Sub
Private Sub Carga()
     Dim monto&
     Dim Unidades%
     For K = 1997 To 1999
          For I = 1 To 12
               For J = 1 To 20
                    monto = monto + 100 + J + I
                    Unidades = Unidades + J + I
                    S = "insert into estadisticas(codart,ano,mes,monto,unidades,cia)"
                    S = S + " values ('01-01-001'," 'El articulo
                    S = S & K & "," 'El año
                    S = S & I & ","   'El mes
                    S = S & monto & "," 'El monto
                    S = S & Unidades & ",'" + CiA + "')" 'La cantidad
                    DatOS.Execute S, 128
               Next
               monto = 0
               Unidades = 0
          Next
     Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Call Posiciona(Me, 1)
     Call Barra("")
End Sub
Private Sub AgregaItem(Tabla As Recordset)
     Set Item = ListView1.ListItems.Add(, , , , 1)
     Item.Text = Mid(Tabla!Fecha, 1, 4)
     Item.SubItems(1) = Meses(Val(Mid(Tabla!Fecha, 5)))
     Item.SubItems(2) = Tabla!Unidades
     Item.SubItems(3) = Format(Tabla!monto, "standard")
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
     ListView1.SortKey = ColumnHeader.Index - 1
     ListView1.SortOrder = IIf(ListView1.SortOrder = 0, 1, 0)
     ListView1.Sorted = True
End Sub
Private Sub Text_Change(Index As Integer)
     ListView1.ListItems.Clear
     If Index = 1 Then
          Label5.Caption = ""
          otro.FindFirst otro(0).Name + "='" + Text(1) + "'"
          If Not otro.NoMatch Then
               Label5.Caption = otro(1)
          End If
     Else
          Label3.Caption = ""
          Articulo.FindFirst "c_articulo='" + Text(0) + "'"
          If Not Articulo.NoMatch Then Label3.Caption = Articulo!d_articulo
     End If
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Call Barra("Presione F3 para buscar por lista")
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF3 Then
          Dim Titulo$
          If Index = 1 Then
               If Indice = 5 Then
                    S = "select codigo as Código,Nombre "
                    S = S + "from clientes where cia='" + CiA + "'  order by nombre"
                    Titulo = "Clientes"
               ElseIf Indice = 6 Then
                    S = "select codigo as Código,Nombre "
                    S = S + "from proveedores where compania='" + CiA
                    S = S + "' order by nombre"
                    Titulo = "Proveedores"
               End If
          Else
               Titulo = "Artículos"
               S = "select c_articulo as Código,d_articulo as Descripción "
               S = S + "from articulos where cia='" + CiA + "' order by d_articulo"
          End If
          Call Lista.Carga(Text(Index), S, Titulo)
          Lista.Show 1
     End If
End Sub
Private Sub Text_LostFocus(Index As Integer)
     Call Barra("")
End Sub
Public Sub Init(Index As Integer)
     Indice = Index
     If Index = 5 Then
          Caption = "Estadistica de Ventas"
          ListView1.ColumnHeaders(4).Text = "Monto Vendido"
          Label4.Caption = "Cliente : "
          S = "select codigo as Código,Nombre from clientes where cia='" + CiA + "'"
          Label3.Tag = "codigo"
          Label4.Tag = "codigo"
     ElseIf Index = 5 Or Index = 6 Then
          Caption = "Estadistica de Compras"
          ListView1.ColumnHeaders(4).Text = "Monto Comprado"
          Label4.Caption = "Proveedor : "
          Label3.Tag = "c_prove"
          Label4.Tag = "codigo"
          S = "select codigo,nombre from proveedores where compania='" + CiA + "'"
     End If
     Set otro = DatOS.OpenRecordset(S, dbOpenSnapshot)
End Sub
