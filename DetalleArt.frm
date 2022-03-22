VERSION 5.00
Begin VB.Form DetalleTipoArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Tipo de Artículo"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DetalleArt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas Contables"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1515
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   6915
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   4
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1020
         Width           =   2625
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   3
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   11
         Top             =   660
         Width           =   2625
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   2
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   7
         Top             =   300
         Width           =   2625
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ingreso :"
         Height          =   225
         Left            =   375
         TabIndex        =   15
         Top             =   1050
         Width           =   705
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   4
         Left            =   3765
         TabIndex        =   14
         Top             =   1020
         Width           =   2985
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   3
         Left            =   3765
         TabIndex        =   12
         Top             =   660
         Width           =   2985
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inventario  :"
         Height          =   225
         Left            =   135
         TabIndex        =   10
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   3765
         TabIndex        =   9
         Top             =   300
         Width           =   2985
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costo :"
         Height          =   225
         Left            =   510
         TabIndex        =   8
         Top             =   690
         Width           =   570
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   3555
      Picture         =   "DetalleArt.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   885
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   585
      Left            =   2325
      Picture         =   "DetalleArt.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   885
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   1410
      TabIndex        =   1
      Top             =   420
      Width           =   5655
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
      Height          =   225
      Left            =   285
      TabIndex        =   5
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   225
      Left            =   660
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "DetalleTipoArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S$
Dim Temp As Recordset
Dim Item As ListItem
Dim Lista1 As ListView
Private Sub Command1_Click()
     If Tag = 1 Then
          S = "insert into tipos(c_tipo_art,d_tipo_art,c_conta,cia,"
          S = S + "ctacostos,ctaventas)"
          S = S + " values ('" + Text(0) + "','" 'El codigo
          S = S + Text(1) + "','" 'La descripcion
          S = S + Text(2) + "','" + CiA + "','" 'El codigo contable
          S = S + Text(3) + "','" + Text(4) + "')"
          If Procesa(S) Then Unload Me
     ElseIf Tag = 2 Then
          S = "update tipos set c_tipo_art='" + Text(0)
          S = S + "',d_tipo_art='" + Text(1)
          S = S + "',c_conta='" + Text(2)
          S = S + "',ctacostos='" + Text(3)
          S = S + "',ctaventas='" + Text(4)
          S = S + "' where cia='" + CiA
          S = S + "' and c_tipo_art='" + Lista1.SelectedItem.Text + "'"
          If Procesa(S) Then Unload Me
     End If
End Sub
Private Function Procesa(SQL As String) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     If Val(Tag) = 1 Then
          Set Item = Lista1.ListItems.Add(, , , , 5)
     ElseIf Val(Tag) = 2 Then
          Set Item = Lista1.SelectedItem
     End If
     Item.Text = Text(0)
     Item.SubItems(1) = Text(1)
     Item.SubItems(2) = Text(2)
     Item.SubItems(3) = Text(3)
     Item.SubItems(4) = Text(4)
     Call GBitacora(Val(Tag), "Tipo de Artículo: " + Text(0) + " " + Text(1))
     Procesa = True
Errores:
     If err.Number = 3022 Then
          MsgBox "Crea una entrada duplicada !", 16, "Error"
          Call Selecciona(Text(0))
     ElseIf err.Number = 3315 Then
          MsgBox "Al menos uno de los campos vacíos es requerido !", 16, "Error"
          Text(0).SetFocus
     ElseIf err.Number = 3075 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
          Text(0).SetFocus
     ElseIf err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Inserta"
     End If
     On Error GoTo 0
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Form_Load()
     Centra Me
     Set Lista1 = TipoArt.ListView1
     S = "select c_cuenta,d_cuenta from detcatalogo,companias "
     S = S + "where companias.c_compania='" + CiA + "' "
     S = S + "and detcatalogo.codigo=companias.catalogo "
     S = S + "and detcatalogo.acepta=1 "
     'Set Temp = DatOS.OpenRecordset(S)
     Refresh
End Sub
Public Sub Carga(Item As ListItem)
     Text(0) = Item.Text
     Text(1) = Item.SubItems(1)
     Text(2) = Item.SubItems(2)
     Text(3) = Item.SubItems(3)
     Text(4) = Item.SubItems(4)
End Sub
Private Sub Text_Change(Index As Integer)
     Dim Valor As Boolean
     Valor = True
     For I = 0 To 1
          If Text(I) = "" Then Valor = False
     Next
     Command1.Enabled = Valor
     If Index >= 2 Then
          Label(Index).Caption = ""
          S = "select c_cuenta,d_cuenta from detcatalogo,companias "
          S = S + "where companias.c_compania='" + CiA + "' "
          S = S + "and detcatalogo.codigo=companias.catalogo "
          S = S + "and detcatalogo.acepta=1 "
          S = S + " and c_cuenta='" + Text(Index) + "'"
          Set Temp = DatOS.OpenRecordset(S)
          If Not Temp.EOF Then Label(Index) = Temp!d_cuenta
     End If
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
     If Index >= 2 Then Call Barra("Presione F3 para buscar por lista")
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case KeyAscii
     Case 39, 44
          KeyAscii = 0
     End Select
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If Index >= 2 And KeyCode = vbKeyF3 Then
          S = "select c_cuenta as Código,d_cuenta as Descripción "
          S = S + "from detcatalogo where codigo=" + CaT
          S = S + " and detcatalogo.acepta=1 order by d_cuenta"
          Call Lista.Carga(Text(Index), S, "Cuentas Contables que aceptan movimientos")
          Lista.Show 1
     End If
End Sub
Private Sub Text_LostFocus(Index As Integer)
     Call Barra("")
End Sub
