VERSION 5.00
Begin VB.Form DetalleUsos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Usos Contables"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DetalleUsos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   2
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   2
      Top             =   780
      Width           =   3225
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   2520
      Picture         =   "DetalleUsos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   585
      Left            =   1290
      Picture         =   "DetalleUsos.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   1
      Top             =   420
      Width           =   3225
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código contable :"
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   870
      Width           =   1410
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nombre :"
      Height          =   225
      Left            =   780
      TabIndex        =   6
      Top             =   510
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   225
      Left            =   885
      TabIndex        =   5
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "DetalleUsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s$
Dim Item As ListItem
Private Sub Command1_Click()
     If Val(Tag) = 1 Then
          If Inserta Then
               Text(1) = ""
               Text(2) = ""
               Text(1).SetFocus
          End If
     ElseIf Val(Tag) = 2 Then
          If Modifica(Usos.ListView1.SelectedItem.Text) Then Unload Me
     End If
End Sub
Private Function Inserta() As Boolean
     s = "insert into usos(codigo,descripcion,contable,cia)"
     s = s + " values ('" + Text(0) + "','" 'El codigo
     s = s + Text(1) + "','" 'La descripcion
     s = s + Text(2) + "','" + CiA + "')" 'El contable
     Inserta = Procesa(s, 1)
End Function
Private Function Procesa(SQL As String, Modo As Integer) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     If Modo = 1 Then
          Set Item = Usos.ListView1.ListItems.Add()
     Else
          Set Item = Usos.ListView1.SelectedItem
     End If
     Item.Text = Text(0)
     Item.SubItems(1) = Text(1)
     Item.SubItems(2) = Text(2)
     Call GBitacora(Modo, "Uso : " + Text(0) + " " + Text(1))
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
Private Function Modifica(Codigo As String) As Boolean
     s = "update usos set codigo='" + Text(0)
     s = s + "',descripcion='" + Text(1)
     s = s + "',contable='" + Text(2)
     s = s + "' where cia='" + CiA + "' and codigo='" + Codigo + "'"
     Modifica = Procesa(s, 2)
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Form_Load()
     Refresh
End Sub
Public Sub DatosCliente(Item As ListItem)
     Text(0).Text = Item.Text
     Text(1).Text = Item.SubItems(1)
     Text(2).Text = Item.SubItems(2)
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
     ElseIf Index = 1 Then
          Select Case KeyAscii
          Case 34, 39
               KeyAscii = 0
          End Select
     End If
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If Index = 2 And KeyCode = vbKeyF3 Then
          s = "select c_cuenta as Código,d_cuenta as Descripción "
          s = s + "from detcatalogo,companias "
          s = s + "where c_compania='" + CiA + "' "
          s = s + "and codigo=catalogo and acepta=1 order by d_cuenta"
          Call Lista.Carga(Text(2), s, "Cuentas Contables")
          Lista.Show 1
     End If
End Sub
