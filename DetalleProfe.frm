VERSION 5.00
Begin VB.Form DetalleEsta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Estación"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
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
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   2
      Left            =   1485
      TabIndex        =   6
      Top             =   930
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   2670
      Picture         =   "DetalleProfe.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   615
      Left            =   1470
      Picture         =   "DetalleProfe.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Height          =   345
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   540
      Width           =   3375
   End
   Begin VB.TextBox Text 
      Height          =   345
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio :"
      Height          =   255
      Left            =   255
      TabIndex        =   7
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripción :"
      Height          =   255
      Left            =   315
      TabIndex        =   5
      Top             =   570
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Código :"
      Height          =   225
      Left            =   810
      TabIndex        =   4
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "DetalleEsta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s$
Dim Item As ListItem
Dim Lista As Object
Private Sub Command1_Click()
     If Val(Tag) = 1 Then
          If Inserta Then Unload Me
     ElseIf Val(Tag) = 2 Then
          If Modifica(Lista.SelectedItem.Text) Then Unload Me
     End If
End Sub
Private Function Inserta() As Boolean
     s = "insert into Estaciones(cia,codigo,Descripcion,precio)"
     s = s + " values ('" + CiA + "','" + Text(0).Text + "','" 'El codigo
     s = s + Trim(Text(1).Text) + "'," & Doble(Text(2)) & ")" 'La descripcion precio
     Inserta = Procesa(s, 1)
End Function
Private Function Procesa(SQL As String, Modo As Integer) As Boolean
     On Error GoTo Errores
     DatOS.Execute s, 128
     If Modo = 1 Then
          Set Item = Lista.ListItems.Add()
     ElseIf Modo = 2 Then
          Set Item = Lista.SelectedItem
     End If
     Item.Text = Trim(Text(0).Text)
     Item.SubItems(1) = Trim(Text(1).Text)
     Item.SubItems(2) = Text(2)
     'Item.SubItems(3) = Trim(Text(3).Text)
     Call GBitacora(Modo, "Estacion: " + Item.Text + " " + Item.SubItems(1))
     Set Item = Nothing
     Procesa = True
Errores:
     If err.Number = 3022 Then
          MsgBox "Crea una entrada duplicada !", 16, "Error"
          Call Selecciona(Text(1))
     ElseIf err.Number = 3315 Then
          MsgBox "El código y la descripción son requeridos !", 16, "Error"
          Call Selecciona(Text(1))
     ElseIf err.Number = 3464 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
          Call Selecciona(Text(1))
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Modifica"
     End If
     On Error GoTo 0
End Function
Private Function Modifica(Codigo As String) As Boolean
     s = "update estaciones set codigo='" + Trim(Text(0).Text)
     s = s + "',descripcion='" + Trim(Text(1).Text)
     s = s + "',precio=" & Doble(Text(2).Text)
     s = s + " where cia='" + CiA + "' and codigo='" + Codigo + "'"
     Modifica = Procesa(s, 2)
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Form_Load()
     Set Lista = Estaciones.ListView1
     Refresh
End Sub
Public Sub Carga(Codigo As ListItem)
     Text(0).Text = Codigo.Text
     Text(1).Text = Codigo.SubItems(1)
     Text(2).Text = Codigo.SubItems(2)
     'Text(3).Text = Codigo.SubItems(3)
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 8 Then
          Select Case KeyAscii
          Case 8, 46, 47 To 58
          Case Else
               KeyAscii = 0
          End Select
     End If
End Sub
