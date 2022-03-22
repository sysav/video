VERSION 5.00
Begin VB.Form DetalleBodega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Bodegas"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DetalleBodega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Bodega principal"
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
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   870
      Width           =   1965
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   2310
      Picture         =   "DetalleBodega.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1170
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   585
      Left            =   1080
      Picture         =   "DetalleBodega.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1170
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   990
      TabIndex        =   1
      Top             =   450
      Width           =   3375
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   990
      MaxLength       =   2
      TabIndex        =   0
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nombre :"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   225
      Left            =   225
      TabIndex        =   5
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "DetalleBodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LV As ListView
Dim S$
Dim Item As ListItem
Private Sub Command1_Click()
     If Val(Tag) = 1 Then
          If Inserta Then
               Text(0) = ""
               Text(1) = ""
          End If
     ElseIf Val(Tag) = 2 Then
          If Modifica Then Unload Me
     End If
End Sub
Private Function Inserta() As Boolean
     S = "insert into bodegas(cia,c_bodega,d_bodega,default)"
     S = S + " values ('" + CiA + "','" + Text(0) + "','" 'El codigo
     S = S + Text(1) + "'," 'La descripcion
     S = S & Str(Check1.Value) & ")"  'Default
     Inserta = Procesa(S, 1)
     If Inserta And Check1.Value = 1 Then Call Actualiza(Text(0))
End Function
Private Function Procesa(SQL As String, Modo As Integer) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     If Modo = 1 Then
          Set Item = LV.ListItems.Add
     Else
          Set Item = LV.SelectedItem
     End If
     Item.Text = Text(0)
     Item.SubItems(1) = Text(1)
     Item.SubItems(2) = IIf(Check1.Value = 1, "Si", "No")
     
     Call GBitacora(Modo, "Bodega: " + Item.Text + " " + Item.SubItems(1))
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
     S = "update bodegas set c_bodega='" + Text(0)
     S = S + "',d_bodega='" + Text(1)
     S = S + "',default=" & Check1.Value
     S = S + " where cia='" + CiA + "' and c_bodega='" + Text(0).Tag + "'"
     Modifica = Procesa(S, 2)
     If Modifica And Check1.Value = 1 Then Call Actualiza(Text(0))
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Form_Load()
     Centra Me
     Set LV = Bodegas.ListView1
     Refresh
End Sub
Public Sub DatosCliente(Item As ListItem)
     Text(0).Tag = Item.Text
     Text(0) = Item.Text
     Text(1) = Item.SubItems(1)
     Check1.Value = IIf(Item.SubItems(2) = "Si", 1, 0)

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
Private Sub Actualiza(Codigo As String)
     S = "update bodegas set default=0 where cia='" + CiA
     S = S + "' and c_bodega<>'" + Codigo + "'"
     DatOS.Execute S, 128
     For I = 1 To Bodegas.ListView1.ListItems.Count
          Set Item = Bodegas.ListView1.ListItems(I)
          If Item.Text <> Codigo Then Item.SubItems(2) = "No"
     Next
End Sub
Private Sub Actualiza2(Codigo As String)
     S = "update bodegas set principal=0 where cia='" + CiA
     S = S + "' and c_bodega<>'" + Codigo + "'"
     DatOS.Execute S, 128
     For I = 1 To Bodegas.ListView1.ListItems.Count
          Set Item = Bodegas.ListView1.ListItems(I)
          If Item.Text <> Codigo Then Item.SubItems(3) = "No"
     Next
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 2 And KeyCode = vbKeyF3 Then
        S = "select c_compania,d_compania from companias"
        Call Lista.Carga(Text(2), S, "Compañías")
        Lista.Show 1
    End If
End Sub
