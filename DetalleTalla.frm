VERSION 5.00
Begin VB.Form DetalleTalla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de la Talla"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
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
   ScaleHeight     =   3450
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TTipo 
      Height          =   330
      Left            =   1410
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   3105
      Picture         =   "DetalleTalla.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   585
      Left            =   1860
      Picture         =   "DetalleTalla.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   1410
      TabIndex        =   1
      Top             =   1260
      Width           =   3525
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   1410
      MaxLength       =   15
      TabIndex        =   0
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Talla :"
      Height          =   225
      Left            =   885
      TabIndex        =   5
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   225
      Left            =   675
      TabIndex        =   4
      Top             =   960
      Width           =   675
   End
End
Attribute VB_Name = "DetalleTalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Padre As Form
Dim Texto As TextBox
Dim S$
Dim Item As ListItem
Private Sub Command1_Click()
     Dim Valor%
     Valor = Val(Tag)
     If Valor = 0 Or Valor = 2 Then
          If Inserta Then
               If Valor = 0 Then
                    TTipo = ""
                    Text(0) = ""
                    Text(1) = ""
                    Text(0).SetFocus
               Else
                    Padre.Talla.Requery
                    Texto.Text = Text(0)
                    Unload Me
               End If
          End If
     ElseIf Valor = 1 Then
          If Modifica() Then Unload Me
     End If
End Sub
Private Function Inserta() As Boolean
     S = "insert into tallas(tipotalla,Ctalla,Dtalla)"
     S = S + " values ('" + TTipo + "','" + Text(0) + "','"
     S = S + Text(1) + "')"                  'La descripcion
     Inserta = Procesa(S, 1)
End Function
Private Function Procesa(SQL$, Modo%) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     If Val(Tag) <> 2 Then
          If Modo = 1 Then
               Set Item = Talla.ListView1.ListItems.Add(, , , , 5)
          Else
               Set Item = Talla.ListView1.SelectedItem
          End If
          Item.Text = TTipo
          Item.SubItems(1) = Text(0)
          Item.SubItems(2) = Text(1)
          Item.EnsureVisible
          Item.Selected = True
     End If
     Call GBitacora(Modo, "Tallas:" + Text(0).Text + " " + Text(1).Text)
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
     S = "update tallas Set Ctalla='" + Text(0)
     S = S + "',Dtalla='" + Text(1) + "'"
     S = S + " where Ctalla='" + Text(0).Tag + "'"
     Modifica = Procesa(S, 2)
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Public Sub DatosCliente(Item As ListItem)
     Tag = 1
     TTipo = Item.Text
     Text(0) = Item.SubItems(1)
     Text(0).Tag = Item.SubItems(1)
     Text(1) = Item.SubItems(2)
End Sub

Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 2 Then
          Select Case KeyAscii
          Case 8, 47 To 57
          Case Else
               KeyAscii = 0
          End Select
     End If
End Sub
Public Sub Exterior(Campo As TextBox, Form As Form)
     Tag = 2
     Set Texto = Campo
     Set Padre = Form
End Sub
