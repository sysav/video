VERSION 5.00
Object = "{B4409115-5405-11D3-943D-0080AD4162AE}#1.0#0"; "ECOMBO.OCX"
Begin VB.Form DetalleProve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Proveedor"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DetalleProve.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
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
      Height          =   885
      Left            =   5130
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   1635
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   510
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin EnhancedCombo.ECombo ECombo1 
      Height          =   345
      Left            =   1290
      TabIndex        =   9
      Top             =   2940
      Width           =   3945
      _ExtentX        =   6959
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
      Index           =   8
      Left            =   3585
      TabIndex        =   3
      Top             =   780
      Width           =   585
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   7
      Left            =   1290
      TabIndex        =   8
      Top             =   2580
      Width           =   3945
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   6
      Left            =   1290
      TabIndex        =   7
      Top             =   2220
      Width           =   3945
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   5
      Left            =   1290
      TabIndex        =   6
      Top             =   1860
      Width           =   1605
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   4
      Left            =   1290
      TabIndex        =   5
      Top             =   1500
      Width           =   1605
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   3
      Left            =   1290
      TabIndex        =   4
      Top             =   1140
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   5580
      Picture         =   "DetalleProve.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2670
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   585
      Left            =   5580
      Picture         =   "DetalleProve.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   1185
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   2
      Left            =   1290
      TabIndex        =   2
      Top             =   780
      Width           =   1605
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   1290
      MaxLength       =   100
      TabIndex        =   1
      Top             =   420
      Width           =   3765
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   0
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo :"
      Height          =   225
      Left            =   750
      TabIndex        =   23
      Top             =   3000
      Width           =   450
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plazo :"
      Height          =   225
      Left            =   3000
      TabIndex        =   22
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cédula :"
      Height          =   225
      Left            =   555
      TabIndex        =   21
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   225
      Left            =   870
      TabIndex        =   20
      Top             =   1920
      Width           =   330
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Contacto :"
      Height          =   225
      Left            =   375
      TabIndex        =   19
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Apartado :"
      Height          =   225
      Left            =   360
      TabIndex        =   18
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Teléfono :"
      Height          =   225
      Left            =   405
      TabIndex        =   17
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dirección :"
      Height          =   225
      Left            =   345
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nombre :"
      Height          =   225
      Left            =   450
      TabIndex        =   15
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   225
      Left            =   525
      TabIndex        =   14
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "DetalleProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s$
Dim Item As ListItem
Dim Lista As ListView
Private Sub Command1_Click()
     Dim NumProv$
     NumProv = Format(Text(0).Text, "0000000000")
     If Tag = 1 Then
          s = "insert into proveedores(codigo,nombre,cedula,direccion,"
          s = s + "telefono,fax,apartado,nom_contact,f_ingreso,compania,"
          s = s + "plazo,c_tipo_pro,moneda)"
          s = s + " values ('" + NumProv + "','" 'El codigo
          s = s + Trim(Text(1).Text) + "','" 'El nombre
          s = s + Trim(Text(2).Text) + "','" 'La cedula
          s = s + Trim(Text(3).Text) + "','" 'La direccion
          s = s + Trim(Text(4).Text) + "','" 'El telefono
          s = s + Trim(Text(5).Text) + "','" 'El fax
          s = s + Trim(Text(6).Text) + "','" 'El apartado
          s = s + Trim(Text(7).Text) + "','" 'El contacto
          s = s + Format(Date, "m/d/yyyy") + "','" + CiA + "'," 'La fecha de ingreso
          s = s + Trim(Val(Text(8).Text)) + ",'" 'El plazo
          s = s + ECombo1.Indice(1) + "','" 'El tipo
          s = s + IIf(Option1, "0", "1") + "')" 'Moneda
          If Procesa(s) Then Unload Me
     ElseIf Tag = 2 Then
          s = "update proveedores set codigo='" + NumProv
          s = s + "',nombre='" + Trim(Text(1).Text)
          s = s + "',cedula='" + Trim(Text(2).Text)
          s = s + "',direccion='" + Trim(Text(3).Text)
          s = s + "',telefono='" + Trim(Text(4).Text)
          s = s + "',fax='" + Trim(Text(5).Text)
          s = s + "',apartado='" + Trim(Text(6).Text)
          s = s + "',nom_contact='" + Trim(Text(7).Text)
          s = s + "',plazo=" + Trim(Val(Text(8).Text))
          s = s + ",c_tipo_pro='" + ECombo1.Indice(1)  'El tipo
          s = s + "',moneda='" + IIf(Option1, "0", "1") 'Moneda
          s = s + "' where codigo='" + Lista.SelectedItem.Text
          s = s + "' and compania='" + CiA + "'"
          If Procesa(s) Then Unload Me
     End If
End Sub
Private Function Procesa(SQL As String) As Boolean
     On Error GoTo Errores
     DatOS.Execute SQL, 128
     If Tag = "1" Then
          Set Item = Lista.ListItems.Add(, , , , 5)
     ElseIf Tag = 2 Then
          Set Item = Lista.SelectedItem
     End If
     Item.Text = Trim(Text(0).Text)
     Item.SubItems(1) = Trim(Text(1).Text)
     Call GBitacora(Val(Tag), "Proveedor: " + Text(0).Text + " " + Text(1).Text)
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
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Inserta"
     End If
     On Error GoTo 0
End Function
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Form_Load()
     Dim Temp As Recordset
     Option1.Caption = MoNLoC
     Option2.Caption = MoNExT
     Set Lista = Provedores.ListView1
     s = "select * from tipoprov where compania='" + CiA + "'"
     Set Temp = DatOS.OpenRecordset(s)
     With Temp
     Do Until Temp.EOF
          ECombo1.AddItem Temp!d_tipo_pro, !c_tipo_pro
          Temp.MoveNext
     Loop
     End With
     Refresh
End Sub
Public Sub Carga(Registro As Recordset)
     On Error Resume Next
     Text(0) = Registro!Codigo
     Text(1) = Registro!Nombre
     Text(2) = Registro!Cedula
     Text(3) = Registro!Direccion
     Text(4) = Registro!Telefono
     Text(5) = Registro!Fax
     Text(6) = Registro!Apartado
     Text(7) = Registro!nom_contact
     Text(8) = Registro!Plazo
     For I = 0 To ECombo1.ListCount - 1
          If ECombo1.List(I, 1) = Registro!c_tipo_pro Then
               ECombo1.ListIndex = I
               Exit For
          End If
     Next
     If Registro!Moneda = "0" Then Option1 = True Else Option2 = True
     On Error GoTo 0
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index <> 8 Then
          Select Case KeyAscii
          Case 39, 44, 96
               KeyAscii = 0
          End Select
     Else
          Select Case KeyAscii
          Case 8, 48 To 57
          Case Else
               KeyAscii = 0
          End Select
     End If
End Sub
Private Sub Text_LostFocus(Index As Integer)
     If Index = 0 Then Text(0).Text = Format(Text(0).Text, "0000000000")
End Sub
