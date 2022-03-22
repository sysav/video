VERSION 5.00
Begin VB.Form Entrada 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clave de Ingreso"
   ClientHeight    =   3285
   ClientLeft      =   2580
   ClientTop       =   1575
   ClientWidth     =   6195
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "Entrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Entrada.frx":0442
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3285
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso al Sistema "
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1800
      Left            =   270
      TabIndex        =   3
      Top             =   675
      Width           =   5700
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   0
         Left            =   2355
         TabIndex        =   5
         ToolTipText     =   "Digite su nombre"
         Top             =   600
         Width           =   2190
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2355
         PasswordChar    =   "¤"
         TabIndex        =   4
         ToolTipText     =   "Digite su clave de acceso"
         Top             =   1050
         Width           =   2190
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   4770
         Picture         =   "Entrada.frx":09CC
         Stretch         =   -1  'True
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   435
         TabIndex        =   7
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         TabIndex        =   6
         Top             =   1050
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   555
      Left            =   3225
      MouseIcon       =   "Entrada.frx":0E0E
      MousePointer    =   99  'Custom
      Picture         =   "Entrada.frx":1118
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir del sistema"
      Top             =   2595
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ingresar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   555
      Left            =   2010
      MouseIcon       =   "Entrada.frx":1262
      MousePointer    =   99  'Custom
      Picture         =   "Entrada.frx":156C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ingresar al sistema"
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   405
      TabIndex        =   2
      Top             =   45
      Width           =   5655
   End
End
Attribute VB_Name = "Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S$
Dim Usuario As String
Dim COnta As Integer
Private Sub Command1_Click()
     Select Case Verifica(Text(0), Text(1))
     Case 0
          MousePointer = 11
          w% = DoEvents
          Dim Nom$
          Dim Valor As Boolean
          Dim N%
          Call GBitacora(0, "Ingreso al sistema")
          Dim Accesos As Recordset
          S = "select * from companias where c_compania='" + CiA + "'"
          Set Accesos = DatOS.OpenRecordset(S)
          If Accesos.EOF Then
               CiA = ""
               S = "La compañía seleccionada no existe en el sistema !"
               MsgBox S, 64, "Inventario y Facturación"
          Else
               CaT = Accesos!catalogo
               Inicio.StatusBar1.Panels(2) = Accesos!d_compania
               Inicio.StatusBar1.Panels(2).ToolTipText = "Compañía seleccionada : " + CiA + " " + Accesos!d_compania
               Call ActuParam(CiA)
               'Inicio.Submenu1(11).Visible = True    ' en lugar de series y lotes colores
               'Inicio.Submenu1(9).Visible = True
          End If
          S = "select * from accesos where login='" + LoGiN + "'"
          Set Accesos = DatOS.OpenRecordset(S)
          On Error Resume Next
          For N = 0 To Inicio.Controls.Count - 1 Step 1
               If TypeName(Inicio.Controls(N)) = "Menu" Then
                    Nom = Inicio.Controls(N).Name + Trim(Inicio.Controls(N).Index)
                    Accesos.FindFirst "codigo='" + Nom + "'"
                    If Not Accesos.NoMatch Then
                         If Accesos!Activo = 0 Then
                              Inicio.Controls(N).Enabled = False
                              Inicio.Toolbar1.Buttons(LCase(Accesos!Codigo)).Enabled = False
                         End If
                    End If
               End If
               w% = DoEvents
          Next N
          On Error GoTo 0
          Call SaveKey("LastUser", LoGiN)
          Inicio.StatusBar1.Panels(3) = " " + Usuario + " "
          Inicio.StatusBar1.Panels(3).ToolTipText = "Usuario actual : " + Usuario
          Unload Me
          Inicio.Show
     Case 1
           MsgBox "Usuario no registrado !", 16, "Inventario y Facturación"
           Selecciona Text(0)
     Case 2
           If COnta < 2 Then
                COnta = COnta + 1
                'StatusBar1.Panels(1) = "Ingreso al Sistema (" & Conta & " de 3 )"
                MsgBox "Contraseña incorrecta !", 16, "Inventario y Facturación"
                Selecciona Text(1)
           Else
                StatusBar1.Panels(1) = "Ingreso al Sistema ( 3 de 3 )"
                MsgBox "Ingreso cancelado por tres intentos fallidos.", 16, "Inventario y Facturación"
                Unload Me
           End If
     Case 4
          S = "No es posible verificar su contraseña."
          S = S + Chr(13) + "    Contacte a su administrador."
          MsgBox S, 16, "Error"
          Unload Me
     End Select
End Sub
Private Sub Command2_Click()
     On Error Resume Next
     Acerca.Show 1
     Selecciona Text(1)
     On Error GoTo 0
End Sub
Private Sub Command3_Click()
     End
End Sub
Private Sub Form_Load()
     On Error GoTo WelErr
     Label3.Caption = ReadKey("nomcia")
     Text(0).Text = ReadKey("lastuser")
     TipoVerS = ReadKey("tipovers")
     Show
     Refresh
     COnta = 0
     MousePointer = 0
     Command1.Enabled = True
     Text(1).SetFocus
     If ReadKey("IYFDV") = "11" Then Caption = Caption + " - Versión Demo"
     On Error GoTo 0
WelErr:
     If err.Number = 68 Or err.Number = 3043 Then
          MsgBox "La unidad del sistema no está disponible !", 16, Mid(DirTrA, 1, 3)
          End
     ElseIf err.Number = 3044 Then
          MsgBox "La ruta del sistema no existe !", 16, Mid(DirTrA, 1, 3)
          End
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number)
          End
     End If
     On Error GoTo 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
     On Error Resume Next
End Sub

Private Sub Text_GotFocus(Index As Integer)
     Text(Index).SelStart = 0
     Text(Index).SelLength = Len(Trim(Text(Index).Text))
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case KeyAscii
          Case 13
               Command1_Click
          Case 27
               Unload Me
     End Select
End Sub
Private Function Verifica(LoG As String, Password As String) As Long
     On Error GoTo VErr
     Dim S$
     Dim Temp As Recordset
     S = "select * from usuarios where login='" + LoG + "'"
     Set Temp = DatOS.OpenRecordset(S)
     If Temp.EOF Then
          Verifica = 1
     Else
          If Decripta(Temp!Clave) <> Password Then
               Verifica = 2
          Else
               LoGiN = Trim(Temp!LoGiN)
               Usuario = Trim(Temp!Nombre)
          End If
     End If
VErr:
     If err.Number = 53 Or err.Number = 3024 Then
          Verifica = 4
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Verifica"
          Verifica = err.Number
     End If
     On Error GoTo 0
     Set Temp = Nothing
End Function
Private Sub Limpiador()
     On Error GoTo Errores
     S = InputBox("Digite la clave a este proceso", "Eliminación de la información en la base de datos actual")
     If S = "tortas" Then
          If MsgBox("Desea borrar toda la información de la base de datos", 36, CiA) = 6 Then
               StatusBar1.Panels(1) = "Limpiando base de datos ..."
               MousePointer = 11
               For K = 1 To 10
                    For I = 1 To DatOS.TableDefs.Count - 1
                         If DatOS.TableDefs(I).Connect = "" Then
                              S = "delete from " + DatOS.TableDefs(I).Name
                              DatOS.Execute S, 128
                         End If
                         w% = DoEvents
                    Next
               Next
               StatusBar1.Panels(1) = ""
               MousePointer = 0
          End If
     Else
          MsgBox "Clave incorrecta", 16
     End If
Errores:
     If err.Number = 3200 Or err.Number = 3112 Then
          Resume Next
     ElseIf err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, DatOS.TableDefs(I).Name
          Resume Next
     End If
     On Error GoTo 0
End Sub
