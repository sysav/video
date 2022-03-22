VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Etiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Etiquetas"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Etiquetas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   6495
   Begin VB.OptionButton Option2 
      Caption         =   "Montos en Dólares"
      Height          =   285
      Left            =   2910
      TabIndex        =   17
      Top             =   1500
      Width           =   1995
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Montos en Colones"
      Height          =   285
      Left            =   2910
      TabIndex        =   16
      Top             =   1230
      Value           =   -1  'True
      Width           =   1995
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   3
      Left            =   4260
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   810
      Width           =   495
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   0
      Left            =   1230
      MaxLength       =   15
      TabIndex        =   0
      Top             =   90
      Width           =   1875
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   1
      Left            =   1230
      MaxLength       =   15
      TabIndex        =   1
      Top             =   450
      Width           =   1875
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   2
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   960
      Width           =   345
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   825
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   180
      Top             =   3180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   615
      Left            =   5220
      Picture         =   "Etiquetas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir el reporte seleccionado"
      Top             =   840
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   615
      Left            =   5220
      Picture         =   "Etiquetas.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar esta ventana"
      Top             =   1500
      Width           =   1185
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2295
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "Text3"
      BuddyDispid     =   196612
      OrigLeft        =   2475
      OrigTop         =   1170
      OrigRight       =   2715
      OrigBottom      =   1515
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown2 
      Height          =   330
      Left            =   4785
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   810
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "Text(3)"
      BuddyDispid     =   196611
      BuddyIndex      =   3
      OrigLeft        =   2475
      OrigTop         =   1170
      OrigRight       =   2715
      OrigBottom      =   1515
      Max             =   18
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empezar en etiqueta # :"
      Height          =   225
      Left            =   2280
      TabIndex        =   14
      Top             =   870
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Del artículo :"
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   150
      Width           =   1005
   End
   Begin VB.Label NomArt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   90
      Width           =   3285
   End
   Begin VB.Label NomArt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   1
      Left            =   3120
      TabIndex        =   11
      Top             =   450
      Width           =   3285
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Al artículo :"
      Height          =   225
      Left            =   270
      TabIndex        =   10
      Top             =   510
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo de precio :"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   1245
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Copias :"
      Height          =   225
      Left            =   720
      TabIndex        =   8
      Top             =   1380
      Width           =   645
   End
End
Attribute VB_Name = "Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumEnt$
Dim CodBod$
Dim Entrada As Boolean
Dim TipoC#
Dim SimBOlo As String * 1
Dim LabelTop%
Dim Articulos As Recordset
Dim S$
Private Function ImpreEtiq(Lado%) As Boolean
     On Error GoTo CMDErr
     Dim Start%
     If Lado = 0 Then
          Start = 4
     Else
          Start = 111
     End If
     With Articulos
     'El codigo
     'Printer.ForeColor = RGB(100, 200, 200)
     Printer.FontBold = False
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 5
     Printer.Print "Modelo : " + !c_articulo
     'La marca
     Printer.CurrentX = Start + 55
     Printer.CurrentY = LabelTop + 5
     Printer.Print "Marca : " + IIf(IsNull(!Marca), "", !Marca)
     'La descripcion
     Printer.FontBold = True
     Printer.CurrentX = Start + (48 - Len(!d_articulo))
     Printer.CurrentY = LabelTop + 10
     Printer.Print Mid(Articulos!d_articulo, 1, 48)
     'Precio
     Printer.FontBold = False
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 15
     Printer.Print "Precio Neto"
     'El Precio
     Printer.FontBold = False
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 20
     Printer.Print SimBOlo + " " + FormatNumber(!monto * TipoC, DeCiMaleS)
     'Leyenda de precio con iv
     Printer.FontBold = False
     Printer.CurrentX = Start + 55
     Printer.CurrentY = LabelTop + 15
     Printer.Print IIf(!sino_impu = 1, "Precio con IV", "Exento")
     'El precio con iv
     Printer.FontBold = False
     Printer.CurrentX = Start + 55
     Printer.CurrentY = LabelTop + 20
     Printer.Print SimBOlo + " " + FormatNumber((!monto * TipoC) * IIf(!sino_impu = 1, (IV + 1), 1), DeCiMaleS)
     ImpreEtiq = True
     End With
CMDErr:
     If err.Number = 482 Then
          S = "La impresora no está lista!, verifique que esté encendida y el papel colocado."
          If MsgBox(S, 21, Printer.DeviceName) = 4 Then
               Resume
          Else
               Exit Function
          End If
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Impresion"
     End If
     On Error GoTo 0
End Function
Private Sub Command1_Click()
     If Text(2) = "" Then
          MsgBox "El tipo de precio es requerido !", 48, Caption
          Text(2).SetFocus
          Exit Sub
     End If
     If Val(Text(3)) = 0 Or Val(Text(3)) > 20 Then
          MsgBox "Número de etiqueta no válido !", 48, Caption
          Selecciona Text(3)
          Exit Sub
     End If
     MousePointer = 11
     If Entrada Then
          S = "select desgent.c_articulo,articulos.d_articulo,"
          S = S + "articulos.marca,listaprecios.monto,articulos.sino_impu,"
          S = S + "desgent.cantidad "
          S = S + "from articulos,listaprecios,desgent "
          S = S + "where desgent.cia='" + CiA
          S = S + "' and desgent.bodega='" + CodBod
          S = S + "' and desgent.n_entrada='" + NumEnt
          S = S + "' and listaprecios.cia=desgent.cia "
          S = S + "  and articulos.cia=desgent.cia "
          S = S + "  and articulos.c_articulo=desgent.c_articulo "
          S = S + "  and listaprecios.codart=desgent.c_articulo "
          S = S + "  and listaprecios.codigo=" + Text(2)
          S = S + " order by desgent.c_articulo"
     Else
          S = "select c_articulo,d_articulo,marca,monto,sino_impu "
          S = S + "from articulos,listaprecios "
          S = S + "where listaprecios.cia='" + CiA
          S = S + "' and listaprecios.cia=articulos.cia "
          S = S + "  and listaprecios.codart=articulos.c_articulo "
          S = S + "  and listaprecios.codigo=" + Text(2)
          If NomArt(0) <> "" Then
               S = S + " and c_articulo >='" + Text(0) + "'"
               If NomArt(1) <> "" Then
                    S = S + " and c_articulo <='" + Text(1) + "'"
               End If
          End If
          S = S + " order by c_articulo"
     End If
     Dim Pag%
     Dim I%
     Pag = Val(Text(3))
     LabelTop = 12
     Select Case Pag
     Case 1, 2
          I = IIf(Pag = 1, 0, 1)
     Case 3, 4
          I = IIf(Pag = 3, 0, 1)
          LabelTop = LabelTop + 26
     Case 5, 6
          I = IIf(Pag = 5, 0, 1)
          LabelTop = LabelTop + 50
     Case 7, 8
          I = IIf(Pag = 7, 0, 1)
          LabelTop = LabelTop + 75
     Case 9, 10
          I = IIf(Pag = 9, 0, 1)
          LabelTop = LabelTop + 100
     Case 11, 12
          I = IIf(Pag = 11, 0, 1)
          LabelTop = LabelTop + 125
     Case 13, 14
          I = IIf(Pag = 13, 0, 1)
          LabelTop = LabelTop + 150
     Case 15, 16
          I = IIf(Pag = 15, 0, 1)
          LabelTop = LabelTop + 175
     Case 17, 18
          I = IIf(Pag = 17, 0, 1)
          LabelTop = LabelTop + 200
     Case 19, 20
          I = IIf(Pag = 19, 0, 1)
          LabelTop = LabelTop + 225
     End Select
     Set Articulos = DatOS.OpenRecordset(S)
     If Articulos.EOF Then
          S = "No se  encontraron  artículos con"
          S = S + Chr(13) + "precio en el rango seleccionado !"
          MsgBox S, 64, Caption
     Else
          TipoC = IIf(Option1, TCambio, 1)
          Dim Cuantos%
          Dim Actual%
          If Entrada Then
               Do Until Articulos.EOF
                    Cuantos = Cuantos + Articulos!Cantidad
                    Articulos.MoveNext
               Loop
               Articulos.MoveFirst
          Else
               Articulos.MoveLast
               Cuantos = Articulos.RecordCount
               Articulos.MoveFirst
          End If
          Printer.ScaleMode = vbMillimeters
          Printer.Font = "Courier New"
          Printer.FontSize = 10
          Printer.Print ""
          If Not Entrada Then
               Do Until Articulos.EOF
                    Actual = Actual + 1
                    Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                    If Pag > 20 And Cuantos > 20 Then
                         Pag = 1
                         Printer.NewPage
                         LabelTop = 12
                    End If
                    Call ImpreEtiq(I)
                    Articulos.MoveNext
                    Pag = Pag + 1
                    If I = 0 Then
                         I = 1
                    Else
                         LabelTop = LabelTop + 25
                         I = 0
                    End If
               Loop
          Else
               Do Until Articulos.EOF
                    For K = 1 To 1 'Articulos!Cantidad
                         Actual = Actual + 1
                         Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                         If Pag > 20 And Cuantos > 20 Then
                              Pag = 1
                              Printer.NewPage
                              LabelTop = 12
                         End If
                         Call ImpreEtiq(I)
                         Pag = Pag + 1
                         If I = 0 Then
                              I = 1
                         Else
                              LabelTop = LabelTop + 25
                              I = 0
                         End If
                    Next
                    Articulos.MoveNext
               Loop
          End If
          Printer.EndDoc
     End If
     MousePointer = 0
     Call Barra("")
     MsgBox "Proceso Finalizado", 64, Caption
     If Entrada Then Unload Me
End Sub
Private Sub Command2_Click()
     Unload Me
End Sub
Private Sub Form_Load()
     Centra Me
     Text3.Text = 1
     Dim SysParam As Recordset
     S = "select * from parametros where cia='" + CiA + "'"
     Set SysParam = DatOS.OpenRecordset(S, 4)
     SimBOlo = Nulo(SysParam!SimBOlo)
     Show
     Refresh
End Sub
Private Sub Text_Change(Index As Integer)
     Select Case Index
     Case 0, 1
          NomArt(Index).Caption = ""
          S = "select c_articulo,d_articulo from articulos "
          S = S + "where cia='" + CiA + "' and c_articulo='" + Text(Index) + "'"
          Set Articulos = DatOS.OpenRecordset(S)
          If Not Articulos.EOF Then NomArt(Index).Caption = Articulos!d_articulo
     End Select
End Sub
Private Sub Text_GotFocus(Index As Integer)
     Call Barra("Presione F3 para buscar por lista")
End Sub
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 2 Or Index = 3 Then
          Select Case KeyAscii
          Case 8, 48 To 57
          Case Else
               KeyAscii = 0
          End Select
     End If
End Sub
Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF3 Then
          Dim Titulo$
          Select Case Index
          Case 0, 1
               Titulo = "Artículos"
               S = "select c_articulo as Código,d_articulo as Descripción "
               S = S + "from articulos where cia='" + CiA + "' order by d_Articulo"
          End Select
          Call Lista.Carga(Text(Index), S, Titulo)
          Lista.Show 1
     End If
End Sub
Private Sub Text_LostFocus(Index As Integer)
     Call Barra("")
End Sub
Public Sub DesdeEntradas(Numero$, Bodega$)
     Entrada = True
     Text(0).Enabled = False
     Text(1).Enabled = False
     Label2.Enabled = False
     Label3.Enabled = False
     NumEnt = Numero
     CodBod = Bodega
     Caption = Caption + " (Entrada No. '" + Numero + "')"
End Sub
