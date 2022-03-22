VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "BARCOD32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Barras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Códigos de Barras"
   ClientHeight    =   2475
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
   Icon            =   "Barras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2640
      TabIndex        =   17
      Top             =   840
      Width           =   2415
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   5
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "1"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   4
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "1"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   1920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Text(4)"
         BuddyDispid     =   196610
         BuddyIndex      =   4
         OrigLeft        =   2475
         OrigTop         =   1170
         OrigRight       =   2715
         OrigBottom      =   1515
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown3 
         Height          =   330
         Left            =   1920
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Text(5)"
         BuddyDispid     =   196610
         BuddyIndex      =   5
         OrigLeft        =   2475
         OrigTop         =   1170
         OrigRight       =   2715
         OrigBottom      =   1515
         Max             =   3
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Columna :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   390
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fila :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin BarcodLib.Barcod Barcode1 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2145
      _Version        =   65543
      _ExtentX        =   3784
      _ExtentY        =   556
      _StockProps     =   75
      BackColor       =   -2147483639
      BarWidth        =   0
      Direction       =   0
      Style           =   18
      UPCNotches      =   3
      Alignment       =   0
      Extension       =   ""
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Montos en Dólares"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3150
      TabIndex        =   15
      Top             =   2100
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Montos en Colones"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3150
      TabIndex        =   14
      Top             =   1830
      Value           =   -1  'True
      Width           =   1995
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   3
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
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
      Left            =   120
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   615
      Left            =   5205
      Picture         =   "Barras.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir el reporte seleccionado"
      Top             =   855
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   615
      Left            =   5220
      Picture         =   "Barras.frx":0544
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
      BuddyDispid     =   196615
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
Attribute VB_Name = "Barras"
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
Dim s$
Dim NumEti As Integer
Private Function ImpreBarra(Lado%) As Boolean
     On Error GoTo CMDErr
     Dim Start%, Xx$
     Start = 1
     LabelTop = 3
With Articulos
     'La descripcion
     Printer.FontSize = 7
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 1
     Xx = Mid(Articulos!d_articulo, 1, 48)
     Printer.Print Xx
     
     'El precio con iv
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 8
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC * IIf(!sino_impu = 1, (IV + 1), 1), "###,###") + " IVI"
     Printer.Print Xx
     Printer.FontBold = False
     
    'Etiqueta de la Izquierda
     Barcode1.PrinterLeft = 2        'Start
     Barcode1.PrinterTop = 15        'LabelTop
     Barcode1.PrinterWidth = 30
     Barcode1.PrinterHeight = 8
     Barcode1.Caption = "*"
     Barcode1.Caption = Articulos!c_articulo
     Barcode1.PrinterHDC = Printer.hDC
     ' La Barra
     Printer.FontSize = 8
     Printer.CurrentX = Barcode1.PrinterLeft
     Printer.CurrentY = LabelTop + 5 ' Barcode1.PrinterTop + Barcode1.PrinterHeight + 2
     Printer.Print Barcode1.Displayed
     
     'El Precio
     Printer.FontSize = 8
     Printer.CurrentX = Start + 18
     Printer.CurrentY = LabelTop
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC, "###,###")
     'Printer.Print Xx
     
     
     'La fecha de hoy
     Printer.FontSize = 6
     Printer.FontBold = False
     Printer.CurrentX = Start + 37
     Printer.CurrentY = 16
     'Printer.Print Format(Date, "mm/yy") + " "

     Printer.FontBold = False
     Printer.FontSize = 10
     'Printer.NewPage
     'Printer.EndDoc
     ImpreBarra = True
     End With
CMDErr:
     If err.Number = 482 Then
          s = "La impresora no está lista!, verifique que esté encendida y el papel colocado."
          If MsgBox(s, 21, Printer.DeviceName) = 4 Then
               Resume
          Else
               Exit Function
          End If
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
     End If
     On Error GoTo 0
End Function
Private Function ImpreBarra2(Lado%) As Boolean
     On Error GoTo CMDErr
     Dim Start%, Xx$
     Start = 40
     LabelTop = 3
With Articulos
     'La descripcion
     Printer.FontSize = 7
     Printer.CurrentX = Start
     Printer.CurrentY = LabelTop + 1
     Xx = Mid(Articulos!d_articulo, 1, 48)
     Printer.Print Xx
     
     'El precio con iv
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start
     Printer.CurrentY = LabelTop + 8
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC * IIf(!sino_impu = 1, (IV + 1), 1), "###,###") + " IVI"
     Printer.Print Xx
     Printer.FontBold = False
     
    'Etiqueta de la Derecha
     Barcode1.PrinterLeft = 40        'Start
     Barcode1.PrinterTop = 15       'LabelTop
     Barcode1.PrinterWidth = 30
     Barcode1.PrinterHeight = 8
     Barcode1.Caption = "*"
     Barcode1.Caption = Articulos!c_articulo
     Barcode1.PrinterHDC = Printer.hDC
     ' La Barra
     Printer.FontSize = 8
     Printer.CurrentX = Barcode1.PrinterLeft
     Printer.CurrentY = LabelTop + 5   ' Barcode1.PrinterTop + Barcode1.PrinterHeight + 2
     Printer.Print Barcode1.Displayed
     'El Precio
     Printer.FontSize = 8
     Printer.CurrentX = Start + 8
     Printer.CurrentY = LabelTop
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC, "###,###")
     'Printer.Print Xx

     'La fecha de hoy
     Printer.FontSize = 6
     Printer.FontBold = False
     Printer.CurrentX = Start + 51
     Printer.CurrentY = 16
     'Printer.Print Format(Date, "mm/yy") + " "

     Printer.FontBold = False
     Printer.FontSize = 10
     'Printer.NewPage
     Printer.EndDoc
     ImpreBarra2 = True
     End With
CMDErr:
     If err.Number = 482 Then
          s = "La impresora no está lista!, verifique que esté encendida y el papel colocado."
          If MsgBox(s, 21, Printer.DeviceName) = 4 Then
               Resume
          Else
               Exit Function
          End If
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
     End If
     On Error GoTo 0
End Function
Private Sub Command1_Click()
Dim Pr As Printer
  s = ReadKey("PrinterSeries")
  For Each Pr In Printers
      If LCase(Pr.DeviceName) = LCase(s) Then
         Set Printer = Pr
         Exit For
      End If
  Next
  
     If Text(2) = "" Then
          MsgBox "El tipo de precio es requerido !", 48, Caption
          Text(2).SetFocus
          Exit Sub
     End If
     If Val(Text(3)) = 0 Or Val(Text(3)) > 30 Then
          MsgBox "Número de etiqueta no válido !", 48, Caption
          Selecciona Text(3)
          Exit Sub
     End If
     MousePointer = 11
     If NomArt(0).Caption <> "" Then Entrada = False
     If Entrada Then
          s = "select desgent.c_articulo,articulos.d_articulo,ALTERNO,ubicacion, "
          s = s + "articulos.marca,listaprecios.monto,articulos.sino_impu,"
          s = s + "desgent.cantidad "
          s = s + "from articulos,listaprecios,desgent "
          s = s + "where desgent.cia='" + CiA
          s = s + "' and desgent.bodega='" + CodBod
          s = s + "' and desgent.n_entrada='" + NumEnt
          s = s + "' and listaprecios.cia=desgent.cia "
          s = s + "  and articulos.cia=desgent.cia "
          s = s + "  and articulos.c_articulo=desgent.c_articulo "
          s = s + "  and listaprecios.codart=desgent.c_articulo "
          s = s + "  and listaprecios.codigo=" + Text(2)
          s = s + " order by desgent.c_articulo"
     Else
          s = "select c_articulo,d_articulo,marca,monto,sino_impu,ALTERNO,ubicacion "
          s = s + " from articulos,listaprecios "
          s = s + " where articulos.cia='" + CiA + "' and listaprecios.cia='" + CiA
          s = s + "' and listaprecios.cia=articulos.cia "
          s = s + "  and listaprecios.codart=articulos.c_articulo "
          s = s + "  and listaprecios.codigo=" + Text(2)
          If NomArt(0) <> "" Then
               If NomArt(1) <> "" Then
                  s = s + " and c_articulo  between '" + Text(0) + "'"
                  s = s + " and '" + Text(1) + "'"
               Else
                  s = s + " and c_articulo ='" + Text(0) + "'"
               End If
          End If
          s = s + " order by c_articulo"
     End If
'=====
     Dim Pag%
     Dim I%
     Dim Eti%
     Pag = Val(Text(3))
     LabelTop = 2
     Select Case Pag
     Case 1, 2, 3
          If Pag = 1 Then
            I = 0
          ElseIf Pag = 2 Then
            I = 1
          Else
            I = 2
          End If
     End Select
'======
     Set Articulos = DatOS.OpenRecordset(s)
     If Articulos.EOF Then
          s = "No se  encontraron  artículos con"
          s = s + Chr(13) + "precio en el rango seleccionado !"
          MsgBox s, 64, Caption
     Else
          TipoC = IIf(Option1, 1, TCambio)      ' COLONES
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
          Dim LasCopias As Double
          Printer.ScaleMode = vbMillimeters
          If Not Entrada Then
               Do Until Articulos.EOF
               For x = 1 To Val(Text3.Text)
                    Actual = Actual + 1
                    'Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                    'If Pag > 30 And Cuantos > 30 Then
                    'Imprime la Etiqueta
                    If I = 0 Then
                        Printer.ScaleMode = vbMillimeters
                        Printer.Font = "Tahoma" '"Courier New"
                        Printer.FontSize = 8
                        Printer.Print ""
                        Barcode1.PrinterScaleMode = Printer.ScaleMode
                        LabelTop = 12
                        Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                        Call ImpreEtiq(I)
                        NumEti = NumEti + 1
                    ElseIf I = 1 Then
                        Printer.ScaleMode = vbMillimeters
                        Printer.Font = "Tahoma" '"Courier New"
                        Printer.FontSize = 8
                        Printer.Print ""
                        Barcode1.PrinterScaleMode = Printer.ScaleMode
                        Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                        Call ImpreEtiq2(I)
                    Else
                        Printer.ScaleMode = vbMillimeters
                        Printer.Font = "Tahoma" '"Courier New"
                        Printer.FontSize = 8
                        Printer.Print ""
                        Barcode1.PrinterScaleMode = Printer.ScaleMode
                        Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                        Call ImpreEtiq3(I)
                    End If
                    
                    Pag = Pag + 1
                    If I = 0 Then
                         I = 1
                    ElseIf I = 1 Then
                        I = 2
                    Else
                         I = 0
                         Printer.NewPage
                    End If
               Next
               Articulos.MoveNext
               Loop
          Else
               Do Until Articulos.EOF
               For x = 1 To Val(Articulos!Cantidad)
                    Actual = Actual + 1
                    'Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                    'If Pag > 30 And Cuantos > 30 Then
                    'Imprime la Etiqueta
                    If I = 0 Then
                        Printer.ScaleMode = vbMillimeters
                        Printer.Font = "Tahoma" '"Courier New"
                        Printer.FontSize = 8
                        Printer.Print ""
                        Barcode1.PrinterScaleMode = Printer.ScaleMode
                        LabelTop = 12
                        Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                        Call ImpreEtiq(I)
                        NumEti = NumEti + 1
                    ElseIf I = 1 Then
                        Printer.ScaleMode = vbMillimeters
                        Printer.Font = "Tahoma" '"Courier New"
                        Printer.FontSize = 8
                        Printer.Print ""
                        Barcode1.PrinterScaleMode = Printer.ScaleMode
                        Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                        Call ImpreEtiq2(I)
                    Else
                        Printer.ScaleMode = vbMillimeters
                        Printer.Font = "Tahoma" '"Courier New"
                        Printer.FontSize = 8
                        Printer.Print ""
                        Barcode1.PrinterScaleMode = Printer.ScaleMode
                        Call Barra("Imprimiendo etiqueta " & Actual & " de " & Cuantos)
                        Call ImpreEtiq3(I)
                    End If
                    
                    Pag = Pag + 1
                    If I = 0 Then
                         I = 1
                    ElseIf I = 1 Then
                        I = 2
                    Else
                         I = 0
                         Printer.NewPage
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
     s = "select * from parametros where cia='" + CiA + "'"
     Set SysParam = DatOS.OpenRecordset(s, 4)
     SimBOlo = Nulo(SysParam!SimBOlo)
     Show
     Refresh
End Sub

Private Sub Text_Change(Index As Integer)
     Select Case Index
     Case 0, 1
          NomArt(Index).Caption = ""
          s = "select c_articulo,d_articulo from articulos "
          s = s + "where cia='" + CiA + "' and c_articulo='" + Text(Index) + "'"
          Set Articulos = DatOS.OpenRecordset(s)
          If Not Articulos.EOF Then
               NomArt(Index).Caption = Articulos!d_articulo
               Barcode1.Caption = Text(Index).Text
          End If
     Case 4
          Text(3) = ((Text(4) - 1) * 3) + Text(5)
     Case 5
          Text(3) = (Text(4) - 1) * 3 + Text(5)
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
               s = "select c_articulo as Código,d_articulo as Descripción "
               s = s + "from articulos where cia='" + CiA + "' order by d_Articulo"
          End Select
          Call Lista.Carga(Text(Index), s, Titulo)
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
Private Function ImpreEtiq(Lado%) As Boolean
     
On Error GoTo CMDErr
Dim Start%, Xx$

Start = 0
LabelTop = 0

With Articulos
     'La descripcion
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 1
     Xx = Mid(Articulos!Alterno, 1, 48)
     Printer.Print Xx
     
     'Ubicacion = fabricante
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 4
     Xx = Articulos!ubicacion
     Printer.Print Xx
     
     'El precio con iv
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 18
     Printer.CurrentY = LabelTop + 1
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC * IIf(!sino_impu = 1, (IV + 1), 1), "###,###")
     Printer.Print Xx
     Printer.FontBold = False
     
    'Etiqueta de la Izquierda
     Barcode1.PrinterLeft = Start + 1
     Barcode1.PrinterTop = 8
     Barcode1.PrinterWidth = 30
     Barcode1.PrinterHeight = 6
     Barcode1.Caption = "*"
     Barcode1.Caption = Articulos!c_articulo
     Barcode1.PrinterHDC = Printer.hDC
     ' La Barra
     Printer.FontSize = 8
     Printer.CurrentX = Barcode1.PrinterLeft
     Printer.CurrentY = LabelTop + 14 ' Barcode1.PrinterTop + Barcode1.PrinterHeight + 2
     Printer.Print Barcode1.Displayed
     Printer.FontBold = False
     Printer.FontSize = 10
     'Printer.NewPage
     'Printer.EndDoc
     ImpreEtiq = True
     End With
CMDErr:
     If err.Number = 482 Then
          s = "La impresora no está lista!, verifique que esté encendida y el papel colocado."
          If MsgBox(s, 21, Printer.DeviceName) = 4 Then
               Resume
          Else
               Exit Function
          End If
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
     End If
     On Error GoTo 0
End Function
Private Function ImpreEtiq2(Lado%) As Boolean
     On Error GoTo CMDErr
     Dim Start%, Xx$
     Start = 35
     LabelTop = 0
     
     
With Articulos
     'La descripcion
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 1
     Xx = Mid(Articulos!Alterno, 1, 48)
     Printer.Print Xx
     
     'Ubicacion = fabricante
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 4
     Xx = Articulos!ubicacion
     Printer.Print Xx
     
     'El precio con iv
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 18
     Printer.CurrentY = LabelTop + 1
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC * IIf(!sino_impu = 1, (IV + 1), 1), "###,###")
     Printer.Print Xx
     Printer.FontBold = False
     
    'Etiqueta de la Izquierda
     Barcode1.PrinterLeft = Start + 1
     Barcode1.PrinterTop = 8
     Barcode1.PrinterWidth = 30
     Barcode1.PrinterHeight = 6
     Barcode1.Caption = "*"
     Barcode1.Caption = Articulos!c_articulo
     Barcode1.PrinterHDC = Printer.hDC
     ' La Barra
     Printer.FontSize = 8
     Printer.CurrentX = Barcode1.PrinterLeft
     Printer.CurrentY = LabelTop + 14 ' Barcode1.PrinterTop + Barcode1.PrinterHeight + 2
     Printer.Print Barcode1.Displayed

     Printer.FontBold = False
     Printer.FontSize = 10
     'Printer.NewPage
     ImpreEtiq2 = True
     End With
CMDErr:
     If err.Number = 482 Then
          s = "La impresora no está lista!, verifique que esté encendida y el papel colocado."
          If MsgBox(s, 21, Printer.DeviceName) = 4 Then
               Resume
          Else
               Exit Function
          End If
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
     End If
     On Error GoTo 0
End Function
Private Function ImpreEtiq3(Lado%) As Boolean
     On Error GoTo CMDErr
     Dim Start%, Xx$
     Start = 70
     LabelTop = 0
     
     
With Articulos
     'La descripcion
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 1
     Xx = Mid(Articulos!Alterno, 1, 48)
     Printer.Print Xx
     
     'Ubicacion = fabricante
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 1
     Printer.CurrentY = LabelTop + 4
     Xx = Articulos!ubicacion
     Printer.Print Xx
     
     'El precio con iv
     Printer.FontBold = True
     Printer.FontSize = 8
     Printer.CurrentX = Start + 18
     Printer.CurrentY = LabelTop + 1
     Xx = SimBOlo + Format(Nulo(!monto) * TipoC * IIf(!sino_impu = 1, (IV + 1), 1), "###,###")
     Printer.Print Xx
     Printer.FontBold = False
     
    'Etiqueta de la Izquierda
     Barcode1.PrinterLeft = Start + 1
     Barcode1.PrinterTop = 8
     Barcode1.PrinterWidth = 30
     Barcode1.PrinterHeight = 6
     Barcode1.Caption = "*"
     Barcode1.Caption = Articulos!c_articulo
     Barcode1.PrinterHDC = Printer.hDC
     ' La Barra
     Printer.FontSize = 8
     Printer.CurrentX = Barcode1.PrinterLeft
     Printer.CurrentY = LabelTop + 14 ' Barcode1.PrinterTop + Barcode1.PrinterHeight + 2
     Printer.Print Barcode1.Displayed
     Printer.FontBold = False
     Printer.FontSize = 10
     'Printer.NewPage
     'Printer.EndDoc
     ImpreEtiq3 = True
     End With
CMDErr:
     If err.Number = 482 Then
          s = "La impresora no está lista!, verifique que esté encendida y el papel colocado."
          If MsgBox(s, 21, Printer.DeviceName) = 4 Then
               Resume
          Else
               Exit Function
          End If
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Imprimir códigos de Barras"
     End If
     On Error GoTo 0
End Function

