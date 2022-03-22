VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FactProd 
   Caption         =   "Facturacion de Productos"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   13
      Left            =   960
      MaxLength       =   15
      TabIndex        =   35
      ToolTipText     =   "Digite con cuanto paga el cliente !"
      Top             =   6345
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   690
      TabIndex        =   31
      Top             =   1530
      Width           =   8100
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   165
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2505
      TabIndex        =   29
      Top             =   1035
      Width           =   6285
   End
   Begin VB.TextBox Text 
      Height          =   330
      Index           =   2
      Left            =   855
      MaxLength       =   2
      TabIndex        =   26
      ToolTipText     =   "La bodega de la cual se procede a vender"
      Top             =   510
      Width           =   435
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      FillColor       =   &H00C0FFC0&
      ForeColor       =   &H0080FFFF&
      Height          =   885
      Left            =   45
      ScaleHeight     =   825
      ScaleWidth      =   8775
      TabIndex        =   11
      Top             =   1965
      Width           =   8835
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         Caption         =   "IV"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4980
         TabIndex        =   33
         Top             =   495
         Width           =   525
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   7365
         MaxLength       =   5
         TabIndex        =   32
         Top             =   435
         Width           =   615
      End
      Begin VB.TextBox Text 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "La cantidad a facturar"
         Top             =   60
         Width           =   5325
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "La cantidad a facturar"
         Top             =   420
         Width           =   1725
      End
      Begin VB.TextBox Text 
         Height          =   330
         Index           =   4
         Left            =   1095
         TabIndex        =   13
         ToolTipText     =   "El código del articulo"
         Top             =   60
         Width           =   1560
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "La cantidad a facturar"
         Top             =   420
         Width           =   645
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1741
         TabIndex        =   14
         Top             =   420
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         BuddyControl    =   "Text(5)"
         BuddyDispid     =   196609
         BuddyIndex      =   5
         OrigLeft        =   2010
         OrigTop         =   450
         OrigRight       =   2250
         OrigBottom      =   795
         Max             =   999999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC(%) :"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6435
         TabIndex        =   34
         Top             =   495
         Width           =   870
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   270
         TabIndex        =   17
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2355
         TabIndex        =   16
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   345
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nueva"
      Height          =   585
      Left            =   1980
      Picture         =   "FactProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Presione para generar una factura nueva"
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Agregar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   585
      Left            =   3120
      Picture         =   "FactProd.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Agregue el artículo seleccionado a la lista"
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Actualizar"
      Enabled         =   0   'False
      Height          =   585
      Left            =   4290
      Picture         =   "FactProd.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Actualize la modificaciones realizadas"
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Borrar"
      Enabled         =   0   'False
      Height          =   585
      Left            =   5430
      Picture         =   "FactProd.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Elimina el artículo seleccionado de la lista"
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Pagar"
      Enabled         =   0   'False
      Height          =   585
      Left            =   6540
      Picture         =   "FactProd.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Graba la factura y la imprime"
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Salir"
      Height          =   585
      Left            =   7695
      Picture         =   "FactProd.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cierra esta ventana"
      Top             =   6765
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1035
      Width           =   600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registro del Pago"
      Height          =   975
      Left            =   5025
      TabIndex        =   0
      Top             =   30
      Width           =   3810
      Begin VB.OptionButton Option1 
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   210
         TabIndex        =   2
         Top             =   450
         Width           =   1725
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Contado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2085
         TabIndex        =   1
         Top             =   465
         Value           =   -1  'True
         Width           =   1560
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   45
      TabIndex        =   22
      Top             =   2925
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Codigo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descripcion"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Unid"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Precio"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SubTotal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Desc(%)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descuento"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "IV(%)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Impuesto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Estacion"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Impuesto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   960
      TabIndex        =   42
      ToolTipText     =   "El impuesto de la factura"
      Top             =   5985
      Width           =   2175
   End
   Begin VB.Label Subtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   960
      TabIndex        =   41
      ToolTipText     =   "El subtotal de la factura"
      Top             =   5265
      Width           =   2175
   End
   Begin VB.Label Descuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   960
      TabIndex        =   40
      ToolTipText     =   "El descuento de la factura"
      Top             =   5625
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descuento :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   -30
      TabIndex        =   39
      Top             =   5685
      Width           =   945
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Subtotal :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   180
      TabIndex        =   38
      Top             =   5325
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Impuesto :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   37
      Top             =   6045
      Width           =   855
   End
   Begin VB.Label PagaCon 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Paga con :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   105
      TabIndex        =   36
      Top             =   6375
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "Nota: "
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   165
      TabIndex        =   30
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bodega :"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   675
   End
   Begin VB.Label NomBod 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1305
      TabIndex        =   27
      Top             =   510
      Width           =   3465
   End
   Begin VB.Label Label1 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5220
      TabIndex        =   25
      Top             =   5340
      Width           =   3045
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número :"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   105
      TabIndex        =   24
      Top             =   135
      Width           =   750
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   915
      TabIndex        =   23
      Top             =   45
      Width           =   1515
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   765
      Left            =   4560
      TabIndex        =   20
      ToolTipText     =   "El total de la factura"
      Top             =   5805
      Width           =   4230
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   3270
      TabIndex        =   19
      Top             =   5940
      Width           =   1230
   End
   Begin VB.Label Label9 
      Caption         =   "ESTACION:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   1035
      Width           =   1650
   End
End
Attribute VB_Name = "FactProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Articulos As Recordset
Dim Bodegas As Recordset
Dim Precios As Recordset
Dim Item As ListItem
Private Sub Command1_Click()
    Call Limpia
End Sub

Private Sub Command2_Click()
    Call Llena(1)
End Sub

Private Sub Command3_Click()
    Call Llena(2)
End Sub

Private Sub Command5_Click()
    If Text2.Text <> "99" And Label1 = "" Then
        MsgBox "La estación no esta definida, Selecione una estacion valida!", vbInformation
        Exit Sub
    End If
    If NomBod.Caption = "" Then
        MsgBox "La Bodega no esta definida, Selecione una Bodega valida!", vbInformation
        Exit Sub
    End If
    Call GuardaFactu
End Sub
Public Function GuardaFactu(Optional ElEfectivo As Double) As Boolean   ' guardar la factura
     EspaCio.BeginTrans
     If Inserta Then
        EspaCio.CommitTrans
        'IMPRIME LA FACTURA
        Call Imprimefactura(Label7.Caption, Text(0), Report1)
        
        'ImpFact ElEfectivo
        S = "update consecutivos set consecutivo='" + Label7 + "' where cia='" + CiA + "'"
        S = S + " and tipoconse='F01'"
        DatOS.Execute S
        MsgBox "Factura Nº:" + Label7.Caption, vbInformation
        Command4.Enabled = False
        ListView1.ListItems.Clear
        ListView1.Enabled = False
        Command1_Click
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text2.Visible = False
        GuardaFactu = True
     Else
        EspaCio.Rollback
        GuardaFactu = False
     End If
End Function
Public Sub Imprimefactura(FactExp$, Bodega$, Reporte As CrystalReport)
     On Error GoTo Errores
     Dim Si%
     S = ReadKey("printerFact")
     Dim Pr As Printer
     For Each Pr In Printers
          If LCase(Pr.DeviceName) = LCase(S) Then
               Set Printer = Pr
               Si = 1
               Exit For
          End If
     Next
     If Si = 0 Then
          MsgBox "No se encontró la impresora de Facturas !", 16, Caption
     End If
     
     FactExp = Format(FactExp, "00000000")
     S = "{facturas.cia}='" + CiA + "' and {facturas.c_bodega}='" + Text2
     S = S + "' and {facturas.n_factura}='" + FactExp + "'"
     Reporte.ReportFileName = DirTrA + "reportes\Factvideo.rpt"
     Reporte.DataFiles(0) = DatOS.Name
     'Reporte.Formulas(5) = "nomclie='" + Text1 + "'"
     'Reporte.Formulas(6) = "comodin='" + Inicio.StatusBar1.Panels(2) + "'"
     Reporte.WindowTitle = "Factura"
     Reporte.WindowShowPrintSetupBtn = True
     Reporte.Destination = 0
     Reporte.PrinterName = Printer.DeviceName
     Reporte.PrinterDriver = Printer.DriverName
     Reporte.PrinterPort = Printer.Port
     Reporte.SelectionFormula = S
     Reporte.Action = 1
Errores:
     If err.Number = 20526 Then
          MsgBox "Debe existir al menos una impresora en el sistema !", 48, "Entradas"
     ElseIf err.Number <> 0 Then
          MsgBox err.Description + Str(err.Number), 16, "ImprimeFactura"
     End If
     On Error GoTo 0
End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Llena(Tipo As Integer)
          Dim Cant#
          Cant = Doble(Text(5))     ' * Doble(Label16)
          If Not AgregaItem(Text(4).Text, Text(1), Cant, _
                 1, Doble(Text(6)), Check1, _
                 Doble(Text(0)), Tipo, Articulos!Minimo, _
                 Articulos!p_compra, 0) Then
                 Exit Sub
          End If
          Text(0) = ""
          Text(1) = ""
          Text(4) = ""
          Text(5) = "1"
          Text(4).SetFocus
End Sub
Private Function AgregaItem(CodArt$, Articulo$, Cant#, Unidades#, _
     PorcDesc#, PorcImp%, PreciO#, Novo%, Minimo&, Costo#, _
      Calculado%) As Boolean
     On Error GoTo Errores
     'If ListView1.ListItems.Count > 11 Then
     '   MsgBox "Ya alcanzó las 12 lineas permitidas", vbCritical, "Agregar Artículo"
     '   AgregaItem = False
     '   Exit Function
     'End If
     Dim IvI#
     Dim Descuento#
     Dim SubTotal#
     Dim Total#
     Dim Impuesto#
     Dim Avail#
     Dim ExisTot#
     Dim Llave$
     Dim Item2 As ListItem
     impvta = IV
     If Check1.Value = 1 Then IvI = IIf(PorcImp = 1, IV, 0)
     If Cant = 0 Then err.Raise 9996
     'If Costo >= Precio Then MsgBox "Costo mayor que el precio", vbExclamation, "Incluir Artículo"      'Err.Raise 9994
     SubTotal = (Cant * PreciO)
     SubTotal = SubTotal - (SubTotal * (PorcDesc / 100))
     SubTotal = SubTotal + (SubTotal * (IvI / 100))
     'SubTotal = 0
     'Cant = Cant * Unidades
     'Existencias
     'If Calculado = 0 Then
          'Avail = ValidaExis(CodArt, Text(0))
          'If (Avail - Cant) < 0 Then err.Raise 9999
          'ExisTot = ValidaMinimo
          'If (ExisTot - Cant) < Minimo Then err.Raise 9998
     'End If
     'Minimo del articulo
     Llave = "*" + CodArt
     If Novo = 1 Then
          Set Item = ListView1.ListItems.Add(, Llave)
          'Set Item2 = ListView2.ListItems.Add
     ElseIf Novo = 2 Then
          Set Item = ListView1.SelectedItem
          'Set Item2 = ListView2.ListItems(item.Index)
     End If
     'La Lista Visible
     Item.Tag = Costo
     Item.Text = CodArt 'La descripcion
     Item.SubItems(1) = Articulo 'La descripcion
     Item.SubItems(2) = Unidades 'Las unidades
     Item.SubItems(3) = Format(PreciO, "standard") 'El costo
     Item.SubItems(4) = Cant 'La cantidad
     SubTotal = Cant * PreciO 'Calcula el subtotal
     Item.SubItems(5) = Format(SubTotal, "standard") 'El subtotal
     Descuento = SubTotal * (PorcDesc / 100) 'Calcula el descuento
     Item.SubItems(6) = PorcDesc 'El % de descuento
     Item.SubItems(7) = Format(Descuento, "standard") 'El descuento
     Impuesto = (SubTotal - Descuento) * IvI 'Calcula el impuesto
     Item.SubItems(8) = Format(IvI * 100, "standard") 'El % de impuesto
     Item.SubItems(9) = Format(Impuesto, "standard") 'El impuesto
     Total = (SubTotal - Descuento) + Impuesto 'Calcula el total
     Item.SubItems(10) = Format(Total, "standard") 'El total
     Item.SubItems(11) = Text2.Text  'La estacion
     'La Lista Invisible
      
     'Item2.Text = ECombo(5).ListIndex 'El numero de precio
     Set Item = Nothing
     AgregaItem = True
     Call CalculaSaldo
Errores:
     If err.Number = 35602 Then
          S = "El artículo ya existe Desea Agregarlo ?"
          'If MsgBox(S, 36, Caption) = 6 Then
               Set Item = ListView1.ListItems("*" + CodArt)
               Item.SubItems(4) = Cant + Item.SubItems(4)
               Item.SubItems(5) = FormatNumber(Item.SubItems(3) * Item.SubItems(4))
               Item.SubItems(7) = FormatNumber(Item.SubItems(6) / 100 * Item.SubItems(5))
               Item.SubItems(9) = FormatNumber(Item.SubItems(8) / 100 * (Item.SubItems(5) - Item.SubItems(7)))
               Item.SubItems(10) = FormatNumber((Item.SubItems(5) - Item.SubItems(7) + Item.SubItems(9)))
               Item.EnsureVisible
               Item.Selected = True
               AgregaItem = True
               Call CalculaSaldo
          'End If
     ElseIf err.Number = 9999 Then
          If Seguridad("factura1") Then
               S = "La cantidad a facturar  es mayor a la existencia"
               S = S + Chr(13) + "   en la bodega seleccionada, desea continuar ?"
               If MsgBox(S, 36, "Existencia: " + FormatNumber(Avail)) = 6 Then
                    Resume Next
               Else
                    Selecciona Text(5)
               End If
          Else
               S = "La cantidad a facturar es mayor a la existencia"
               S = S + Chr(13) + "          en la bodega seleccionada!"
               MsgBox S, 64, "Existencia: " + FormatNumber(Avail)
               Selecciona Text(5)
          End If
     ElseIf err.Number = 9996 Then
          MsgBox "La cantidad debe ser mayor a cero !", 16, "Error"
          Text(5).SetFocus
     ElseIf err.Number = 9998 Then
          If Seguridad("factura2") Then
               S = "La cantidad solicitada menos la existencia es menor"
               S = S + Chr(13) + "     al mínimo del articulo, desea continuar ?"
               If MsgBox(S, 36, "Existencia: " + FormatNumber(Avail) + " Mínimo: " + FormatNumber(Articulos!Minimo)) = 6 Then
                    Resume Next
               Else
                    Exit Function
               End If
          Else
               S = "La cantidad solicitada menos la existencia es menor al mínimo del articulo !"
               MsgBox S, 64, "Existencia: " + FormatNumber(Avail) + " Mínimo: " + FormatNumber(Articulos!Minimo)
               Selecciona Text(5)
          End If
     ElseIf err.Number = 9995 Then
          If Seguridad("factura3") Then
               S = "El monto a facturar sobrepasa el disponible de crédito del cliente,"
               S = S + Chr(13) + "                         Desea continuar ?"
               If MsgBox(S, 36, "Disponible: " + FormatNumber(Avail)) = 6 Then
                    Resume Next
               Else
                    Exit Function
               End If
          Else
               S = "El monto a facturar sobrepasa el límite de crédito del cliente !"
               MsgBox S, 64, "Disponible: " + FormatNumber(Avail)
               Selecciona Text(5)
          End If
     ElseIf err.Number = 9994 Then
          S = "El costo del artículo seleccionado es mayor o igual "
          S = S + Chr(13) + "   al precio de venta, no es posible facturarlo."
          MsgBox S, 16, "Costo: " + FormatNumber(Costo)
          Text(4).SetFocus
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Agrega Articulo"
     End If
End Function
Private Sub CalculaSaldo()
     's = "select * from parametros where cia='" + CiA + "'"
     'Set param = DatOS.OpenRecordset(s)
     If ListView1.ListItems.Count > 0 Then
          Dim IV#
          Dim SubTot#
          Dim Costo#
          Dim Tot#
          Dim DesC#
          Dim Item As ListItem
          For I = 1 To ListView1.ListItems.Count
               Set Item = ListView1.ListItems(I)
               SubTot = SubTot + Doble(Item.SubItems(5))
               DesC = DesC + Doble(Item.SubItems(7))
               IV = IV + Doble(Item.SubItems(9))
               Tot = Tot + Doble(Item.SubItems(10))
               Costo = Costo + (Doble(Item.Tag) * Doble(Item.SubItems(4)))
          Next
          'SubTotal.Tag = Costo
          'If Check4.Value = 1 Then
          '     If Text2.Text = "" Then Text2.Text = 0
          '     desc = Format(desc + ((SubTot - desc) * (Text2.Text / 100)), "standard")
          '     If IV > 0 Then IV = (SubTot - desc) * impvta
          'End If
          SubTotal.Caption = Format(SubTot, "standard")
          SubTot = SubTot - DesC
          'Descuento global
          'Desc = Desc + (SubTot * (Doble(Text(7).Text) / 100))
          'Le resta el descuento
          'Tot = SubTot - DesC
          'Le suma el impuesto
          'Tot = Tot + IV
          'Le suma el flete
          'Tot = Tot + Doble(Text(10))
          Descuento.Caption = Format(DesC, "standard")
          Impuesto.Caption = Format(IV, "standard")
          'If Option2.Value = True Then
          '  Label26 = Format((tot * ((param!impcons / 100))), "standard")
          '  tot = (tot * (1 + (param!impcons / 100)))
          'Else
          '  Label26 = "0.00"
          'End If
          Total.Caption = Format(Tot, "standard") 'Redondea a 5 Colones Siguientes
          Command5.Enabled = True
     Else
          Total.Caption = ""
          Command5.Enabled = False
     End If
End Sub
Private Sub ActualizaConsec()         ' 01........
      EsTaC = "01"
      Label7.Caption = Format(Consecutivo("F01"), "00000000")
End Sub
Private Function Inserta() As Boolean
     On Error GoTo INSErr
     Call Barra("Generando Factura ...")
     Call ActualizaConsec
     'Inserta la factura
     DatOS.Execute Arma, 129
     'Inserta el desglose de la factura y actualiza existencias y estadisticas
     Call Barra("Actualizando desglose de la factura ...")
     For I = 1 To ListView1.ListItems.Count
          Set Item = ListView1.ListItems(I)
          'Set Item2 = ListView2.ListItems(item.Index)
          'Desglose
          S = "insert into desgfact(c_articulo,n_factura,cantidad,"
          S = S + "unidades,precio,total_brut,descuento,porc_imp,"
          S = S + "total_neto,c_bodega,costo,lote,cia,serie) values ('"
          S = S + Item.Text + "','" 'El codigo de articulo
          S = S + Label7.Caption + "'," 'El numero de factura
          S = S + Item.SubItems(4) + "," 'La cantidad vendida
          S = S + Item.SubItems(2) + "," 'El numero de unidades
          S = S + Format(Doble(Item.SubItems(3))) + "," 'El precio de cada unidad
          S = S + Format(Doble(Item.SubItems(5))) + "," 'El total sin IV y sin descuento
          S = S + Format(Doble(Item.SubItems(6))) + "," 'El % de descuento
          S = S + Format(Doble(Item.SubItems(8))) + "," 'El % de impuesto
          S = S + Format(Doble(Item.SubItems(10))) + ",'" 'El total con IV y con descuento
          S = S + Text2 + "'," 'LA BODEGA
          S = S + Trim(Item.Tag)  'El costo,
          S = S + ",'','" + CiA + "','" + Text2 + "')" 'El Lote cia serie
          DatOS.Execute S, 128
          'Actualiza existencias
          S = "update existencias set existencia=existencia-"
          S = S & Doble(Item.SubItems(4))
          S = S + ", f_ult_sal=#" + Format(Date, "m/d/yyyy")
          S = S + "#,doc_ult_sa='" + Label7.Caption
          S = S + "',tiposalida=0 " + " where cia='" + CiA + "' and c_bodega='" + Text(2)
          S = S + "' and c_articulo='" + Item.Text + "'"
          DatOS.Execute S, 128
     Next
   
     'Las facturas Pendiente
     If Option1.Value = True Then
          S = "insert into Apartados(cia,c_bodega,n_factura,estado,"
          S = S + "nombre,monto,saldo,ultpago,telefono,codclie) "
          S = S + "Values('" + CiA + "','" + Text2 + "','"
          S = S + Label7.Caption + "','" 'El numero de factura
          S = S + "01','"        ' estado
          S = S + "'," 'nombre
          S = S & Doble(Total) & "," 'El total de la factura
          S = S & Doble(Total) & ","      ' el saldo
          S = S + "#" + Format(Date, "m/d/yyyy") + "#,'" 'La fecha
          S = S + Text(2) + "','9999')"  ' telefono
          DatOS.Execute S, 128
     End If
     
     Inserta = True
     Call GBitacora(1, "Factura No. " + Label7.Caption)
     On Error GoTo 0
INSErr:
     If err.Number = 3022 Then
          Label7.Caption = Format(Val(Label7) + 1, "0000000000")
          Resume
     ElseIf err.Number = 3315 Then
          MsgBox "El campo:" + err.Description, 16, "Campos en Blanco"
     ElseIf err.Number = 3201 Then
          MsgBox "Al menos uno de los datos es incorrecto !", 16, "Error"
     ElseIf err.Number = 3003 Then
          S = "Ha ocurrido un error transaccional en la base de datos,"
          S = S + Chr(13)
          S = S + "       Cierre esta ventana y abrala de nuevo."
          MsgBox S, 16, "Error"
          Command6.SetFocus
     ElseIf err.Number > 0 Then
          MsgBox err.Description + Str(err.Number), 16, "Inserta"
     End If
     Call Barra("")
     On Error GoTo 0
End Function
Private Function Arma() As String
     S = "insert into facturas (n_factura,c_bodega,impuesto,"
     S = S + "monto,monto_real,f_factura,plazo,c_cliente,descuento,"
     S = S + "c_tipo_fac,vendedor,estado,flete,costo,nota,"
     S = S + "dolares,tipocambio,cia,autoriza,ctf_1,ctf_2,ctf_3,pago_1,pago_2"
     S = S + ",pago_3,telefono,Tipofact,financiamiento,Cliente) "
     S = S + "values('" + Label7.Caption + "','" 'El numero de factura
     S = S + Text2 + "'," 'El codigo de bodega
     S = S & Impuesto & "," 'El monto del impuesto
     S = S & Doble(SubTotal) & "," 'El subtotal
     S = S & Doble(Total) & ",#" 'El total de la factura
     S = S + Format(Date, "m/d/yyyy") + "#," 'La fecha
     S = S & "0,'" 'El plazo
     S = S + "9999'," 'El cliente
     S = S & Descuento & ",'" 'El descuento
     S = S & 0 & "','" 'El tipo de pago
     S = S & "'," 'El vendedor
     S = S & "0," 'El estado 0= No nula 1=Nula
     S = S & "0," 'El flete
     S = S & "0,'" 'El costo
     S = S & Text3 & "',0," 'La nota
     S = S & TCambio & ",'" + CiA + "','" + LoGiN + "','"
     S = S & 0 & "','" & 0 & "','" & 0 & "',"
     S = S & 0 & "," & 0 & "," & 0
     S = S & ",''," & Option1.Value & ",0,'" & Text1 & "')"
     Arma = S
     Debug.Print S
End Function

Private Sub Command6_KeyUp(KeyCode As Integer, Shift As Integer)
     Dim Titulo$
     Select Case KeyCode
     Case vbKeyF3
          Select Case Index
          Case 4
               S = "select c_articulo as Código,d_articulo as Descripción "
               S = S + "from articulos where cia='" + CiA + "' order by d_articulo"
               Titulo = "Artículos"
          End Select
          Call Lista.Carga(Text(Index), S, Titulo)
          Lista.Show 1
    End Select
End Sub

Private Sub Form_Load()
Centra Me
Text2 = "99"
Text(2) = "01"
End Sub

Private Sub ListView1_Click()
     Command3.Enabled = True
     Command4.Enabled = True
     Call CargaItem(Item)
End Sub
Private Sub CargaItem(Item As ListItem)
     Set Item = ListView1.SelectedItem
     Text(4) = Item.Text
     Text(5) = Item.SubItems(4)
     Text(0) = Item.SubItems(3)
     'Text(6) = item.SubItems(6)
     'Check1.Value = IIf(Val(item.SubItems(8)) = 0, 0, 1)
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        Text2 = ""
        Text2.Enabled = True
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Text2 = "99"
        Text2.Enabled = False
    End If
End Sub
Private Sub Limpia()
     ListView1.ListItems.Clear
     Option2 = True
     Text2 = "99"
     Text(4) = ""
     Text(5) = "1"
     Text(1) = ""
     Text(0) = ""
     Total.Caption = ""
     Label7 = ""
End Sub

Private Sub Text_Change(Index As Integer)
Select Case Index
     Case 2
            S = "select * from bodegas where cia='" + CiA + "'"
            S = S + " and c_bodega='" + Text(2) + "'"
            Set Bodegas = DatOS.OpenRecordset(S)
            If Not Bodegas.EOF Then
                NomBod = Bodegas!d_bodega
            Else
                NomBod = ""
            End If
     Case 4
            S = "select * from articulos where cia='" + CiA + "'"
            S = S + " and c_articulo='" + Text(4) + "'"
            Set Articulos = DatOS.OpenRecordset(S)
            If Not Articulos.EOF Then
                Text(1) = Articulos!d_articulo
                Check1.Value = Articulos!sino_impu 'El impuesto
                S = "select * from listaprecios where cia='" + CiA + "'"
                S = S + " and codart='" + Text(4) + "' and codigo=1"
                Set Precios = DatOS.OpenRecordset(S)
                If Not Precios.EOF Then
                    Text(0) = Precios!Monto
                End If
            End If
     End Select
     If Index = 4 Or Index = 5 Or Index = 0 Then Call Valida
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     Dim Titulo$
     Select Case KeyCode
     Case vbKeyF3
          Select Case Index
          Case 2
               S = "select c_Bodega as Codigo,d_bodega as Nombre from bodegas"
               S = S + " where cia='" + CiA + "'"
               Titulo = "Bodegas"
          Case 4
               S = "select c_articulo as Código,d_articulo as Descripción "
               S = S + "from articulos where cia='" + CiA + "' order by d_articulo"
               Titulo = "Artículos"
          End Select
          Call Lista.Carga(Text(Index), S, Titulo)
          Lista.Show 1
     End Select
End Sub
Private Function Valida() As Boolean
     Command2.Enabled = False
     Command3.Enabled = False
     If Text(4) = "" Then Exit Function
     If Text(0) = "" Then Exit Function
      
     If Text(5) = "" Then Exit Function
     Command2.Enabled = True
     Command3.Enabled = True
     Valida = True
End Function

Private Sub Text2_Change()
Dim esta As Recordset
    S = "select * from estaciones where cia='" + CiA + "'"
    S = S + " and codigo='" + Text2.Text + "'"
    Set esta = DatOS.OpenRecordset(S)
    If Not esta.EOF Then
        Label1 = esta!Descripcion
    Else
        Label1 = ""
    End If
End Sub
