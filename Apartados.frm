VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Apartados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apartados y Ventas de Crédito"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Apartados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   4905
   ScaleWidth      =   10635
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Apartados.frx":030A
      Height          =   3135
      Left            =   60
      OleObjectBlob   =   "Apartados.frx":031E
      TabIndex        =   7
      Top             =   120
      Width           =   10425
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pagar"
      Height          =   315
      Left            =   6570
      TabIndex        =   16
      Top             =   3480
      Width           =   1050
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   4680
      TabIndex        =   15
      Top             =   3465
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busquedas  "
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   4095
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   2760
         TabIndex        =   1
         Top             =   315
         Width           =   285
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1035
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Código :"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Teléfono :"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre :"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
   End
   Begin ComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   7470
      Visible         =   0   'False
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   5370
      TabIndex        =   8
      Top             =   4140
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label DTienda 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3990
      TabIndex        =   10
      Top             =   5820
      Width           =   4185
   End
   Begin VB.Label Label1 
      Caption         =   "Tienda:"
      Height          =   225
      Left            =   4650
      TabIndex        =   9
      Top             =   4200
      Width           =   615
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":0D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":0E2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":0F41
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":1053
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":136D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":1687
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":1799
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":18AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":1BC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":1EDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":21F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Apartados.frx":2513
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Apartados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Boton%
Dim Item As ListItem
Dim S$
Dim Temp As Recordset

Private Sub Command1_Click()
    Call Carga
End Sub

Private Sub Command2_Click()
    Call Carga
End Sub
Private Sub Command3_Click()
    Call Carga
End Sub

Private Sub Command4_Click()
    PagaApa.Text1(9) = Mid(Text5.Text, 3)
    PagaApa.Show
End Sub
Private Sub DBGrid1_DblClick()
     DBGrid1.Bookmark = Data1.Recordset.Bookmark
     Text5.Text = DBGrid1.Columns(0)
End Sub
Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
     S = "select n_factura as Apartado,Nombre,Telefono,Monto,Saldo,ultpago AS Ultimo_Pago "
     S = S + "from apartados where cia='" + CiA + "' "
     S = S + "and c_bodega='" + Text1 + "'"
     Hay = False
     If ColIndex = 0 Then
          S = S + " order by apartado"
          Hay = True
     ElseIf ColIndex = 1 Then
          S = S + " order by nombre"
          Hay = True
     ElseIf ColIndex = 5 Then
          S = S + " order by telefono"
          Hay = True
     End If
     If Hay Then
        Data1.RecordSource = S

        Call Columnas(False)
     End If
End Sub

Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     'If Button = 2 Then PopupMenu Inicio.Popup, , , , Inicio.SubPop(0)
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Errores
     DBGrid1.Bookmark = Data1.Recordset.Bookmark
     Text5.Text = DBGrid1.Columns(0)
Errores:
    If err.Number <> 0 Then
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
     'Call Menus(True, Me)
End Sub
Private Sub Form_Deactivate()
     'Call Menus(False, Me)
End Sub
Private Sub Form_Load()
     Centra Me
     Set Actual = Me
     Data1.DatabaseName = DatOS.Name
     Data1.Refresh
     MousePointer = 11
     Text1 = LaBodega
     Call Barra("Un momento ... Cargando artículos !")
     Show
     Refresh
     Call Barra(Carga & " apartados encontrados", 1)
     MousePointer = 0
End Sub

Private Function Carga() As Long
     S = "select apartados.n_factura as Apartado,Nombre,apartados.Telefono,"
     S = S + "apartados.Monto,Saldo,ultpago AS Ultimo_Pago,IIf(tipofact = 1, 'CR', 'AP') as Tipo"
     S = S + " from apartados,facturas where apartados.cia='" + CiA + "' "
     S = S + " and apartados.c_bodega='" + Text1 + "'"
     S = S + " and apartados.cia=facturas.cia and apartados.n_factura=facturas.n_factura"
     If Text2 <> "" Then
        S = S + " and nombre like '*" + Text2 + "*'"
     End If
     If Text3 <> "" Then
        S = S + " and apartados.telefono like '*" + Text3 + "*'"
     End If
     If Text4 <> "" Then
        S = S + " and codclie like '*" + Text4 + "*'"
     End If
     Data1.RecordSource = S
     Data1.Refresh
     Call Columnas
End Function
Public Sub Columnas(Optional Refresca As Boolean)
     If Refresca Then Data1.Recordset.Requery
     DBGrid1.Columns(0).Width = 1300
     DBGrid1.Columns(1).Width = 3500
     DBGrid1.Columns(2).Width = 900
     DBGrid1.Columns(3).Width = 800       ' existen
     DBGrid1.Columns(4).Width = 700       ' porc descuento
     DBGrid1.Columns(5).Width = 2000      ' el grupo
     DBGrid1.Columns(2).NumberFormat = "####,##0"
     DBGrid1.Columns(2).Alignment = 1
     DBGrid1.Columns(3).NumberFormat = "###,##0"
     DBGrid1.Columns(4).NumberFormat = "###0"
End Sub
 
Private Sub Text1_LostFocus()
S = "Select d_bodega from bodegas where cia='" + CiA + "' and c_bodega='" + Text1.Text + "'"
Set Temp = DatOS.OpenRecordset(S)
If Temp.EOF Then
     DTienda.Caption = "**"
     MsgBox "No existe la Tienda", vbCritical, "APARTADOS"
     Text1.SetFocus
     Exit Sub
Else
     DTienda.Caption = Temp(0)
End If
Carga
End Sub
