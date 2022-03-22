VERSION 5.00
Begin VB.Form DescAdi 
   Caption         =   "Descuento"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2280
      Picture         =   "DesAdi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   4095
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Costo"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Precio Venta"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "P Venta c/ I.V.I"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Imprimir al :"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Descuento Adicional:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "DescAdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Porcentaje del Descuento Adicional
    If Text1.Text = "" Then
        DesAdi = 0
    Else
        DesAdi = Text1.Text
    End If
    'Tipo de Precio a imprimir la Entrada
    If Option1.Value Then
        monImp = "1"
    ElseIf Option2.Value Then
        monImp = "2"
    ElseIf Option3.Value Then
        monImp = "3"
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    Centra Me
End Sub
