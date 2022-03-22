VERSION 5.00
Begin VB.Form CantHoras 
   BackColor       =   &H80000003&
   Caption         =   "Cantidad de Horas"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   645
      Left            =   2160
      Picture         =   "CantHoras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1890
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   645
      Left            =   825
      Picture         =   "CantHoras.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1890
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1020
      Left            =   1440
      TabIndex        =   0
      Text            =   " "
      Top             =   660
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Digite la Cantidad de Horas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   345
      TabIndex        =   3
      Top             =   195
      Width           =   4050
   End
End
Attribute VB_Name = "CantHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Monitoreo.HAlq.Text = Doble(Text1)
    Unload Me
End Sub

Private Sub Command2_Click()
    Monitoreo.HAlq.Text = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Centra Me
End Sub
