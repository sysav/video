VERSION 5.00
Begin VB.Form Acerca 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   9135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Acerca.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Acerca.frx":0E42
   ScaleHeight     =   2619.375
   ScaleMode       =   0  'User
   ScaleWidth      =   8578.236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   180
      TabIndex        =   7
      Top             =   2460
      Width           =   8745
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   7500
      TabIndex        =   0
      Top             =   3300
      Width           =   1560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Web : www.hovisys.com   Soporte : soporte@hovisys.com"
      Height          =   225
      Left            =   480
      TabIndex        =   9
      Top             =   1620
      Width           =   5565
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   180
      Picture         =   "Acerca.frx":24764
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   5880
   End
   Begin VB.Label DEscripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   225
      Left            =   480
      TabIndex        =   8
      Top             =   900
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "San José, Costa Rica"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   6
      Top             =   1380
      Width           =   5640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   5
      Top             =   1140
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compania licencia"
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
      Height          =   300
      Left            =   630
      TabIndex        =   4
      Top             =   2070
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   630
      Width           =   615
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Licencia de uso para :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   270
      TabIndex        =   1
      Top             =   1860
      Width           =   1830
   End
End
Attribute VB_Name = "Acerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
     Unload Me
End Sub
Private Sub cmdOK_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Load()
     lblTitle.Caption = App.ProductName
     Descripcion.Caption = App.FileDescription
     S = "Version "
     If ReadKey("IYFDV") = "11" Then S = S + " de demostración "
     S = S & App.Major & "." & App.Minor & "." & App.Revision
     S = S + " (" + App.Comments + ")"
     lblVersion.Caption = S
     Label1.Caption = ReadKey("nomcia") 'App.LegalTrademarks ' ReadKey("nomcia")
     Label2.Caption = "Copyright © 1997 - " + Format(Date, "yyyy") + " " + App.CompanyName
     Refresh
End Sub

