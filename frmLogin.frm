VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   6255
      Begin VB.CommandButton btnLogar 
         Caption         =   "Login"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton btnSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtsenha 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox txtlogin 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "Senha"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "PADARIA 21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogar_Click()

' Valida usuario do administrador

If txtlogin.Text = "Adm" And txtsenha.Text = "adm" Then
frmAdministrador.Show

Unload Me
ElseIf txtlogin.Text = "Caixa" And txtsenha.Text = "caixa" Then
frmCaixa.Show

Unload Me
Else
    MsgBox "Por favor, verifique os dados informados.", vbInformation, "Login - Padaria21"

End If

Unload Me

End Sub

Private Sub btnSair_Click()

If btnSair = True Then
MsgBox "Deseja realmente sair?", vbOKOnly

frmLogin.Hide
Unload Me


' OLHAR QUAL COMANDO PARA FECHAR O FORMULARIO

End If


End Sub
