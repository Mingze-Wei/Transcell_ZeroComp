VERSION 5.00
Begin VB.Form frmTempPass 
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox tbPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "请输入密码："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmTempPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public closed As Boolean

Private Sub Command1_Click()
    If tbPass.Text = "123456" Then
        fMainForm.hasTempAuth = True
        closed = True
        Unload Me
    Else
        MsgBox ("密码不正确！")
        closed = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    closed = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closed = True
End Sub
