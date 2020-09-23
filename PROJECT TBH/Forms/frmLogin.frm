VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   "
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   735
      Left            =   2880
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   735
      Left            =   1560
      Picture         =   "frmLogin.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   270
      TabIndex        =   5
      Top             =   4815
      Width           =   3840
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLogin.frx":2104
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   270
         TabIndex        =   6
         Top             =   270
         Width           =   3390
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   -225
      Picture         =   "frmLogin.frx":2192
      ScaleHeight     =   2355
      ScaleWidth      =   6000
      TabIndex        =   4
      Top             =   0
      Width           =   6000
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3285
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2700
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      Height          =   330
      Left            =   225
      TabIndex        =   3
      Top             =   3375
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME:"
      Height          =   375
      Left            =   225
      TabIndex        =   0
      Top             =   2745
      Width           =   1050
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdClose_Click()

    Unload Me
    Unload frmMain

End Sub

Private Sub cmdOk_Click()

    Static try As Integer
    
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    
    Call DBConnection
    
    
    rs.Open "select * from tbl_user where username = '" & Replace(Trim(txtUser.Text), "'", "") & "' and password = '" & Replace(Trim(txtPass.Text), "'", "") & "'", conn, adOpenStatic, adLockReadOnly
        
        If rs.EOF Then
        
             MsgBox "Invalid username or password! " & (2 - try) & " Try remaining...", vbExclamation, "FAILED"
             txtPass.Text = ""
             txtUser.SetFocus
             try = try + 1
             
                If try = 3 Then
                
                    MsgBox "You have reached the maximum trial!", vbInformation, "PROGRAM TERMINATED"
                    End
                
                End If
             
        Else
        
            MsgBox "Welcome " & rs.Fields("username") & " to THE BLAZING HERALD", vbInformation, "SUCCESS"
            frmMain.StatusBar1.Panels(2).Text = rs.Fields("username")
            try = 0
            isLoginOk = True
            Unload Me
            
        End If

End Sub

Private Sub Form_Load()

    isLoginOk = False

End Sub

