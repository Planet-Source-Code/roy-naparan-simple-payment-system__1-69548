VERSION 5.00
Begin VB.Form frmEditUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT USER"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4575
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   615
      Left            =   3120
      Picture         =   "frmEditUser.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   615
      Left            =   1920
      Picture         =   "frmEditUser.frx":3E84
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1980
      TabIndex        =   5
      Top             =   2835
      Width           =   2265
   End
   Begin VB.TextBox txtPass2 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2205
      Width           =   2265
   End
   Begin VB.TextBox txtNewPass 
      BackColor       =   &H00F0F0F0&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1575
      Width           =   2265
   End
   Begin VB.TextBox txtOldPass 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   990
      Width           =   2265
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   360
      Width           =   2265
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      Height          =   510
      Left            =   225
      TabIndex        =   9
      Top             =   2925
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type Password:"
      Height          =   375
      Left            =   225
      TabIndex        =   8
      Top             =   2295
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      Height          =   330
      Left            =   225
      TabIndex        =   7
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   1080
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   420
      Left            =   270
      TabIndex        =   0
      Top             =   450
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()

        If txtOldPass.Text = frmUser.AdoUser.Recordset.Fields("password") Then
    
            If txtPass2.Text = txtNewPass.Text Then
        
                With frmUser.AdoUser.Recordset
                    
                    .Fields("username") = (txtUsername.Text)
                    .Fields("password") = (txtPass2.Text)
                    .Fields("type") = (txtType.Text)
                    .Update
                    
                    
                End With
        
            MsgBox "Record was updated successfully!", vbInformation, "EDITED"
            frmUser.DataGrid1.ReBind
            frmUser.DataGrid1.Refresh
            frmUser.AdoUser.Refresh
            Unload Me
       
        
            Else
            
                    MsgBox "Re-type new password!", vbExclamation, "INVALID"
            
            End If
        
    Else
        
        MsgBox "Invalid old password!", vbExclamation, "INVALID"
        
    End If

End Sub

Private Sub Form_Load()

    With frmUser.AdoUser.Recordset
    
        txtUsername = .Fields("username")
        txtType.Text = .Fields("type")
    
    End With

End Sub

