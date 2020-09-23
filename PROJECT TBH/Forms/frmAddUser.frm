VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD USER"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4620
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   615
      Left            =   3240
      Picture         =   "frmAddUser.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   615
      Left            =   2040
      Picture         =   "frmAddUser.frx":3E84
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H00F0F0F0&
      Height          =   465
      Left            =   2025
      TabIndex        =   4
      Top             =   2025
      Width           =   2265
   End
   Begin VB.TextBox txtPass2 
      BackColor       =   &H00F0F0F0&
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2265
   End
   Begin VB.TextBox txtPass1 
      BackColor       =   &H00F0F0F0&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   900
      Width           =   2265
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00F0F0F0&
      Height          =   420
      Left            =   2025
      TabIndex        =   1
      Top             =   360
      Width           =   2265
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type"
      Height          =   420
      Left            =   315
      TabIndex        =   7
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type Password:"
      Height          =   465
      Left            =   315
      TabIndex        =   6
      Top             =   1575
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   375
      Left            =   315
      TabIndex        =   5
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   285
      Left            =   315
      TabIndex        =   0
      Top             =   495
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAdd_Click()

    With frmUser
    
        .AdoUser.Recordset.AddNew
    
        
        
            If txtPass1.Text = txtPass2.Text Then
            
                If Trim(txtUsername) = "'" Then
                
                      MsgBox "Please put a valid USER NAME!", vbExclamation, "INVALID"
                    
                Else
                
                    
                     .AdoUser.Recordset.Fields("username") = txtUsername.Text
                     .AdoUser.Recordset.Fields("password") = txtPass2.Text
                     .AdoUser.Recordset.Fields("type") = txtType.Text
                     .AdoUser.Recordset.Update
                     
                     MsgBox "User was added successfully.", vbInformation, "Process Completed"
                    
                     .DataGrid1.ReBind
                     .AdoUser.Refresh
                     
                     Unload Me
                
                End If
            
            Else
            
                MsgBox "Incorrect Password!", vbExclamation, "Invalid"
                txtPass1.Text = ""
                txtPass2.Text = ""
                txtPass1.SetFocus
              
            End If
    
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub
