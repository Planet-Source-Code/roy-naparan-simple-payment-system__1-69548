VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   735
      Left            =   4320
      Picture         =   "frmAdd.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   735
      Left            =   2880
      Picture         =   "frmAdd.frx":4504
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   -135
      Picture         =   "frmAdd.frx":5586
      ScaleHeight     =   2310
      ScaleWidth      =   6540
      TabIndex        =   14
      Top             =   0
      Width           =   6540
   End
   Begin VB.TextBox txtStudentNo 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1620
      TabIndex        =   0
      Top             =   2475
      Width           =   1680
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   1575
      TabIndex        =   5
      Top             =   5445
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   17694721
      CurrentDate     =   39379
   End
   Begin VB.TextBox txtCourse 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1575
      TabIndex        =   6
      Top             =   6030
      Width           =   1905
   End
   Begin VB.ComboBox cmbLevel 
      BackColor       =   &H00F0F0F0&
      Height          =   315
      ItemData        =   "frmAdd.frx":1000D
      Left            =   1575
      List            =   "frmAdd.frx":10020
      TabIndex        =   4
      Top             =   4950
      Width           =   1905
   End
   Begin VB.TextBox txtFName 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   3735
      Width           =   3930
   End
   Begin VB.TextBox txtLName 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   3105
      Width           =   3930
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Student No."
      Height          =   240
      Left            =   315
      TabIndex        =   13
      Top             =   2610
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   240
      Left            =   360
      TabIndex        =   12
      Top             =   5580
      Width           =   1410
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   6165
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "College Level"
      Height          =   240
      Left            =   315
      TabIndex        =   10
      Top             =   5040
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      Height          =   240
      Left            =   315
      TabIndex        =   9
      Top             =   4365
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Left            =   315
      TabIndex        =   8
      Top             =   3825
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   240
      Left            =   315
      TabIndex        =   7
      Top             =   3240
      Width           =   1050
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdClose_Click()


    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    If txtStudentNo.Text = "" Then
    
        MsgBox "Student number is required!", vbInformation, "INVALID"
        
    Else
    
        If txtLName.Text = "" Then
    
            MsgBox "Last name is required!", vbInformation, "INVALID"
        
    Else
    
        If txtFName.Text = "" Then
    
            MsgBox "First name is required!", vbInformation, "INVALID"
        
    Else
    
        If txtCourse.Text = "" Then
    
            MsgBox "Course is required!", vbInformation, "INVALID"
        
    Else
    
        If cmbLevel.Text = "" Then
    
            MsgBox "Year is required!", vbInformation, "INVALID"
        
    Else
    
        If txtAmount.Text = "" Then
    
            MsgBox "Amount is required!", vbInformation, "INVALID"
        
    Else
    
            rs.Open "select * from tbl_student", conn, adOpenDynamic, adLockBatchOptimistic
            
            With rs
            
                .AddNew
                .Fields("StudentNo") = txtStudentNo.Text
                .Fields("LName") = txtLName.Text
                .Fields("GName") = txtFName.Text
                .Fields("Course") = txtCourse.Text
                .Fields("Year") = cmbLevel.Text
                .Fields("Date") = dtpDate.Value
                .Fields("Amount") = txtAmount.Text
                .UpdateBatch
                .Close
                
            End With
        
            
            MsgBox "New student was added successfully!", vbInformation, "Process Completed"
            
            txtStudentNo.Text = ""
            txtLName.Text = ""
            txtFName.Text = ""
            txtCourse.Text = ""
            cmbLevel.Text = ""
            dtpDate.Value = Date
            txtAmount.Text = ""
        
        
    End If
    End If
    End If
    End If
    End If
    End If

End Sub

Private Sub Form_Load()

    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    Call DBConnection

End Sub

