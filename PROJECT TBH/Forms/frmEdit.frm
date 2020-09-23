VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmEdit.frx":3482
   ScaleHeight     =   7425
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   735
      Left            =   4440
      Picture         =   "frmEdit.frx":DCA2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   735
      Left            =   3000
      Picture         =   "frmEdit.frx":ED24
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   0
      Picture         =   "frmEdit.frx":FDA6
      ScaleHeight     =   2355
      ScaleWidth      =   6000
      TabIndex        =   14
      Top             =   0
      Width           =   6000
   End
   Begin VB.TextBox txtStudentNo 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1755
      TabIndex        =   0
      Top             =   2520
      Width           =   1680
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   1755
      TabIndex        =   5
      Top             =   5130
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      _Version        =   393216
      Format          =   17694721
      CurrentDate     =   39379
   End
   Begin VB.TextBox txtCourse 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1755
      TabIndex        =   3
      Top             =   4140
      Width           =   1815
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1755
      TabIndex        =   6
      Top             =   5625
      Width           =   1905
   End
   Begin VB.ComboBox cmbLevel 
      BackColor       =   &H00F0F0F0&
      Height          =   315
      ItemData        =   "frmEdit.frx":1A5C6
      Left            =   1755
      List            =   "frmEdit.frx":1A5D9
      TabIndex        =   4
      Top             =   4680
      Width           =   1860
   End
   Begin VB.TextBox txtFName 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1755
      TabIndex        =   2
      Top             =   3600
      Width           =   3930
   End
   Begin VB.TextBox txtLName 
      BackColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   1755
      TabIndex        =   1
      Top             =   3060
      Width           =   3930
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Student No."
      Height          =   240
      Left            =   405
      TabIndex        =   13
      Top             =   2610
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   240
      Left            =   450
      TabIndex        =   12
      Top             =   5220
      Width           =   1410
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   285
      Left            =   450
      TabIndex        =   11
      Top             =   5715
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "College Level"
      Height          =   240
      Left            =   405
      TabIndex        =   10
      Top             =   4770
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      Height          =   240
      Left            =   450
      TabIndex        =   9
      Top             =   4230
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Left            =   405
      TabIndex        =   8
      Top             =   3690
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   240
      Left            =   405
      TabIndex        =   7
      Top             =   3150
      Width           =   1050
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()

    With frmSearch.AdoStudent.Recordset
    
        .Fields("StudentNo") = txtStudentNo.Text
        .Fields("LName") = txtLName.Text
        .Fields("GName") = txtFName.Text
        .Fields("Course") = txtCourse.Text
        .Fields("Year") = cmbLevel.Text
        .Fields("Date") = dtpDate.Value
        .Fields("Amount") = txtAmount.Text
        .UpdateBatch
        
    End With
    
    MsgBox "Student record was updated successfully!", vbInformation, "Process Completed"
    
    Unload Me

End Sub

Private Sub Form_Load()

    With frmSearch.AdoStudent.Recordset
    
        txtStudentNo.Text = .Fields("StudentNo")
        txtLName.Text = .Fields("LName")
        txtFName.Text = .Fields("GName")
        txtCourse.Text = .Fields("Course")
        cmbLevel.Text = .Fields("Year")
        dtpDate.Value = .Fields("Date")
        txtAmount.Text = .Fields("Amount")
    
    End With

End Sub
