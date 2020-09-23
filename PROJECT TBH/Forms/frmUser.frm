VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USERS"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Picture         =   "frmUser.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Picture         =   "frmUser.frx":4504
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Picture         =   "frmUser.frx":5586
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frmUser.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   -90
      Picture         =   "frmUser.frx":768A
      ScaleHeight     =   2355
      ScaleWidth      =   6900
      TabIndex        =   1
      Top             =   0
      Width           =   6900
   End
   Begin MSAdodcLib.Adodc AdoUser 
      Height          =   330
      Left            =   3060
      Top             =   6345
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2940
      Left            =   405
      TabIndex        =   0
      Top             =   2475
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   5186
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "UserID"
         Caption         =   "UserID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "UserName"
         Caption         =   "User Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Password"
         Caption         =   "Password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Type"
         Caption         =   "User Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    
    frmAddUser.Show

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "DELETE") = vbYes Then
    
        If InputBox("Enter Account Password:", "PASSWORD") = AdoUser.Recordset.Fields("password") Then
        
            With AdoUser.Recordset
                .Delete
                .MoveNext
            End With
        
        Else
        
            MsgBox "Invalid password", vbExclamation, "INVALID"
            
        End If
        
    End If

End Sub

Private Sub cmdEdit_Click()

    frmEditUser.Show

End Sub

Private Sub Form_Load()

    AdoUser.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\Database\TBH.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    AdoUser.RecordSource = "select * from tbl_user"
    AdoUser.Refresh
    
        Set DataGrid1.DataSource = AdoUser
        DataGrid1.ReBind

End Sub
