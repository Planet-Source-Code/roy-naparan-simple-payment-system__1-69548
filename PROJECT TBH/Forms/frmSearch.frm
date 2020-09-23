VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEARCH"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   11820
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   2640
      Picture         =   "frmSearch.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   1920
      Picture         =   "frmSearch.frx":4504
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   1200
      Picture         =   "frmSearch.frx":5586
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   480
      Picture         =   "frmSearch.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8760
      Width           =   615
   End
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
      Height          =   855
      Left            =   10080
      Picture         =   "frmSearch.frx":768A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      Picture         =   "frmSearch.frx":870C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&REFRESH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      Picture         =   "frmSearch.frx":978E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8520
      Width           =   1335
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
      Height          =   855
      Left            =   5760
      Picture         =   "frmSearch.frx":A810
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   1335
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
      Height          =   855
      Left            =   4320
      Picture         =   "frmSearch.frx":B892
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   855
      Left            =   4680
      Picture         =   "frmSearch.frx":C914
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click to Search"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   -135
      Picture         =   "frmSearch.frx":FD96
      ScaleHeight     =   2355
      ScaleWidth      =   12030
      TabIndex        =   5
      Top             =   -180
      Width           =   12030
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00F0F0F0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   450
      TabIndex        =   1
      Top             =   3060
      Width           =   1995
   End
   Begin VB.ComboBox cmbYear 
      BackColor       =   &H00F0F0F0&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearch.frx":1B7A1
      Left            =   2610
      List            =   "frmSearch.frx":1B7B4
      TabIndex        =   3
      Top             =   3060
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2610
      TabIndex        =   2
      Top             =   2565
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   71761921
      CurrentDate     =   39379
   End
   Begin VB.ComboBox cmbSelect 
      BackColor       =   &H00F0F0F0&
      Height          =   315
      ItemData        =   "frmSearch.frx":1B7EA
      Left            =   450
      List            =   "frmSearch.frx":1B7FA
      TabIndex        =   0
      Top             =   2565
      Width           =   1995
   End
   Begin MSAdodcLib.Adodc AdoStudent 
      Height          =   375
      Left            =   630
      Top             =   9180
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
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
      Height          =   4695
      Left            =   450
      TabIndex        =   4
      Top             =   3600
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   8281
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "StudentID"
         Caption         =   "StudentID"
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
         DataField       =   "StudentNo"
         Caption         =   "Student No."
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
         DataField       =   "LName"
         Caption         =   "Last Name"
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
         DataField       =   "GName"
         Caption         =   "First Name"
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
      BeginProperty Column04 
         DataField       =   "Course"
         Caption         =   "Course"
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
      BeginProperty Column05 
         DataField       =   "Year"
         Caption         =   "Year"
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
      BeginProperty Column06 
         DataField       =   "Date"
         Caption         =   "Date of Payment"
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
      BeginProperty Column07 
         DataField       =   "Amount"
         Caption         =   "Amount"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10440
      TabIndex        =   8
      Top             =   2835
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Students:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8325
      TabIndex        =   7
      Top             =   2835
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search for?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   495
      TabIndex        =   6
      Top             =   2295
      Width           =   1320
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmbSelect_Click()

    If cmbSelect.Text = "Last Name" Then
    
        txtName.Enabled = True
        cmbYear.Enabled = False
        dtpDate.Enabled = False
        
        
    ElseIf cmbSelect.Text = "Date of Payment" Then
    
        dtpDate.Enabled = True
        cmbYear.Enabled = False
        txtName.Enabled = False
        
        
    ElseIf cmbSelect.Text = "Year" Then
    
        cmbYear.Enabled = True
        txtName.Enabled = False
        dtpDate.Enabled = False
        
    ElseIf cmbSelect.Text = "View All" Then
    
        txtName.Enabled = False
        dtpDate.Enabled = False
        cmbYear.Enabled = False
        
        AdoStudent.Refresh
        
    End If
    

End Sub

Private Sub cmbYear_Click()

            
                With AdoStudent.Recordset
                
                    .MoveFirst
                    .Filter = "Year= '" & (cmbYear.Text) & "'"
                    
                        If .EOF Then
                        
                           MsgBox "Year not found!", vbExclamation, "Filter Failed"
                           
                        End If
                        
                End With
  

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    If MsgBox("Are you sure you to delete this record?", vbQuestion + vbYesNo, "DELETE") = vbYes Then
    
        
        With AdoStudent.Recordset
        
            .Delete
            .MoveNext
            
        End With
        
        MsgBox "Student record was successfully deleted!", vbInformation, "DELETED"
        
    End If

End Sub

Private Sub cmdEdit_Click()

    frmEdit.Show

End Sub

Private Sub cmdPrint_Click()

    Set rptStudentRecords.DataSource = AdoStudent.Recordset
    rptStudentRecords.Show
    
    DataGrid1.SetFocus

End Sub

Private Sub cmdRefresh_Click()

    AdoStudent.Refresh

End Sub

Private Sub cmdSearch_Click()

    Dim strFind As String
    strFind = txtName.Text
    
        If strFind = "" Then Exit Sub
        
            If cmbSelect.Text = "Last Name" Then
            
                With AdoStudent.Recordset
                
                    .MoveFirst
                    .Find "LName like '" & strFind & "%'"
                    
                        If .EOF Then
                        
                            MsgBox "Record not found!", vbExclamation, "Search Failed"
                            
                        End If
                    
                End With
            
        End If
        
       
    
    txtName.Text = ""

End Sub

Private Sub Command1_Click()

    AdoStudent.Recordset.MoveFirst

End Sub

Private Sub Command2_Click()

    AdoStudent.Recordset.MovePrevious
    
    If AdoStudent.Recordset.BOF Then
    
        AdoStudent.Recordset.MoveFirst
        
    End If

End Sub

Private Sub Command3_Click()

    AdoStudent.Recordset.MoveNext
    
    If AdoStudent.Recordset.EOF Then
    
        AdoStudent.Recordset.MoveLast
        
    End If

End Sub

Private Sub Command4_Click()

    AdoStudent.Recordset.MoveLast

End Sub

Private Sub dtpDate_Change()
            
                With AdoStudent.Recordset
                
                    .MoveFirst
                    .Filter = "Date= '" & (dtpDate.Value) & "'"
                    
                        If .EOF Then
                        
                            MsgBox "Date not found!", vbExclamation, "Filter Failed"
                            
                        End If
                
                End With

End Sub

Private Sub Form_Load()

    AdoStudent.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\Database\TBH.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    AdoStudent.RecordSource = "SELECT * FROM tbl_Student"
    AdoStudent.Refresh
    
    
        Set DataGrid1.DataSource = AdoStudent
        lblNo.Caption = AdoStudent.Recordset.RecordCount

End Sub
