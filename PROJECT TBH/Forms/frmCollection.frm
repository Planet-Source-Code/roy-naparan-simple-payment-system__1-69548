VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCollection 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COLLECTION"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCollection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7065
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   2520
      Picture         =   "frmCollection.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   1800
      Picture         =   "frmCollection.frx":4504
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   1080
      Picture         =   "frmCollection.frx":5586
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   360
      Picture         =   "frmCollection.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   735
      Left            =   5280
      Picture         =   "frmCollection.frx":768A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   735
      Left            =   3840
      Picture         =   "frmCollection.frx":870C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&VIEW ALL"
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   -45
      Picture         =   "frmCollection.frx":978E
      ScaleHeight     =   2310
      ScaleWidth      =   7845
      TabIndex        =   3
      Top             =   0
      Width           =   7845
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   2385
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Format          =   17694721
      CurrentDate     =   39379
   End
   Begin MSAdodcLib.Adodc AdoTotal 
      Height          =   330
      Left            =   585
      Top             =   5310
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
      Height          =   3075
      Left            =   360
      TabIndex        =   0
      Top             =   2925
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5424
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CountOfStudentNo"
         Caption         =   "Number of Students"
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
      BeginProperty Column02 
         DataField       =   "Amount"
         Caption         =   "Total Amount"
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
            Locked          =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date:"
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   2430
      Width           =   1230
   End
End
Attribute VB_Name = "frmCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()

    Unload Me

End Sub


Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdPrint_Click()

    Set rptTotal.DataSource = AdoTotal.Recordset
    
        rptTotal.Show

End Sub

Private Sub cmdView_Click()

    AdoTotal.Refresh

End Sub

Private Sub Command1_Click()

    AdoTotal.Recordset.MoveFirst

End Sub

Private Sub Command2_Click()

    AdoTotal.Recordset.MovePrevious
    
        If AdoTotal.Recordset.BOF Then
        
            AdoTotal.Recordset.MoveFirst
            
        End If

End Sub

Private Sub Command3_Click()

    AdoTotal.Recordset.MoveNext
    
        If AdoTotal.Recordset.EOF Then
        
            AdoTotal.Recordset.MoveLast
            
        End If
    
End Sub

Private Sub Command4_Click()

    AdoTotal.Recordset.MoveLast

End Sub

Private Sub DTPicker1_Change()

       With AdoTotal.Recordset
        
        .MoveFirst
        .Filter = "Date= '" & (DTPicker1.Value) & "'"
       
            If .EOF Then
            
                MsgBox "Date not found", vbInformation, "FILTER FAILED"
                AdoTotal.Refresh
                
            End If
        
    End With

End Sub

Private Sub Form_Load()

    AdoTotal.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\Database\TBH.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    AdoTotal.RecordSource = "SELECT * FROM qry_Total"
    AdoTotal.Refresh
    
    
        Set DataGrid1.DataSource = AdoTotal

End Sub

