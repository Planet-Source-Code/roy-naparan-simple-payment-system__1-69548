VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "THE BLAZING HERALD v 1.0"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   12450
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":3482
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6165
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2569B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2702D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":289BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A351
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn1"
            Object.ToolTipText     =   "Add Student"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn2"
            Object.ToolTipText     =   "Search Record"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn3"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn5"
            Object.ToolTipText     =   "View Collection"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn4"
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   5
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmMain.frx":2BCE3
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Picture         =   "frmMain.frx":2BE45
            Text            =   "USER:"
            TextSave        =   "USER:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6579
            MinWidth        =   2822
            Key             =   "stbUser"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6297
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmMain.frx":2CED7
            Text            =   "TIME:"
            TextSave        =   "TIME:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "6:40 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuUser 
         Caption         =   "User Accounts"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Log Off"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Student"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View Records"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCollection 
         Caption         =   "View Collection"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnupop 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAddStudent 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

    frmLoader.Show vbModal
    Me.Show
    isLoginOk = False
    frmLogin.Show 1
    
    If isLoginOk = False Then Unload Me: Exit Sub
    
        DE1.DBConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\Database\TBH.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"

End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        PopupMenu mnupop
    
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

       If isLoginOk = False Then Exit Sub

        Dim reply As Integer
        
            reply = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "EXIT")
            
                If reply = vbYes Then
                
                    End
                
                Else
                
                    Cancel = 1
                    
                End If
                
End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show 1

End Sub

Private Sub mnuAdd_Click()

    frmAdd.Show

End Sub

Private Sub mnuAddStudent_Click()

    frmAdd.Show

End Sub

Private Sub mnuCollection_Click()

    frmCollection.Show

End Sub

Private Sub mnuExit_Click()

    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "EXIT") = vbYes Then End

End Sub

Private Sub mnuHelp_Click()

    frmHelp.Show 1

End Sub

Private Sub mnuLog_Click()

    frmLogin.Show

End Sub

Private Sub mnuSearch_Click()

    frmSearch.Show

End Sub

Private Sub mnuUser_Click()

    frmUser.Show

End Sub

Private Sub mnuView_Click()

    frmSearch.Show

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    
        Case "btn1"
            frmAdd.Show
          
        Case "btn2"
            frmSearch.Show
                   
        Case "btn4"
            Shell "Calc.exe", vbMaximizedFocus
            
        Case "btn5"
            frmCollection.Show
            
    End Select

End Sub
