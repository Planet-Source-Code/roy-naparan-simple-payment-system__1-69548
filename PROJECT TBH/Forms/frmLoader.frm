VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoader 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Picture         =   "frmLoader.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5580
      Top             =   3420
   End
   Begin MSComctlLib.ProgressBar pbLoader 
      Height          =   150
      Left            =   810
      TabIndex        =   0
      Top             =   4185
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblPercent 
      BackStyle       =   0  'Transparent
      Caption         =   "12 %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   225
      TabIndex        =   2
      Top             =   4140
      Width           =   645
   End
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   810
      TabIndex        =   1
      Top             =   3915
      Width           =   4290
   End
End
Attribute VB_Name = "frmLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()

    pbLoader.Value = pbLoader.Value + 1
        
        lblPercent.Caption = pbLoader.Value / pbLoader.Max * 100 & " %"
        
        
            If pbLoader.Value > 1 And pbLoader.Value < 20 Then
            
                        lblLoad.Caption = "Loading Program...."
                    
            Else
                    If pbLoader.Value > 20 And pbLoader.Value < 35 Then
                    
                        lblLoad.Caption = "Loading Components...."
                        
            Else
                    If pbLoader.Value > 35 And pbLoader.Value < 45 Then
                    
                        lblLoad.Caption = "Building Tools...."
                        
            End If
            End If
            End If
            
            If pbLoader.Value = 70 Then
            
                    lblLoad.Caption = "loading Database...."
                    Timer1.Interval = 150
                    
            End If
            
            If pbLoader.Value = 80 Then
            
                    Timer1.Interval = 100
                    lblLoad.Caption = "Executing Program...."
                    
            End If
            
            If pbLoader.Value = pbLoader.Max Then
            
                    Timer1.Interval = 0
                    
                    
                    Call DBConnection

                    Unload Me
            End If

End Sub
