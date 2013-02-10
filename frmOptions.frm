VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkGetIPAddress 
      Caption         =   "Get IP Address of site on lookup. Takes longer"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Whois Servers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   19
      Top             =   120
      Width           =   4215
      Begin VB.Label Label6 
         Caption         =   "To edit Whois servers open the whois.ini located in the root directory of this software."
         Height          =   975
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CheckBox chkRawWhoisData 
      Caption         =   "Save Raw Whois Data"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CheckBox chkSaveDomains 
      Caption         =   "Auto save domains every minute."
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CheckBox chkStartUp 
      Caption         =   "Run Domain Manager on Start Up"
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "30"
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame FrameWhois 
      Caption         =   "Whois Servers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtDotNet 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtDotInfo 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox txtDotBiz 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtDotOrg 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CommandButton cmdDefaultWhois 
         Caption         =   "&Restore Default Whois Servers"
         Height          =   435
         Left            =   1200
         TabIndex        =   6
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtDotCom 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   ".Com Whois Server:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   ".Info Whois Server"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   ".Biz Whois Server"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   ".Org Whois Server"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   ".Net Whois Server:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Save Changes"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblDomainExpire 
      Caption         =   "Check if domains are expiring within how many days:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   1560
      Width           =   4455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGetIPAddress_Click()
    If Me.chkGetIPAddress.Value = vbChecked Then
        SaveSetting "Domain Manager Pro", "Options", "GetIP", "1"
        blnGetIPAddress = True
    Else
        SaveSetting "Domain Manager Pro", "Options", "GetIP", "0"
        blnGetIPAddress = False
    End If
End Sub

Private Sub chkRawWhoisData_Click()
    If chkRawWhoisData.Value = vbChecked Then
        Call SaveSetting("Domain Manager Pro", "Options", "RawWhois", "True")
        blnRawWhois = True
    Else
        Call SaveSetting("Domain Manager Pro", "Options", "RawWhois", "False")
        blnRawWhois = False
    End If
End Sub

Private Sub chkSaveDomains_Click()
    If chkSaveDomains.Value = vbChecked Then
        frmMain.tmrAutoSave.Enabled = True
    Else
        frmMain.tmrAutoSave.Enabled = False
    End If
End Sub

Private Sub chkStartUp_Click()
    If chkStartUp.Value = vbChecked Then
        Call modMain.RegRun(App.Path & "\DNManagerPro.exe tray", "DomainManagerPro")
        Call SaveSetting("Domain Manager Pro", "Options", "StartUp", "True")
        blnStartUp = True
    Else
        Call modMain.RemoveRegRun("DomainManagerPro")
        Call SaveSetting("Domain Manager Pro", "Options", "StartUp", "False")
        blnStartUp = False
    End If
        
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDefaultWhois_Click()
    Me.txtDotCom.Text = DotCom
    Me.txtDotNet.Text = DotNet
    Me.txtDotOrg.Text = DotOrg
    Me.txtDotBiz.Text = DotBiz
    Me.txtDotInfo.Text = DotInfo
End Sub

Private Sub cmdUpdate_Click()
    strWhoisDotCom = Me.txtDotCom.Text
    strWhoisDotNet = Me.txtDotNet.Text
    strWhoisDotOrg = Me.txtDotOrg.Text
    strWhoisDotBiz = Me.txtDotBiz.Text
    strWhoisDotInfo = Me.txtDotInfo.Text
    ExpireDays = Me.txtDay.Text
    
    If Me.chkGetIPAddress.Value = vbChecked Then
        SaveSetting "Domain Manager Pro", "Options", "GetIP", "1"
        blnGetIPAddress = True
    Else
        SaveSetting "Domain Manager Pro", "Options", "GetIP", "0"
        blnGetIPAddress = False
    End If
    
    SaveSetting "Domain Manager Pro", "Options", "DotCom", strWhoisDotCom
    SaveSetting "Domain Manager Pro", "Options", "DotNet", strWhoisDotNet
    SaveSetting "Domain Manager Pro", "Options", "DotOrg", strWhoisDotOrg
    SaveSetting "Domain Manager Pro", "Options", "DotBiz", strWhoisDotBiz
    SaveSetting "Domain Manager Pro", "Options", "DotInfo", strWhoisDotInfo
    SaveSetting "Domain Manager Pro", "Options", "ExpireCheck", txtDay.Text
    
    
    MsgBox "Changes Saved", vbInformation
    Unload Me
End Sub

Private Sub Form_Load()
'Load Servers
    Me.txtDotCom.Text = strWhoisDotCom
    Me.txtDotNet.Text = strWhoisDotNet
    Me.txtDotOrg.Text = strWhoisDotOrg
    Me.txtDotBiz.Text = strWhoisDotBiz
    Me.txtDotInfo.Text = strWhoisDotInfo
    
    Me.txtDay = ExpireDays
    If blnRawWhois = True Then Me.chkRawWhoisData.Value = vbChecked
    If frmMain.tmrAutoSave.Enabled = True Then Me.chkSaveDomains.Value = vbChecked
    If blnStartUp = True Then Me.chkStartUp.Value = vbChecked
    If blnGetIPAddress = True Then Me.chkGetIPAddress.Value = vbChecked
End Sub

Private Sub txtDay_Change()
    If IsNumeric(txtDay.Text) = False Then
        txtDay.Text = 30
    End If
End Sub
