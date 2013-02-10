VERSION 5.00
Begin VB.Form frmTrayIcon 
   Caption         =   "Tray Icon"
   ClientHeight    =   30
   ClientLeft      =   -40305
   ClientTop       =   105
   ClientWidth     =   2340
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   30
   ScaleWidth      =   2340
   Visible         =   0   'False
   Begin VB.Menu mnuMenu 
      Caption         =   "File"
      Begin VB.Menu mnuHideDomainManager 
         Caption         =   "Hide Domain Manager Pro"
      End
      Begin VB.Menu mnuFileExpire 
         Caption         =   "Check for Expiring Domains"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'Domain Manager Pro
'VisualBasicZone.com 2006
'Jonathan Valentin
'********************************************
Option Explicit
#Const Trial = False
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONDBLCLK = &H203


Private Sub Form_Load()

    Me.Hide
    MySysTray.PopUpMessage = "Domain Manager Pro"
    MySysTray.Initialize Me.hWnd, Me.Icon, MySysTray.PopUpMessage
    MySysTray.ShowIcon
    Me.Hide
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim msgCallBackMessage As Long
  msgCallBackMessage = x / Screen.TwipsPerPixelX
  Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
      MySysTray.TipText = MySysTray.PopUpMessage
   Case WM_RBUTTONDOWN
        Me.PopupMenu mnuMenu
   Case WM_LBUTTONDBLCLK

        frmMain.Show
        IsInTray = False

   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MySysTray.HideIcon
    End
End Sub



Private Sub mnuExit_Click()
    Dim Response As VbMsgBoxResult
    Response = MsgBox("Are you sure you want to quit?", vbYesNo + vbInformation, "Quit Domain Manager Pro?")
    If Response = vbYes Then
        If frmMain.txtAccessKey.Text <> "" Then
            SaveSetting "Domain Manager Pro", "Options", "AlexaID", frmMain.txtAccessKey.Text
        Else
            SaveSetting "Domain Manager Pro", "Options", "AlexaID", frmMain.txtAlexaKey.Text
        End If
        If frmMain.txtAlexaSig.Text <> "" Then
            SaveSetting "Domain Manager Pro", "Options", "AlexaSig", frmMain.txtAlexaSig.Text
        Else
            SaveSetting "Domain Manager Pro", "Options", "AlexaSig", frmMain.txtAlexaSecretKey.Text
        End If
        Unload frmOptions
        Unload frmMain
        Unload Me
        Call MruList.Save
        End
    End If
End Sub





Private Sub mnuFileExpire_Click()
    Call frmMain.CheckDomainExpire
End Sub

Private Sub mnuHideDomainManager_Click()

    If mnuHideDomainManager.Caption = "Hide Domain Manager Pro" Then
        Unload frmOptions
        frmMain.Hide
        IsInTray = True
        frmTrayIcon.mnuHideDomainManager.Caption = "Show Domain Manager Pro"
    Else
        frmMain.Show
        IsInTray = False
        frmTrayIcon.mnuHideDomainManager.Caption = "Hide Domain Manager Pro"

    End If
End Sub

Private Sub mnuOptions_Click()

    frmOptions.Show
End Sub


