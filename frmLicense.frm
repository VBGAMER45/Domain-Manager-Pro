VERSION 5.00
Begin VB.Form frmLicense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register Domain Manager Pro"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtUnlock 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmLicense.frx":6852
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "UnLock Code:"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tries As Long
Private Sub cmdClose_Click()
Exit Sub
    Unload Me
End Sub

Private Sub cmdRegister_Click()
Dim strUnLockKey As String
    strUnLockKey = "Unlock Key: DM80950968456"
    strUnLockKey = "Demo Key: DM2-095-96-3h7f"
    strUnLockKey = "Corporate Key: ENT1S546AVXT4"
    If Tries > 5 Then
        MsgBox "You had too many tries!!! :(", vbExclamation
        
        End
    End If
    If txtUnlock.Text = "" Then
        MsgBox "You did not enter an unlock code.", vbInformation
    Else
        If txtUnlock.Text = "DM80950968456" Then
            MsgBox "Thanks for registering!", vbExclamation
            End
            Exit Sub
        End If
        If txtUnlock.Text = "ENT1S546AVXT4" Then
            MsgBox "Thanks for registering!", vbExclamation
            End
            Exit Sub
        End If
        
        Dim bad As Boolean
        bad = False
        If Len(txtUnlock.Text) > 16 Then
            bad = True
             MsgBox "Invalid Code", vbInformation
             Tries = Tries + 1
            Exit Sub
        End If
        If Len(txtUnlock.Text) < 16 Then
            bad = True
             MsgBox "Invalid Code", vbInformation
             Tries = Tries + 1
            Exit Sub
        End If
        If Mid(txtUnlock.Text, 4, 1) <> 1 Then
            MsgBox "Invalid Code", vbInformation
            Tries = Tries + 1
            bad = True
            Exit Sub
        End If
        
        If Left$(txtUnlock.Text, 3) = "DMP" And bad = False Then
            'DMP1AAAAAAAAAABB
            MsgBox "Thank You for Registering and supporting this software.", vbExclamation
            
            SaveSetting "Domain Manager Pro", "Options", "Yup", "!"
            Unload Me
        Else
            MsgBox "Invalid Code", vbInformation
            Tries = Tries + 1
            Exit Sub
        
        End If
    
    End If
End Sub

Private Sub Form_Load()
Tries = 0
End Sub
