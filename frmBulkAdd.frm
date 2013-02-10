VERSION 5.00
Begin VB.Form frmBulkAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bulk Add Domains"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDomains 
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddDomains 
      Caption         =   "&Add Domains"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblDomains 
      Caption         =   "Enter Domain Names - One domain per line."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBulkAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddDomains_Click()
    If txtDomains.Text = "" Then
        MsgBox "You need to enter a domain name.", vbInformation
        Exit Sub
    Else
        txtDomains.Text = Replace(txtDomains.Text, "http://www", "")
        txtDomains.Text = Replace(txtDomains.Text, "http://", "")
        Dim strTemp() As String
        Dim i As Long, g As Long
        strTemp = Split(txtDomains.Text, vbCrLf)
        Dim blnGood As Boolean
        For g = 0 To UBound(strTemp)
            blnGood = True
            'Check if domain is already in list
            For i = 1 To frmMain.lstDomains.ListItems.Count
                If frmMain.lstDomains.ListItems.Item(i).Text = strTemp(g) Then
                    MsgBox "You already have that domain listed. For: " & strTemp(g), vbInformation
                    blnGood = False
                End If
            Next
            'Sanity checks
            If InStr(strTemp(g), ".") = False Then
                MsgBox "You have entered an invalid domain name. For:" & strTemp(g), vbInformation
                blnGood = False
            End If
    
            If blnGood = True Then
                frmMain.lstDomains.ListItems.Add , , strTemp(g)
            End If
        Next g
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
