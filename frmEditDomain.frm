VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditDomain 
   Caption         =   "Edit Domain Information"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRaw 
      Caption         =   "Raw Whois"
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin RichTextLib.RichTextBox txtRawData 
         Height          =   4095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7223
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmEditDomain.frx":0000
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label lblDate 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdViewRawWhois 
      Caption         =   "&Raw Whois Informaiton"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtEditDomain 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdEditDomain 
      Caption         =   "&Edit Domain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblEditDomain 
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "frmEditDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub EditDomain(index As Integer)
On Error Resume Next
    Dim i As Long
    For i = 1 To Me.txtEditDomain.count - 1
        Unload txtEditDomain(i)
        Unload lblEditDomain(i)
    Next
    
    For i = 1 To frmMain.lstDomains.ColumnHeaders.count
        If i = 1 Then
            lblEditDomain(0).Caption = frmMain.lstDomains.ColumnHeaders(1).Text
            txtEditDomain(0).Enabled = False
            txtEditDomain(0).Text = frmMain.lstDomains.ListItems(index).Text
        Else
            Load Me.lblEditDomain(i - 1)
            With Me.lblEditDomain(i - 1)
                .Caption = frmMain.lstDomains.ColumnHeaders(i).Text
                .Top = lblEditDomain(i - 2).Top + 300
                .Tag = index
                .Visible = True
            End With
            Load Me.txtEditDomain(i - 1)
            
            With Me.txtEditDomain(i - 1)
                .Enabled = True
                .Top = txtEditDomain(i - 2).Top + 300
                .Visible = True
                .Text = frmMain.lstDomains.ListItems(index).SubItems(i - 1)
                .Tag = (i - 1)
            End With
            
        End If
    Next
End Sub

Private Sub cmdEditDomain_Click()
On Error Resume Next
    Dim i As Long
    For i = 1 To txtEditDomain.count - 1
        frmMain.lstDomains.ListItems(Int(lblEditDomain(2).Tag)).SubItems(txtEditDomain(i).Tag) = txtEditDomain(i).Text
       ' frmMain.lstDomains.ListItems(1).SubItems(1) = "Agag"
    Next
    MsgBox "Domain Edited", vbInformation
End Sub

Private Sub cmdHide_Click()
    FrameRaw.Visible = False
End Sub

Private Sub cmdViewRawWhois_Click()
On Error GoTo errHandle
    If FileExists(App.Path & "\rawwhois\" & txtEditDomain.ITem(0).Text & ".txt") = True Then
        Dim data As String, strData As String
        Dim f As Long
        f = FreeFile
        lblDate.Caption = "Date: " & FileDateTime(App.Path & "\rawwhois\" & txtEditDomain.ITem(0).Text & ".txt")

        Open App.Path & "\rawwhois\" & txtEditDomain.ITem(0).Text & ".txt" For Input As #f
            Do While Not EOF(f)
                Line Input #f, data
                strData = strData & data & vbCrLf
            Loop
            
        Close #f
        FrameRaw.Visible = True
        txtRawData.Text = strData
    Else
        MsgBox "Raw whois data does not exists for that domain. Please rescan the domain and make sure you have raw whois data turned on in the options dialog.", vbInformation, "No Raw Whois Data"
    End If
Exit Sub
errHandle:
    MsgBox "Error: cmdViewRawWhois: " & Err.Number & " " & Err.Description
End Sub
