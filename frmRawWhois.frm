VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRawWhois 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Whois Data Viewer"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FileList 
      Height          =   480
      Left            =   5760
      Pattern         =   "*.txt"
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox txtRawData 
      Height          =   4155
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7329
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmRawWhois.frx":0000
   End
   Begin VB.ListBox lstDomains 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblDate 
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Click on a domain to load the whois information."
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label lblWhois 
      Caption         =   "Whois Information:"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblDomains 
      Caption         =   "Domains:"
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
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmRawWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FileList.Path = App.Path & "\rawwhois\"
    
    Dim i As Long
    For i = 0 To FileList.ListCount - 1
        lstDomains.AddItem Replace(FileList.List(i), ".txt", "")
    Next
End Sub

Private Sub lstDomains_Click()
On Error GoTo errHandle
    If FileExists(App.Path & "\rawwhois\" & lstDomains.List(lstDomains.ListIndex) & ".txt") = True Then
        Dim data As String, strData As String
        Dim f As Long
        f = FreeFile
        lblDate.Caption = "Date: " & FileDateTime(App.Path & "\rawwhois\" & lstDomains.List(lstDomains.ListIndex) & ".txt")
        Open App.Path & "\rawwhois\" & lstDomains.List(lstDomains.ListIndex) & ".txt" For Input As #f
            Do While Not EOF(f)
                Line Input #f, data
                strData = strData & data & vbCrLf
            Loop
            
        Close #f
        txtRawData.Text = strData
    Else
        MsgBox "Raw whois data does not exists for that domain. Please rescan the domain and make sure you have raw whois data turned on in the options dialog.", vbInformation, "No Raw Whois Data"
    End If
Exit Sub
errHandle:
    MsgBox "Error: lstDomains_Click: " & Err.Number & " " & Err.Description

End Sub
