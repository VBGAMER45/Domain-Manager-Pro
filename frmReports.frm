VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "&Generate Report"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Report Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdChangeFolder 
         Caption         =   "..."
         Height          =   255
         Left            =   6120
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtReportDirectory 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label lblReportDirectory 
         Caption         =   "Report Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
    End Type

Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Sub cmdChangeFolder_Click()
    Dim sPath As String
    Dim structFolder As BROWSEINFO
    Dim iNull As Integer
    Dim ret As Long
    structFolder.hOwner = Me.hWnd
    structFolder.lpszTitle = "Browse for folder"
    structFolder.ulFlags = BIF_NEWDIALOGSTYLE  'To create make new folder option

    ret = SHBrowseForFolder(structFolder)
    If ret Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList ret, sPath
        'free the block of memory
        CoTaskMemFree ret
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    If sPath = vbNullString Then Exit Sub
    If Right$(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
   
    txtReportDirectory.Text = sPath

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGenerateReport_Click()
On Error GoTo errHandle
    If txtReportDirectory.Text = "" Then
        MsgBox "You need to enter the report directory!", vbInformation
        Exit Sub
    End If
    
    Dim f As Long, i As Long, g As Long
    f = FreeFile
    
    'Select the report type
    Select Case Me.Tag
    
        Case 1
            Open txtReportDirectory.Text & "domains.html" For Output As #f
                Print #f, "<html>"
                Print #f, "<head><title>Domain Report</title></head><body>"
                Print #f, "<h1>Domain Report</h1><br>"
                Print #f, "<table border=" & Chr(34) & "1" & Chr(34) & " width=" & Chr(34) & "95%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">"
                Print #f, "<tr>"
                For i = 1 To frmMain.lstDomains.ColumnHeaders.Count
                    Print #f, "<td>" & frmMain.lstDomains.ColumnHeaders.Item(i).Text & "</td>"
                Next
                Print #f, "</tr>"
                For i = 1 To frmMain.lstDomains.ListItems.Count
                    Print #f, "<tr>"
                    Print #f, "<td>" & frmMain.lstDomains.ListItems.Item(i).Text & "</td>"
                    For g = 1 To frmMain.lstDomains.ListItems.Item(i).ListSubItems.Count
                        Print #f, "<td>" & frmMain.lstDomains.ListItems.Item(i).ListSubItems(g).Text & "</td>"
                    Next
                    Print #f, "</tr>"
                Next
                Print #f, "</table>"
                Print #f, "<p>Generated by <a href=" & Chr(34) & "http://www.dnmanagerpro.com" & Chr(34) & ">Domain Manager Pro</a></p>"
                Print #f, "</body></html>"
                
            Close #f
            ShellExecute Me.hWnd, vbNullString, txtReportDirectory.Text & "domains.html", vbNullString, "C:\", SW_SHOWNORMAL

        Case 2
            Open txtReportDirectory.Text & "results.html" For Output As #f
                Print #f, "<html>"
                Print #f, "<head><title>Results Report</title></head><body>"""
                Print #f, "<h1>Results Report</h1><br>"
                Print #f, "<table border=" & Chr(34) & "1" & Chr(34) & " width=" & Chr(34) & "95%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">"
                Print #f, "<tr>"
                For i = 1 To frmMain.lstResults.ColumnHeaders.Count
                    Print #f, "<td>" & frmMain.lstResults.ColumnHeaders.Item(i).Text & "</td>"
                Next
                Print #f, "</tr>"
                For i = 1 To frmMain.lstResults.ListItems.Count
                    Print #f, "<tr>"
                    Print #f, "<td>" & frmMain.lstResults.ListItems.Item(i).Text & "</td>"
                    For g = 1 To frmMain.lstResults.ListItems.Item(i).ListSubItems.Count
                        Print #f, "<td>" & frmMain.lstResults.ListItems.Item(i).ListSubItems(g).Text & "</td>"
                    Next
                    Print #f, "</tr>"
                Next
                Print #f, "</table>"
                Print #f, "<p>Generated by <a href=" & Chr(34) & "http://www.dnmanagerpro.com" & Chr(34) & ">Domain Manager Pro</a></p>"
                Print #f, "</body></html>"
            Close #f
           ShellExecute Me.hWnd, vbNullString, txtReportDirectory.Text & "results.html", vbNullString, "C:\", SW_SHOWNORMAL
        
        Case 3
            Open txtReportDirectory.Text & "alexa.html" For Output As #f
                Print #f, "<html>"
                Print #f, "<head><title>Alexa Report</title></head><body>"
                Print #f, "<h1>Alexa Report</h1><br>"
                Print #f, "<table border=" & Chr(34) & "1" & Chr(34) & " width=" & Chr(34) & "95%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">"
                Print #f, "<tr>"
                For i = 1 To frmMain.lstAlexa.ColumnHeaders.Count
                    Print #f, "<td>" & frmMain.lstAlexa.ColumnHeaders.Item(i).Text & "</td>"
                Next
                Print #f, "</tr>"
                For i = 1 To frmMain.lstAlexa.ListItems.Count
                    Print #f, "<tr>"
                    Print #f, "<td>" & frmMain.lstAlexa.ListItems.Item(i).Text & "</td>"
                    For g = 1 To frmMain.lstAlexa.ListItems.Item(i).ListSubItems.Count
                        Print #f, "<td>" & frmMain.lstAlexa.ListItems.Item(i).ListSubItems(g).Text & "</td>"
                    Next
                    Print #f, "</tr>"
                Next
                Print #f, "</table>"
                Print #f, "<p>Generated by <a href=" & Chr(34) & "http://www.dnmanagerpro.com" & Chr(34) & ">Domain Manager Pro</a></p>"
                Print #f, "</body></html>"
            Close #f
            ShellExecute Me.hWnd, vbNullString, txtReportDirectory.Text & "alexa.html", vbNullString, "C:\", SW_SHOWNORMAL

    End Select
Exit Sub
errHandle:
    MsgBox "Error: cmdGenerateReport: " & Err.Number & " " & Err.Description
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtReportDirectory.Text = App.Path & "\reports\"
    MkDir (App.Path & "\reports")
    
    
End Sub

