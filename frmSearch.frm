VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search Domains"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSearchResult 
      Caption         =   "Complete Search Result"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   2775
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
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
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox cboSearchCol 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Search is case sensitive."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label lblResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   3480
      Width           =   3000
   End
   Begin VB.Label lblSearchFor 
      Caption         =   "Search For:"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Search In:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then
        MsgBox "You need to enter something to search for!", vbInformation
        Exit Sub
    End If
    If cboSearchCol.Text = "" Then
        MsgBox "You did not select a column to search in.", vbInformation
        Exit Sub
    End If
    Dim i As Long
    Dim cId As Long

    lstSearch.ColumnHeaders.Clear
    lstSearch.ListItems.Clear
    'Get the Column Id
    For i = 1 To frmMain.lstDomains.ColumnHeaders.count
        If frmMain.lstDomains.ColumnHeaders(i) = cboSearchCol.Text Then
            lstSearch.ColumnHeaders.Add , , "Domain"
            lstSearch.ColumnHeaders.Add , , cboSearchCol.Text
            cId = i

            Exit For
        End If
    Next
    Dim g As Long
        'MsgBox cId
    For i = 1 To frmMain.lstDomains.ListItems.count
        If cId = 1 Then
            If InStr(1, frmMain.lstDomains.ListItems.ITem(i).Text, txtSearch.Text) Then
                lstSearch.ListItems.Add , , frmMain.lstDomains.ListItems.ITem(i).Text
                If chkSearchResult.Value = vbChecked Then
                    For g = 2 To frmMain.lstDomains.ColumnHeaders.count
                        
                    Next
                End If
            End If
        Else
            If InStr(1, frmMain.lstDomains.ListItems.ITem(i).SubItems(cId - 1), txtSearch.Text) Then
                With lstSearch.ListItems.Add(, , frmMain.lstDomains.ListItems.ITem(i).Text)
                    .SubItems(1) = frmMain.lstDomains.ListItems.ITem(i).SubItems(cId - 1)
                End With
                If chkSearchResult.Value = vbChecked Then
                    For g = 2 To frmMain.lstDomains.ColumnHeaders.count
             
                        If g <> cId Then
                        
                           ' frmMain.lstDomains.frmMain.lstDomains.ListItems.Item(i).SubItems (g - 1)
                        End If
                    Next
                End If

            End If
        End If
      '  frmMain.lstDomains.item(0).
    Next
    lblResults.Caption = "Total Records Found: " & Me.lstSearch.ListItems.count
End Sub


Private Sub Form_Load()
On Error Resume Next
    Dim i As Long
    Me.cboSearchCol.Clear
    Me.cboSearchCol.Text = frmMain.lstDomains.ColumnHeaders(1).Text
    For i = 1 To frmMain.lstDomains.ColumnHeaders.count
        Me.cboSearchCol.AddItem frmMain.lstDomains.ColumnHeaders(i).Text
    Next
    
End Sub



Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click

    End If
End Sub
