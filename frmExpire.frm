VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmExpire 
   Caption         =   "Warning Domains Expiring Soon"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy to Clipboard"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtExpire 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmExpire.frx":0000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtExpire2 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "The following domains are expiring soon."
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmExpire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.SetText Me.txtExpire.Text
End Sub

Private Sub Form_Resize()
On Error Resume Next
    txtExpire.Width = Me.Width - 100
    cmdClose.Left = txtExpire.Left + txtExpire.Width - 1600
End Sub
