VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Custom Quick Links"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Menu Item"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ListBox lstMenu 
      Height          =   1230
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   5415
   End
   Begin VB.TextBox txtEditMenuTitle 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4560
      Width           =   4455
   End
   Begin VB.TextBox txtEditMenuLink 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "http://"
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox txtMenuTitle 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtMenuLink 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "http://"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton cmdAddMenu 
      Caption         =   "&Add Menu Item"
      Height          =   495
      Left            =   2078
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditMenu 
      Caption         =   "&Edit Menu Item"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Menu Title:"
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
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Menu Link:"
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
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5640
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      Caption         =   $"frmMenu.frx":6852
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "Menu Link:"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Menu Title:"
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
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Edit Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblAddMenu 
      Caption         =   "Add Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddMenu_Click()
    If txtMenuTitle.Text = "" Then
        MsgBox "You did not enter a menu title.", vbInformation
        Exit Sub
    End If
    If txtMenuLink.Text = "http://" Then
        MsgBox "You did not enter a link.", vbInformation
        Exit Sub
    End If
    Dim i As Long
    For i = 0 To UBound(CustomMenu)
        If CustomMenu(i).MenuTitle = "" And CustomMenu(i).MenuLink = "" Then
            CustomMenu(i).MenuLink = txtMenuLink.Text
            CustomMenu(i).MenuTitle = txtMenuTitle.Text
            Call ReloadMenus
            Call frmMain.SaveMenus
            Call frmMain.ReloadMenuItems
            Exit Sub
        End If
    Next
    ReDim Preserve CustomMenu(UBound(CustomMenu) + 1)
    CustomMenu(UBound(CustomMenu)).MenuLink = txtMenuLink.Text
    CustomMenu(UBound(CustomMenu)).MenuTitle = txtMenuTitle.Text
    
    Call ReloadMenus
    Call frmMain.SaveMenus
    Call frmMain.ReloadMenuItems
    MsgBox "Menu Item Added", vbInformation
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub ReloadMenus()
    lstMenu.Clear
    
    Dim i As Long
    For i = 0 To UBound(CustomMenu)
        If CustomMenu(i).MenuTitle <> "" Then
            lstMenu.AddItem i & ":" & CustomMenu(i).MenuTitle
        End If
    Next
End Sub

Private Sub cmdDelete_Click()
    If lstMenu.ListIndex = -1 Then
        MsgBox "You need to select a menu item to delete", vbInformation
        Exit Sub
    End If
    
    Dim strTemp() As String
    Dim strResponse As String
    strResponse = MsgBox("Are you sure you want to delete this menu item?", vbYesNo, "Delete?")
    If strResponse = vbYes Then
        strTemp = Split(lstMenu.Text, ":")
        CustomMenu(Int(strTemp(0))).MenuLink = ""
        CustomMenu(Int(strTemp(0))).MenuTitle = ""
        Call frmMain.SaveMenus
        Call Me.ReloadMenus
        Call frmMain.ReloadMenuItems
    End If
End Sub

Private Sub cmdEditMenu_Click()
    If lstMenu.ListIndex = -1 Then
        MsgBox "You need to select a menu item to edit.", vbInformation
        Exit Sub
    End If
    If txtEditMenuTitle.Text = "" Then
        MsgBox "You did not enter a menu title.", vbInformation
        Exit Sub
    End If
    If txtEditMenuLink.Text = "http://" Then
        MsgBox "You did not enter a link.", vbInformation
        Exit Sub
    End If
    Dim strTemp() As String
    strTemp = Split(lstMenu.Text, ":")
    CustomMenu(Int(strTemp(0))).MenuLink = Me.txtEditMenuLink.Text
    CustomMenu(Int(strTemp(0))).MenuTitle = Me.txtEditMenuTitle.Text
    Call ReloadMenus
    Call frmMain.SaveMenus
    Call frmMain.ReloadMenuItems
End Sub

Private Sub Form_Load()
    Call ReloadMenus
End Sub

Private Sub lstMenu_Click()
    If lstMenu.Text = "" Then
    
    Else
        Dim strTemp() As String
        strTemp = Split(lstMenu.Text, ":")
        Me.txtEditMenuLink.Text = CustomMenu(Int(strTemp(0))).MenuLink
        Me.txtEditMenuTitle.Text = CustomMenu(Int(strTemp(0))).MenuTitle

    End If
End Sub
