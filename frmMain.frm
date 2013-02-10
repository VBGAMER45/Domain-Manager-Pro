VERSION 5.00
Object = "{79423413-BFDF-483D-BBB2-0D3B88187EB4}#1.0#0"; "WhoIs.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   Caption         =   "Domain Manager Pro"
   ClientHeight    =   6390
   ClientLeft      =   600
   ClientTop       =   1155
   ClientWidth     =   10185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin WhoIsControl.WhoIs WhoIs1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1508
      _ExtentY        =   661
      Server          =   ""
      Query           =   ""
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5040
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Domain Manager"
      TabPicture(0)   =   "frmMain.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSelectAll"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUnSelectAll"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblClearAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstDomains"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frameQuickAdd"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdLoadDomains"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdScanDomains"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExportCSV"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdStop"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Result Scanner"
      TabPicture(1)   =   "frmMain.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "llbPageRank"
      Tab(1).Control(1)=   "lblSecret"
      Tab(1).Control(2)=   "lblAWSKey"
      Tab(1).Control(3)=   "lblAlexa"
      Tab(1).Control(4)=   "lblSignup"
      Tab(1).Control(5)=   "ScriptControl1"
      Tab(1).Control(6)=   "txtBulkResult"
      Tab(1).Control(7)=   "chkPageRank"
      Tab(1).Control(8)=   "chkAlexa"
      Tab(1).Control(9)=   "chkOverture"
      Tab(1).Control(10)=   "chkAvailable"
      Tab(1).Control(11)=   "chkGoogleLinks"
      Tab(1).Control(12)=   "chkYahooLinks"
      Tab(1).Control(13)=   "chkMSNLinks"
      Tab(1).Control(14)=   "cmdScanResults"
      Tab(1).Control(15)=   "lstResults"
      Tab(1).Control(16)=   "cmdExportResults"
      Tab(1).Control(17)=   "cmdClearResults"
      Tab(1).Control(18)=   "txtAlexaSig"
      Tab(1).Control(19)=   "txtAccessKey"
      Tab(1).Control(20)=   "txtCode"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Whois"
      TabPicture(2)   =   "frmMain.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtWhoisInfo"
      Tab(2).Control(1)=   "Whois"
      Tab(2).Control(2)=   "cmdWhois"
      Tab(2).Control(3)=   "txtDomainWhois"
      Tab(2).Control(4)=   "lblWhoisInfo"
      Tab(2).Control(5)=   "lblWhois"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Alexa Information"
      TabPicture(3)   =   "frmMain.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblAlexaDomains"
      Tab(3).Control(1)=   "Label1"
      Tab(3).Control(2)=   "Label2"
      Tab(3).Control(3)=   "Label3"
      Tab(3).Control(4)=   "lblSignUp2"
      Tab(3).Control(5)=   "lblAlexaTotal"
      Tab(3).Control(6)=   "lstAlexa"
      Tab(3).Control(7)=   "txtAlexaDomains"
      Tab(3).Control(8)=   "txtAlexaKey"
      Tab(3).Control(9)=   "txtAlexaSecretKey"
      Tab(3).Control(10)=   "cmdCheckAlexa"
      Tab(3).Control(11)=   "cmdAlexaExportCSV"
      Tab(3).Control(12)=   "cmdAlexaClearList"
      Tab(3).Control(13)=   "tmrAutoSave"
      Tab(3).ControlCount=   14
      Begin VB.Timer tmrAutoSave 
         Interval        =   59000
         Left            =   -70080
         Top             =   3240
      End
      Begin VB.CommandButton cmdAlexaClearList 
         Caption         =   "Clear &List"
         Height          =   375
         Left            =   -68400
         TabIndex        =   49
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAlexaExportCSV 
         Caption         =   "E&xport CSV"
         Height          =   375
         Left            =   -72960
         TabIndex        =   48
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCheckAlexa 
         Caption         =   "Check &Alexa"
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
         Left            =   -74760
         TabIndex        =   47
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtAlexaSecretKey 
         Height          =   285
         Left            =   -73200
         TabIndex        =   44
         ToolTipText     =   "AWS Secret Access Key"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtAlexaKey 
         Height          =   285
         Left            =   -73200
         TabIndex        =   43
         ToolTipText     =   "AWS Access Key"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtAlexaDomains 
         Height          =   1575
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   720
         Width           =   9615
      End
      Begin VB.TextBox txtCode 
         Height          =   1815
         Left            =   -65280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   37
         Text            =   "frmMain.frx":68C2
         Top             =   3240
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox txtAccessKey 
         Height          =   285
         Left            =   -73320
         TabIndex        =   32
         ToolTipText     =   "AWS Access Key"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtAlexaSig 
         Height          =   285
         Left            =   -73320
         TabIndex        =   31
         ToolTipText     =   "AWS Secret Access Key"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.CommandButton cmdClearResults 
         Caption         =   "&Clear Results"
         Height          =   375
         Left            =   -71040
         TabIndex        =   27
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "STOP"
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
         Left            =   8280
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportResults 
         Caption         =   "&Export Results"
         Height          =   375
         Left            =   -72840
         TabIndex        =   25
         Top             =   3600
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstResults 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   24
         Top             =   4080
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PageRank"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Alexa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Overture"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Google Result"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Yahoo Result"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Bing Result"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdScanResults 
         Caption         =   "&Scan Domains"
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CheckBox chkMSNLinks 
         Caption         =   "Bing Links"
         Height          =   255
         Left            =   -69000
         TabIndex        =   22
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkYahooLinks 
         Caption         =   "Yahoo Links"
         Height          =   255
         Left            =   -70320
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkGoogleLinks 
         Caption         =   "Google Links"
         Height          =   255
         Left            =   -71640
         TabIndex        =   20
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox chkAvailable 
         Caption         =   "Available?"
         Height          =   255
         Left            =   -67560
         TabIndex        =   19
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkOverture 
         Caption         =   "Overture"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72720
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkAlexa 
         Caption         =   "Alexa"
         Height          =   255
         Left            =   -73560
         TabIndex        =   17
         Top             =   3240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPageRank 
         Caption         =   "PageRank"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox txtWhoisInfo 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   15
         Top             =   1200
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8070
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":7D7A
      End
      Begin WhoIsControl.WhoIs Whois 
         Left            =   -66600
         Top             =   480
         _ExtentX        =   873
         _ExtentY        =   661
         Server          =   ""
         Query           =   ""
      End
      Begin VB.CommandButton cmdWhois 
         Caption         =   "&Check Whois"
         Height          =   375
         Left            =   -68880
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDomainWhois 
         Height          =   285
         Left            =   -72360
         TabIndex        =   13
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtBulkResult 
         Height          =   1695
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   720
         Width           =   9375
      End
      Begin VB.CommandButton cmdExportCSV 
         Caption         =   "&Export CSV"
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdScanDomains 
         Caption         =   "&Scan Domains"
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
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoadDomains 
         Caption         =   "&Load Text File"
         Height          =   375
         Left            =   5400
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame frameQuickAdd 
         Caption         =   "Quick Domain Add"
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton cmdBulkAddDomain 
            Caption         =   "&Bulk Add"
            Height          =   375
            Left            =   3840
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtDomain 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton cmdAddDomain 
            Caption         =   "&Add"
            Height          =   375
            Left            =   2880
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lstDomains 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8070
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Registrar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Whois Server"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name Server 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Name Server 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Updated Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Created Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Expiration Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   -67800
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         Language        =   "JavaScript"
      End
      Begin MSComctlLib.ListView lstAlexa 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   40
         Top             =   3720
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Alexa Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rank"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Reach"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Full Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Phone Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Email"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DMOZ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Online Since"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "1 Month Reach Per Million"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "1 Month Pageviews per Million"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Language"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblAlexaTotal 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   -71280
         TabIndex        =   50
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label lblSignUp2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Signup Alexa Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   -69480
         TabIndex        =   46
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "<< These two fields require an AWS account you can signup by clicking the link below"
         Height          =   495
         Left            =   -69360
         TabIndex        =   45
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Secret Access Key:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "AWSAccessKeyId:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblAlexaDomains 
         Caption         =   "Enter Domains:"
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblSignup 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Signup Alexa Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -69360
         TabIndex        =   36
         Top             =   2950
         Width           =   3735
      End
      Begin VB.Label lblAlexa 
         BackStyle       =   0  'Transparent
         Caption         =   "<< These two fields require an AWS account you can signup by clicking the link below"
         Height          =   495
         Left            =   -69360
         TabIndex        =   35
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label lblAWSKey 
         Caption         =   "AWSAccessKeyId:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblSecret 
         Caption         =   "Secret Access Key:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblClearAll 
         Alignment       =   2  'Center
         Caption         =   "Clear All Domains"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   8040
         TabIndex        =   30
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblUnSelectAll 
         Alignment       =   2  'Center
         Caption         =   "UnSelect All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblSelectAll 
         Alignment       =   2  'Center
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblWhoisInfo 
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblWhois 
         Caption         =   "Enter domain name here:"
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
         Left            =   -74640
         TabIndex        =   11
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label llbPageRank 
         Caption         =   "Enter Domains:"
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
         Left            =   -74760
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileImport 
         Caption         =   "Import Domains"
         Begin VB.Menu mnuFileLoadDomains 
            Caption         =   "&Load Domains Text File"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuImportSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileLoadCSVData 
            Caption         =   "Load &CSV Data"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuFileImportSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImportRecent 
            Caption         =   "Recent Files"
            Begin VB.Menu mnuMRUFiles 
               Caption         =   "Files:"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu mnuMRUFiles 
               Caption         =   ""
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu mnuMRUFiles 
               Caption         =   ""
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu mnuMRUFiles 
               Caption         =   ""
               Index           =   3
               Visible         =   0   'False
            End
            Begin VB.Menu mnuMRUFiles 
               Caption         =   ""
               Index           =   4
               Visible         =   0   'False
            End
            Begin VB.Menu mnuMRUFiles 
               Caption         =   ""
               Index           =   5
               Visible         =   0   'False
            End
         End
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "Export Domains"
         Begin VB.Menu mnuFileExportDomains 
            Caption         =   "&Export Domains CSV"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuExportSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileExportSelectedCSV 
            Caption         =   "Export Selected Domains CSV"
         End
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print Domains Infomation"
      End
      Begin VB.Menu mnuFileCheckExpire 
         Caption         =   "C&heck For Expiring Domains"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileClearAllDomains 
         Caption         =   "&Clear All Domains"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMinimizeTray 
         Caption         =   "&Minimize To Tray"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuToolsCustomize 
      Caption         =   "&Domains"
      Begin VB.Menu mnuCustomColums 
         Caption         =   "Database Columns"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDomainsRawWhoisData 
         Caption         =   "Raw Whois Data Viewer"
      End
      Begin VB.Menu mnuDomainsImportRawWhois 
         Caption         =   "Import Raw Whois Data for current domains"
      End
      Begin VB.Menu mnuCustomizeMenu 
         Caption         =   "Quick Links"
      End
      Begin VB.Menu mnuCustomizeSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsDomainReport 
         Caption         =   "Generate Domains Report"
      End
      Begin VB.Menu mnuResultsReport 
         Caption         =   "Generate Results Report"
      End
      Begin VB.Menu mnuReportsAlexaReport 
         Caption         =   "Generate Alexa Report"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpUpdates 
         Caption         =   "Check for Updates"
      End
      Begin VB.Menu mnuHelpLicenseKey 
         Caption         =   "Enter License Key"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupViewBrowser 
         Caption         =   "View in Browser"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCopyLine 
         Caption         =   "Copy Whole Line"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "Copy to clipboard Domain"
      End
      Begin VB.Menu mnuPopupRescanDomain 
         Caption         =   "Rescan Domain"
      End
      Begin VB.Menu mnuPopUpEditDomain 
         Caption         =   "Edit Domain Information"
      End
      Begin VB.Menu mnuPopupResult 
         Caption         =   "Result Check"
      End
      Begin VB.Menu mnuPopUpWhois 
         Caption         =   "Check Full Whois"
      End
      Begin VB.Menu mnuPopUpAlexaInformation 
         Caption         =   "Check Alexa Information"
      End
      Begin VB.Menu mnuPopUpCustom 
         Caption         =   "Custom"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Visible         =   0   'False
      Begin VB.Menu mnuSortAZ 
         Caption         =   "Ascending Order"
      End
      Begin VB.Menu mnuSortZA 
         Caption         =   "Decending Order"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Dim ClickedItem As Long
Dim AlexaTotal As Long
Dim prevOrder As Integer
Dim blnStop As Boolean

Dim ThreadInfo(9) As Boolean

Private Sub cmdAlexaClearList_Click()
    Dim Response As String
    Response = MsgBox("Are you sure you want to delete the Alexa results list?", vbYesNo + vbInformation, "Delete Alexa results?")
    If Response = vbYes Then
        lstAlexa.ListItems.Clear
    End If
End Sub

Private Sub cmdAlexaExportCSV_Click()
On Error GoTo errHandle
    If Me.lstAlexa.ListItems.count < 1 Then
        MsgBox "No data on the domains. Use Check Alexa first", vbInformation
        Exit Sub
    End If
    CD.Filename = ""
    CD.DialogTitle = "Save CSV"
    CD.DefaultExt = "csv"
    CD.Filter = "CSV Files (*.csv)|*.csv"
    CD.Flags = cdlOFNOverwritePrompt
    CD.ShowSave
    
    If CD.Filename <> "" Then
            Dim f As Long
            f = FreeFile
            Open CD.Filename For Output As #f
                'Print #f, "Domain, PageRank, Alexa Rank, Google Result"
                Dim g As Long, strData As String
                For g = 1 To Me.lstAlexa.ColumnHeaders.count
                    If g = Me.lstAlexa.ColumnHeaders.count Then
                        strData = strData & Me.lstAlexa.ColumnHeaders.ITem(g).Text
                    Else
                        strData = strData & Me.lstAlexa.ColumnHeaders.ITem(g).Text & ", "
                    End If
                Next
                Print #f, strData
                
                Dim i As Long, data As String
                If Me.lstAlexa.ListItems.count > 0 Then
                    For i = 1 To Me.lstAlexa.ListItems.count
                        data = ""

                            data = Chr$(34) & Me.lstAlexa.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstAlexa.ListItems.ITem(i).ListSubItems.count
                                data = data & "," & Chr$(34) & lstAlexa.ListItems.ITem(i).ListSubItems(g).Text & Chr$(34)
                            Next g
                           Print #f, data
                    
                    Next
                End If

            Close #f
    End If
Exit Sub
errHandle:
    MsgBox "Error: cmdAlexaExportCSV:" & Err.Description
End Sub

Private Sub cmdBulkAddDomain_Click()
    frmBulkAdd.Show
End Sub

Private Sub cmdCheckAlexa_Click()
    'Check if any domains where entered
    If txtAlexaDomains.Text = "" Then
        MsgBox "You need to enter a domain name.", vbInformation
        Exit Sub
    End If
    
        If txtAlexaKey.Text = "" Then
            MsgBox "You need to enter your AWS Access Key in order to get alexa results!", vbInformation
            Exit Sub
        End If
        If txtAlexaSecretKey.Text = "" Then
            MsgBox "You need to enter your Alexa Secert Key in order to get alexa results!", vbInformation
            Exit Sub
        End If
    'Disable the check Alexa button
    cmdCheckAlexa.Enabled = False
    Call CheckAlexa(txtAlexaKey.Text, Me.txtAlexaSecretKey.Text, txtAlexaDomains.Text, False)
    cmdCheckAlexa.Enabled = True
End Sub
Function CheckAlexa(ByVal AlexaID As String, ByVal AlexaSig As String, ByVal strUrl As String, SingleMode As Boolean)
On Error Resume Next
    
    Dim strAlexa As String
    Dim strEmail As String, strPhone As String, strName As String, strReach As String, strDMOZ As String
    Dim strDomains() As String, AlexaQueryCount As Long
    Dim FinalSig As String
    Dim strtimestamp As String, strStatus As String
    Dim strTitle As String, strOnlineSince As String
    Dim strLastSig As String, sUrl As String
    Dim strLanguage As String, strReachOneMonth As String, strPageViews As String
    Dim i As Long

 AlexaQueryCount = 0


    
    'Split the domains in the textbox by a newline
    strDomains = Split(strUrl, vbCrLf)
    

    For i = 0 To UBound(strDomains)
    
        strReach = ""
        strName = ""
        strPhone = ""
        strEmail = ""
        strDMOZ = ""
        strOnlineSince = ""
        strTitle = ""
        strStatus = ""
        strLanguage = ""
        strReachOneMonth = ""
        strPageViews = ""
        'Check if there is a domain
        If Trim$(strDomains(i)) <> "" Then

            strAlexa = "NA"
            'Find the Alexa ranking information
                Dim XMLDOC As MSXML2.DOMDocument
                'Remove http:// from the string if they entered it
                sUrl = "http://" & Replace(strDomains(i), "http://", "")
                Set XMLDOC = New MSXML2.DOMDocument
                XMLDOC.async = False
RedoAlexa:
                'Get the Alexa ranking. Here is the full request

                Dim strSigFull As String, Temp() As String
                strSigFull = "calculateSignature(" & Chr(34) & AlexaSig & Chr(34) & "," & Chr(34) & "AlexaWebInfoService" & Chr(34) & "," & Chr(34) & "UrlInfo" & Chr(34) & "," & Chr(34) & Chr(34) & ")"
                Temp = Split(ScriptControl1.Eval(strSigFull), ":::")
                strtimestamp = Temp(0)
                FinalSig = Temp(1)
                If FinalSig = strLastSig Then
                    
                    'MsgBox "Agag"
                    strLastSig = FinalSig
                    GoTo RedoAlexa
                End If
                
                strLastSig = FinalSig
                
                FinalSig = modEncode.sURLEncode(FinalSig)
              'frmMain.txtAlexaDomains.Text = GetUrl("http://awis.amazonaws.com/?Service=AlexaWebInfoService&Operation=UrlInfo&AWSAccessKeyId=" & AlexaID & "&Signature=" & FinalSig & "&Timestamp=" & strtimestamp & "&Url=" & sUrl & "&ResponseGroup=SiteData,TrafficData,Categories,ContactInfo,Language")
               ' XMLDOC.loadXML (GetUrl("http://awis.amazonaws.com/onca/xml?Service=AlexaWebInfoService&Operation=UrlInfo&AWSAccessKeyId=" & AlexaID & "&Signature=" & FinalSig & "&Timestamp=" & strtimestamp & "&Url=" & sUrl & "&ResponseGroup=SiteData,TrafficData,Categories,ContactInfo,Language"))
               'XMLDOC.loadXML (GetUrl("http://awis.amazonaws.com/?Service=AlexaWebInfoService&Operation=UrlInfo&AWSAccessKeyId=" & AlexaID & "&Signature=" & FinalSig & "&Timestamp=" & strtimestamp & "&Url=" & sUrl & "&ResponseGroup=SiteData,TrafficData,Categories,ContactInfo,Language"))
            XMLDOC.loadXML (GetUrl("http://awis.amazonaws.com/?Service=AlexaWebInfoService&Operation=UrlInfo&AWSAccessKeyId=" & AlexaID & "&Signature=" & FinalSig & "&Timestamp=" & strtimestamp & "&Url=" & sUrl & "&ResponseGroup=SiteData,TrafficData,Categories,ContactInfo,Language"))
              
                 AlexaQueryCount = AlexaQueryCount + 1
                

                Dim oelemlist As IXMLDOMNodeList
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:OperationRequest//aws:Errors//aws:Error//aws:Code")
                'Store the POPULARITY
                strStatus = oelemlist.ITem(0).Text
                If strStatus = "" Then
                    strStatus = "Complete"
                Else

                End If
                
                
                'Scan the xml for the POPULARITY of the domain
                
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:TrafficData//aws:Rank")
                'Store the POPULARITY
                strAlexa = oelemlist.ITem(0).Text
                
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:ContactInfo//aws:OwnerName")

                strName = oelemlist.ITem(0).Text

                
               Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:ContactInfo//aws:PhoneNumbers//aws:PhoneNumber")
                
                strPhone = oelemlist.ITem(0).Text
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:ContactInfo//aws:Email")
                
                strEmail = oelemlist.ITem(0).Text

                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:TrafficData//aws:UsageStatistics//aws:UsageStatistic//aws:Reach//aws:Value")
                
                strReach = oelemlist.ITem(0).Text
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:Related//aws:Categories//aws:CategoryData//aws:Title")
                
                strDMOZ = oelemlist.ITem(0).Text
                
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:ContentData//aws:SiteData//aws:OnlineSince")

                strOnlineSince = oelemlist.ITem(0).Text

                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:ContentData//aws:SiteData//aws:Title")
                strTitle = oelemlist.ITem(0).Text

                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:Language//aws:Locale")
                strLanguage = oelemlist.ITem(0).Text
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:TrafficData//aws:UsageStatistics//aws:UsageStatistic//aws:Reach//aws:PerMillion//aws:Value")
                strReachOneMonth = oelemlist.ITem(1).Text
                
                Set oelemlist = XMLDOC.selectNodes("//aws:UrlInfoResponse//aws:Alexa//aws:TrafficData//aws:UsageStatistics//aws:UsageStatistic//aws:PageViews//aws:PerMillion//aws:Value")
                strPageViews = oelemlist.ITem(1).Text
                  
            
                'Erase XMLDoc
                Set XMLDOC = Nothing
                
                

       
            'Add information to the treeview
            If SingleMode = False Then
                With lstAlexa.ListItems.Add(, , strDomains(i))
                    
                    .ListSubItems.Add , , strStatus
                    .ListSubItems.Add , , strAlexa
                    .ListSubItems.Add , , strReach
                    .ListSubItems.Add , , strName
                    .ListSubItems.Add , , strPhone
                    .ListSubItems.Add , , strEmail
                    .ListSubItems.Add , , strDMOZ
                    .ListSubItems.Add , , strTitle
                    .ListSubItems.Add , , strOnlineSince
                    .ListSubItems.Add , , strReachOneMonth
                    .ListSubItems.Add , , strPageViews
                    .ListSubItems.Add , , strLanguage
                    
                End With
            Else
                CheckAlexa = strAlexa
            
            End If
        End If
        DoEvents
    Next
AlexaTotal = AlexaTotal + AlexaQueryCount
lblAlexaTotal.Caption = "Alexa Total Query Count: " & AlexaTotal


End Function
Private Sub cmdClearResults_Click()
    lstResults.ListItems.Clear
End Sub

Private Sub cmdExportCSV_Click()
    Call ExportDomains
End Sub

Private Sub cmdExportResults_Click()
On Error GoTo errHandle
    If Me.lstResults.ListItems.count = 0 Then
        MsgBox "No Data to export. Run Scan Domains first", vbInformation
        Exit Sub
    End If
    CD.Filename = ""
    CD.DialogTitle = "Save CSV"
    CD.DefaultExt = "csv"
    CD.Filter = "CSV Files (*.csv)|*.csv"
    CD.Flags = cdlOFNOverwritePrompt
    CD.ShowSave
    
    If CD.Filename <> "" Then
        Dim f As Long
        f = FreeFile
        Open CD.Filename For Output As #f
            Print #f, "Domain, PageRank, Alexa, Overture, Google, Yahoo, Msn"
            Dim i As Long
            If Me.lstResults.ListItems.count > 1 Then
                For i = 1 To Me.lstResults.ListItems.count
                    Print #f, Chr$(34) & Me.lstResults.ListItems.ITem(i).Text & Chr$(34) & "," & Chr$(34) & lstResults.ListItems.ITem(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstResults.ListItems.ITem(i).ListSubItems(2).Text & Chr$(34) & "," & Chr$(34) & lstResults.ListItems.ITem(i).ListSubItems(3).Text & Chr$(34) & "," & Chr$(34) & lstResults.ListItems.ITem(i).ListSubItems(4).Text & Chr$(34) & "," & Chr$(34) & lstResults.ListItems.ITem(i).ListSubItems(5).Text & Chr$(34) & "," & Chr$(34) & lstResults.ListItems.ITem(i).ListSubItems(6).Text & Chr$(34)
                Next
            End If
        Close #f
    End If
Exit Sub
errHandle:
    MsgBox "Error: cmdExportCSV:" & Err.Description
End Sub

Private Sub cmdLoadDomains_Click()
    Call LoadDomains
End Sub

Private Sub cmdScanDomains_Click()
    Call DomainScanner(-1)
End Sub
Private Sub DomainScanner(Range As Long)
On Error GoTo errHandle
    Dim i As Long, Min As Long, Max As Long
    Dim strDomainName As String
    cmdStop.Visible = True
    cmdAddDomain.Enabled = False
    cmdBulkAddDomain.Enabled = False
    If Range = -1 Then
        Min = 1
        Max = Me.lstDomains.ListItems.count
    Else
        Min = Range
        Max = Range
    End If
    blnStop = False
    
    For i = Min To Max
        If blnStop = True Then
            Exit Sub
        End If
        strDomainName = lstDomains.ListItems.ITem(i).Text
        
        'Check what ending of domain is and then choose the whois server
        
        Dim c As Long, fDomainWhoisSet As Boolean
        fDomainWhoisSet = False
        For c = 1 To UBound(WhoisServers)
            If Right$(LCase$(strDomainName), Len(WhoisServers(c).Extension)) = WhoisServers(c).Extension Then
                WhoIs1.Server = Trim$(WhoisServers(c).Server)
                fDomainWhoisSet = True
                Exit For
            End If
        Next
        
        If fDomainWhoisSet = False Then
            WhoIs1.Server = Trim$(WhoisServers(1).Server)
        End If
        
    
        'If Right$(LCase$(strDomainName), 4) = ".com" Then
        '    WhoIs1.Server = strWhoisDotCom
        'ElseIf Right$(LCase$(strDomainName), 4) = ".net" Then
         '   WhoIs1.Server = strWhoisDotNet
        'ElseIf Right$(LCase$(strDomainName), 4) = ".org" Then
        '    WhoIs1.Server = strWhoisDotOrg
        'ElseIf Right$(LCase$(strDomainName), 4) = ".biz" Then
        '    WhoIs1.Server = strWhoisDotBiz
        'ElseIf Right$(LCase$(strDomainName), 5) = ".info" Then
        '    WhoIs1.Server = strWhoisDotInfo
        'Else
        '    WhoIs1.Server = strWhoisDotCom
        'End If
        
        WhoIs1.Query = strDomainName
        WhoIs1.Connect
        Dim strData As String
        strData = WhoIs1.Result
        Dim J As Long
        Dim strTemp() As String
        If WhoIs1.Server = strWhoisDotCom Or WhoIs1.Server = strWhoisDotNet Then
            strTemp = Split(strData, vbLf)
        Else
            strTemp = Split(strData, vbCrLf)
        End If
      
        Dim NS1 As Boolean, TwoNS As Boolean
        TwoNS = False
        NS1 = False
        
        ' Check if we are getting the IP Address
        If blnGetIPAddress = True Then
         If SocketsInitialize() Then
            'obtain and pass the host address to the function
            
            lstDomains.ListItems.ITem(i).SubItems(GetColByName("IP Address")) = GetIPFromHostName(strDomainName)
           
            SocketsCleanup
           
         End If
        End If
        
        'MsgBox UBound(strTemp)
        For J = 0 To UBound(strTemp)
            With lstDomains.ListItems.ITem(i)
                If InStr(1, strTemp(J), "Registrar:") <> 0 Then
                    .SubItems(GetColByName("Registrar")) = Trim$(Replace(strTemp(J), "Registrar:", ""))

                ElseIf InStr(1, strTemp(J), "Whois Server:") <> 0 Then
                    .SubItems(GetColByName("Whois Server")) = Trim$(Replace(strTemp(J), "Whois Server:", ""))
                ElseIf InStr(1, strTemp(J), "Name Server:") <> 0 And TwoNS = False Then
    
                    If NS1 = False Then
                       .SubItems(GetColByName("Name Server 1")) = Trim$(Replace(strTemp(J), "Name Server:", ""))
                        NS1 = True
                    Else
                        .SubItems(GetColByName("Name Server 2")) = Trim$(Replace(strTemp(J), "Name Server:", ""))
                        TwoNS = True
                    End If
                ElseIf InStr(1, strTemp(J), "Status:") <> 0 Then
                    .SubItems(GetColByName("Status")) = Trim$(Replace(strTemp(J), "Status:", ""))
                ElseIf InStr(1, strTemp(J), "Updated Date:") <> 0 Then
                    .SubItems(GetColByName("Updated Date")) = Trim$(Replace(strTemp(J), "Updated Date:", ""))
                ElseIf InStr(1, strTemp(J), "Last Updated On:") <> 0 Then
                    .SubItems(GetColByName("Updated Date")) = Trim$(Replace(strTemp(J), "Last Updated On:", ""))
                ElseIf InStr(1, strTemp(J), "Created On:") <> 0 Then
                    .SubItems(GetColByName("Created Date")) = Trim$(Replace(strTemp(J), "Created On:", ""))
                ElseIf InStr(1, strTemp(J), "Creation Date:") <> 0 Then
                    .SubItems(GetColByName("Created Date")) = Trim$(Replace(strTemp(J), "Creation Date:", ""))
                ElseIf InStr(1, strTemp(J), "Expiration Date:") <> 0 Then
                    .SubItems(GetColByName("Expiration Date")) = Trim$(Replace(strTemp(J), "Expiration Date:", ""))
                End If
            End With
            
            
            DoEvents
        Next
        If blnRawWhois = True Then
            Dim f As Long
            f = FreeFile
            Open App.Path & "\rawwhois\" & strDomainName & ".txt" For Output As #f
                Print #f, strData
            Close #f
        End If
        
        DoEvents
    Next
    cmdStop.Visible = False
    cmdAddDomain.Enabled = True
    cmdBulkAddDomain.Enabled = True
    'Call CheckDomainExpire
Exit Sub
errHandle:
    MsgBox "Error: DomainScanner " & Err.Number & " " & Err.Description
End Sub
Private Sub cmdScanResults_Click()
On Error Resume Next
    Dim strDomains() As String, sUrl As String, strTmp() As String
    Dim strRes As String, strAlexa As String, strOverture As String
    Dim Rank As Long, i As Long, chSum
    Dim strYahoo As String, strMSN As String, strGoogle As String

    If txtBulkResult.Text = "" Then
        MsgBox "You did not enter any domains to scan.", vbInformation
        Exit Sub
    End If
    If chkAlexa.Value = vbChecked Then
        If txtAccessKey.Text = "" Then
            MsgBox "You need to enter your AWS Access Key in order to get alexa results!", vbInformation
            chkAlexa.Value = vbUnchecked
            Exit Sub
        End If
        If txtAlexaSig.Text = "" Then
            MsgBox "You need to enter your Alexa Secret Key in order to get alexa results!", vbInformation
            chkAlexa.Value = vbUnchecked
            Exit Sub
        End If
    End If

    strDomains = Split(txtBulkResult.Text, vbCrLf)


    For i = 0 To UBound(strDomains)
        'Check if there is a domain
        If Trim$(strDomains(i)) <> "" Then
            'PageRank
            If chkPageRank.Value = vbChecked Then
                'Remove http:// from the string if they entered it
                sUrl = "http://" & Replace(strDomains(i), "http://", "")
                'Calculate Google Checksum
                chSum = CalculateChecksum(sUrl)
                'Read from Google
                strRes = GetUrl("http://www.google.com/search?client=navclient-auto&ch=" & chSum & "&features=Rank&q=info:" & sUrl)
                strTmp = Split(strRes, ":")
                Rank = CInt(strTmp(2))

 
                
            End If
            strAlexa = "NA"
            'Find the Alexa ranking information
            If chkAlexa.Value = vbChecked Then
                'Dim XMLDOC As MSXML2.DOMDocument
                'Remove http:// from the string if they entered it
                'sUrl = "http://" & Replace(strDomains(i), "http://", "")
                'Set XMLDOC = New MSXML2.DOMDocument
                'XMLDOC.async = False
                'Get the Alexa ranking. Here is the full request
                'XMLDOC.loadXML (GetUrl("http://209.237.237.101/data/rdys313tqmO3uv?cli=10&dat=snba&ver=7.0&cdt=alx_vw%3D20%26wid%3D22032%26act%3D00000000000%26ss%3D800x600%26bw%3D792%26t%3D0%26ttl%3D520%26vis%3D1%26rq%3D21&url=" & sUrl))
                'Dim oelemlist As IXMLDOMNodeList
                'Scan the xml for the POPULARITY of the domain
                'Set oelemlist = XMLDOC.selectNodes("//ALEXA//SD//POPULARITY")
                'Store the POPULARITY
                'strAlexa = oelemlist.Item(0).Attributes.Item(1).Text
                'Erase XMLDoc
                'Set XMLDOC = Nothing
                strAlexa = CheckAlexa(Me.txtAccessKey.Text, txtAlexaSig.Text, strDomains(i), True)
                
            End If
            strOverture = "NA"
            If chkOverture.Value = vbChecked Then

            End If
            strGoogle = "NA"
            If chkGoogleLinks.Value = vbChecked Then
                Dim strGoogleData As String
                strGoogleData = GetUrl("http://www.google.com/search?hl=en&lr=&q=link%3A" & Replace(strDomains(i), "http://", ""))
                If InStr(1, strGoogleData, "- did not match any documents. ") <> 0 Then
                    strGoogle = "0"
                Else
                    Dim gNum As Long, gI As Long
                    gNum = InStr(1, strGoogleData, "</b> of about <b>")
                    strGoogle = ""
                    gI = 17
                    Do
                        strGoogle = strGoogle & Mid$(strGoogleData, gNum + gI, 1)
                        gI = gI + 1
                    Loop Until Mid$(strGoogleData, gNum + gI, 1) = "<"
                    
                End If
            End If
            strYahoo = "NA"
            If chkYahooLinks.Value = vbChecked Then
            
               Dim XMLDOC2 As MSXML2.DOMDocument
                 sUrl = "http://www." & Replace(strDomains(i), "http://", "")
               'Set XMLDOC2 = New MSXML2.DOMDocument
               XMLDOC2.async = False
                'Get the Alexa ranking. Here is the full request
               
                
                XMLDOC2.loadXML (GetUrl("http://api.search.yahoo.com/SiteExplorerService/V1/inlinkData?appid=AZAZ1975&query=" & sUrl))
                Dim oelemlist2 As IXMLDOMNodeList
                'Scan the xml for the POPULARITY of the domain
                Set oelemlist2 = XMLDOC2.selectNodes("//ResultSet")
                
                'Store the POPULARITY
                strYahoo = oelemlist2.ITem(1).Attributes.ITem(3).Text
               ' MsgBox oelemlist2.ITem(2).Attributes.ITem(0).Text
                
                'Erase XMLDoc
                Set XMLDOC2 = Nothing
                
                'Dim strYahooData As String
                'strYahooData = GetUrl("http://siteexplorer.search.yahoo.com/search?p=link%3Awww." & Replace(strDomains(i), "http://", "") & "&bwm=i&bwmf=u")
               '  MsgBox strYahooData
               '  If InStr(1, strYahooData, "<span class=" & Chr(34) & "btn" & Chr(34) & ">Inlinks (") <> 0 Then
                '    strYahoo = "0"
               ' Else
                '    Dim YahooNum As Long, YahooI As Long
                '    YahooNum = InStr(1, strYahooData, "<span class=" & Chr(34) & "btn" & Chr(34) & ">Inlinks (")
                    
                 '   If strYahooData <> "" Then
                  '  Do
                           
                    '    strYahoo = strYahoo & Mid$(strYahooData, YahooNum + YahooI, 1)
                    '        YahooI = YahooI + 1
                    '    Loop Until Mid$(strYahooData, YahooNum + YahooI, 1) = ")"
                   ' End If
              '  End If
                
                
            End If
            strMSN = "NA"
            If chkMSNLinks.Value = vbChecked Then
                Dim strMSNData As String
                strMSNData = GetUrl("http://search.msn.com/results.aspx?q=link%3Awww." & Replace(strDomains(i), "http://", ""))
                'Debug.Print strMSNData
           
                If InStr(1, strMSNData, "We couldn't find any results containing ") <> 0 Then
                    strMSN = "0"
                Else
                    Dim msnNum As Long, mI As Long
                    msnNum = InStr(1, strMSNData, "<h1>Web Results</h1><h5>Page 1 of ")
                    strMSN = ""
                    mI = 34
                    If strMSNData <> "" Then
                        Do
                           
                            strMSN = strMSN & Mid$(strMSNData, msnNum + mI, 1)
                            mI = mI + 1
                        Loop Until Mid$(strMSNData, msnNum + mI, 1) = " "
                    End If
                End If
            End If
            
            With lstResults.ListItems.Add(, , strDomains(i))
                
                .ListSubItems.Add , , Rank
                .ListSubItems.Add , , strAlexa
                .ListSubItems.Add , , strOverture
                .ListSubItems.Add , , strGoogle
                .ListSubItems.Add , , strYahoo
                .ListSubItems.Add , , strMSN
                
            End With
            DoEvents
        End If
    Next
End Sub

Private Sub cmdStop_Click()
    blnStop = True
    cmdStop.Visible = False
    cmdAddDomain.Enabled = True
    cmdBulkAddDomain.Enabled = True
End Sub

Private Sub cmdWhois_Click()
    If txtDomainWhois.Text <> "" Then
        txtWhoisInfo.Text = ""
        
        
        'Check the extension of the domain and choose whois server
     
        'If Right$(LCase$(txtDomainWhois.Text), 4) = ".com" Then
        '    Whois.Server = strWhoisDotCom
        'ElseIf Right$(LCase$(txtDomainWhois.Text), 4) = ".net" Then
        '    Whois.Server = strWhoisDotNet
        'ElseIf Right$(LCase$(txtDomainWhois.Text), 4) = ".org" Then
        '    Whois.Server = strWhoisDotOrg
        'ElseIf Right$(LCase$(txtDomainWhois.Text), 4) = ".biz" Then
        '    Whois.Server = strWhoisDotBiz
        'ElseIf Right$(LCase$(txtDomainWhois.Text), 5) = ".info" Then
        '    Whois.Server = strWhoisDotInfo
        'Else
        '    Whois.Server = strWhoisDotCom
        'End If
        Dim c As Long, fDomainWhoisSet As Boolean
        fDomainWhoisSet = False
        For c = 1 To UBound(WhoisServers)
            If Right$(LCase$(txtDomainWhois.Text), Len(WhoisServers(c).Extension)) = WhoisServers(c).Extension Then
                Whois.Server = Trim$(WhoisServers(c).Server)
                fDomainWhoisSet = True
                Exit For
            End If
        Next
        
        If fDomainWhoisSet = False Then
            Whois.Server = Trim$(WhoisServers(1).Server)
        End If
        'Debug.Print "#" & WhoIs1.Server & "#"
        '/Debug.Print "#" & strWhoisDotCom & "#"
        ' Whois.Server = strWhoisDotCom
        'MsgBox WhoIs1.Server
        
        Whois.Query = Trim$(txtDomainWhois.Text)
        Whois.Connect
  
        txtWhoisInfo.Text = Whois.Result
    Else
        MsgBox "You did not enter a domain name!", vbInformation
    End If
End Sub



Private Sub cmdAddDomain_Click()
    If txtDomain.Text = "" Then
        MsgBox "You need to enter a domain name.", vbInformation
        Exit Sub
    Else
        Dim i As Long
        'Check if domain is already in list
        For i = 1 To Me.lstDomains.ListItems.count
            If lstDomains.ListItems.ITem(i).Text = txtDomain.Text Then
                MsgBox "You already have that domain listed.", vbInformation
                Exit Sub
            End If
        Next
        'Sanity checks
        If InStr(Me.txtDomain.Text, ".") = False Then
            MsgBox "You have entered an invalid domain name", vbInformation
            Exit Sub
        End If
        txtDomain.Text = Replace(txtDomain.Text, "http://www", "")
        txtDomain.Text = Replace(txtDomain.Text, "http://", "")
        
        
        lstDomains.ListItems.Add , , txtDomain.Text
        
    End If
End Sub



Private Sub Form_Load()
On Error GoTo errHandle:
    Me.Tag = "Copyright VisualBasicZone.com 2006"
    blnStartUp = GetSetting("Domain Manager Pro", "Options", "StartUp", "True")
    If blnStartUp = True Then
        Call modMain.RegRun(App.Path & "\DNManagerPro.exe tray", "DomainManagerPro")
    Else
        Call modMain.RemoveRegRun("DomainManagerPro")

    End If
    
    blnRawWhois = GetSetting("Domain Manager Pro", "Options", "RawWhois", "True")
    
    Call MruList.Load
    Call MruList.Update(frmMain)
    
    Dim ipSetting As String
    ipSetting = GetSetting("Domain Manager Pro", "Options", "GetIP", "0")
    If ipSetting = "1" Then
        blnGetIPAddress = True
    Else
        blnGetIPAddress = False
    End If
    
    'Setup Whois servers
    strWhoisDotCom = DotCom
    strWhoisDotNet = DotNet
    strWhoisDotOrg = DotOrg
    strWhoisDotBiz = DotBiz
    strWhoisDotInfo = DotInfo
    
    strWhoisDotCom = GetSetting("Domain Manager Pro", "Options", "DotCom", strWhoisDotCom)
    strWhoisDotNet = GetSetting("Domain Manager Pro", "Options", "DotNet", strWhoisDotNet)
    strWhoisDotOrg = GetSetting("Domain Manager Pro", "Options", "DotOrg", strWhoisDotOrg)
    strWhoisDotBiz = GetSetting("Domain Manager Pro", "Options", "DotBiz", strWhoisDotBiz)
    strWhoisDotInfo = GetSetting("Domain Manager Pro", "Options", "DotInfo", strWhoisDotInfo)
    
    'Alexa Key
    txtAccessKey.Text = GetSetting("Domain Manager Pro", "Options", "AlexaID", "")
    txtAlexaSig.Text = GetSetting("Domain Manager Pro", "Options", "AlexaSig", "")
    Me.txtAlexaKey.Text = txtAccessKey.Text
    Me.txtAlexaSecretKey.Text = txtAlexaSig.Text
    
    ExpireDays = GetSetting("Domain Manager Pro", "Options", "ExpireCheck", "30")
    Dim RegTest As String
         RegTest = GetSetting("Domain Manager Pro", "Options", "Yup", "#")
    If RegTest = "!" Then
        frmMain.mnuHelpLicenseKey.Visible = False
        
    Else
        frmLicense.Show vbModal, Me
        frmLicense.cmdClose.Visible = False
        frmMain.mnuHelpLicenseKey.Visible = False
    End If

    Call LoadCol
    
    Call LoadDomainInfo
    
    Call LoadMenus
    
    Call modMain.LoadWhoisData
    
    ScriptControl1.AddCode txtCode.Text
    
    frmTrayIcon.Show
    frmTrayIcon.Hide
    If Command <> "Tray" Then
        frmMain.Show
    Else
        frmMain.Hide
    End If
    
    Call MakeFolder(App.Path & "\rawwhois\")
Exit Sub
errHandle:
    MsgBox "Error: frmMain_Load() " & Err.Number & " " & Err.Description
End Sub
Private Sub MakeFolder(Path As String)
On Error Resume Next
    MkDir Path
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveDomains
    Cancel = True
    Me.Hide
    IsInTray = True
    frmTrayIcon.mnuHideDomainManager.Caption = "Show Domain Manager Pro"
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    SSTab.Width = Me.Width - 300
    SSTab.Height = Me.Height - 700
    lstDomains.Width = SSTab.Width - 200
    lstDomains.Height = SSTab.Height - 1300
    txtWhoisInfo.Width = SSTab.Width - 300
    txtWhoisInfo.Height = SSTab.Height - 1300
    txtBulkResult.Width = SSTab.Width - 300
    txtAlexaDomains.Width = SSTab.Width - 300
    lstAlexa.Width = SSTab.Width - 300
    lstAlexa.Height = SSTab.Height - 4000
    lstResults.Width = SSTab.Width - 300
    lstResults.Height = SSTab.Height - 4200
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Call SaveDomains
End Sub



Private Sub lblClearAll_Click()
    Dim Response As String
    Response = MsgBox("Are you sure you want to clear all domains?", vbYesNo + vbExclamation, "Clear Domains?")
    If Response = vbYes Then
        lstDomains.ListItems.Clear
    End If
End Sub

Private Sub lblSelectAll_Click()
    Dim i As Long
    For i = 1 To Me.lstDomains.ListItems.count
        Me.lstDomains.ListItems(i).Selected = True
    Next
End Sub



Private Sub lblSignup_Click()
On Error Resume Next
 ShellExecute Me.hWnd, vbNullString, "http://www.amazon.com/gp/browse.html/ref=sc_fe_c_0_15763381_6/104-8609174-5267946?%5Fencoding=UTF8&node=12782661&no=15763381&me=A36L942TSJ2AJA", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub lblSignUp2_Click()
On Error Resume Next
 ShellExecute Me.hWnd, vbNullString, "http://www.amazon.com/gp/browse.html/ref=sc_fe_c_0_15763381_6/104-8609174-5267946?%5Fencoding=UTF8&node=12782661&no=15763381&me=A36L942TSJ2AJA", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub lblUnSelectAll_Click()
    Dim i As Long
    For i = 1 To Me.lstDomains.ListItems.count
        Me.lstDomains.ListItems(i).Selected = False
    Next
End Sub

Private Sub lstDomains_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo errHandle
    PopupMenu mnuSort
    
    Dim currSortKey As Integer
    
    lstDomains.SortKey = ColumnHeader.index - 1
    currSortKey = lstDomains.SortKey
    
    lstDomains.SortOrder = Abs(Not lstDomains.SortOrder = 1)
    lstDomains.Sorted = True
    
    
    mnuSortAZ.Checked = lstDomains.SortOrder = 0
    mnuSortZA.Checked = mnuSortAZ.Checked = False
  
    If currSortKey > -1 Then
      prevOrder% = currSortKey
    End If
Exit Sub
errHandle:
    MsgBox "Error_lstDomains_ColumnClick: " & Err.Number & " " & Err.Description
End Sub

Private Sub lstDomains_DblClick()
    If Me.lstDomains.ListItems.count = 0 Then Exit Sub
    frmEditDomain.Show
    frmEditDomain.EditDomain (ClickedItem)
    
End Sub

Private Sub lstDomains_ItemClick(ByVal ITem As MSComctlLib.ListItem)
   ClickedItem = ITem.index
    
End Sub

Private Sub lstDomains_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstDomains.ListItems.count = 0 Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu Me.mnuPopupMenu
    End If

End Sub

Private Sub mnuCustomizeMenu_Click()
    frmMenu.Show vbModal, Me
End Sub

Private Sub mnuDomainsImportRawWhois_Click()
    Call DomainScannerCache(-1)
    MsgBox "Raw Whois Data Imported", vbInformation
End Sub

Private Sub mnuDomainsRawWhoisData_Click()
    frmRawWhois.Show
End Sub

Private Sub mnuFileCheckExpire_Click()
    If Me.lstDomains.ListItems.count = 0 Then
        MsgBox "No Domains are loaded.", vbInformation
        Exit Sub
    End If
    Call CheckDomainExpire
    
End Sub

Private Sub mnuFileClearAllDomains_Click()
Dim Response As String
    Response = MsgBox("Are you sure you want to clear all domains?", vbYesNo + vbExclamation, "Clear Domains?")
    If Response = vbYes Then
        lstDomains.ListItems.Clear
    End If
End Sub

Private Sub mnuFileExit_Click()
    If txtAccessKey.Text <> "" Then
    SaveSetting "Domain Manager Pro", "Options", "AlexaID", txtAccessKey.Text
    Else
        SaveSetting "Domain Manager Pro", "Options", "AlexaID", Me.txtAlexaKey.Text
    End If
    If txtAlexaSig.Text <> "" Then
    SaveSetting "Domain Manager Pro", "Options", "AlexaSig", txtAlexaSig.Text
    Else
        SaveSetting "Domain Manager Pro", "Options", "AlexaSig", Me.txtAlexaSecretKey.Text
    End If
    Call MruList.Save
    Call SaveDomains
    MySysTray.HideIcon
    
    End
End Sub

Private Sub mnuFileExportDomains_Click()
    Call ExportDomains
End Sub

Private Sub mnuFileExportSelectedCSV_Click()
On Error Resume Next
    If Me.lstDomains.ListItems.count = 0 Then
        MsgBox "No Data to export. Run Scan Domains first", vbInformation
        Exit Sub
    End If
    CD.Filename = ""
    CD.DialogTitle = "Save CSV"
    CD.DefaultExt = "csv"
    CD.Filter = "CSV Files (*.csv)|*.csv"
    CD.Flags = cdlOFNOverwritePrompt
    CD.ShowSave
    
    If CD.Filename <> "" Then
            Dim f As Long
            f = FreeFile
   
            Open CD.Filename For Output As #f
            
                'Print #f, "Domain, Register, Whois Server, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
                'Dim i As Long
                'If Me.lstDomains.ListItems.Count > 0 Then
                '    For i = 1 To Me.lstDomains.ListItems.Count
                '        Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(2).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(3).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(4).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(5).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(6).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(7).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(8).Text & Chr$(34)
                '    Next
                'End If
                Dim g As Long, strData As String
                For g = 1 To Me.lstDomains.ColumnHeaders.count
                    If g = Me.lstDomains.ColumnHeaders.count Then
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text
                    Else
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text & ", "
                    End If
                Next
                Print #f, strData
 
                Dim i As Long, data As String
                If Me.lstDomains.ListItems.count > 0 Then
                    For i = 1 To Me.lstDomains.ListItems.count
                        data = ""
                        If lstDomains.ListItems(i).Selected = True Then
                            If Me.lstDomains.ListItems.ITem(i).ListSubItems.count <> 0 Then
                                data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                                For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                                    data = data & "," & Chr$(34) & lstDomains.ListItems.ITem(i).ListSubItems(g).Text & Chr$(34)
                                Next g
                               Print #f, data
                            Else
                                data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                                For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                                    data = data & "," & Chr$(34) & Chr$(34)
                                Next g
                               Print #f, data
                            End If
                        End If
                    
                    Next
                End If
                

            Close #f
    End If
Exit Sub
errHandle:
    MsgBox "Error: mnuFileExportSelectedCSV:" & Err.Description
End Sub
Sub LoadCSVDomains(Filename As String)
            Dim f As Long
            f = FreeFile
            Dim strData As String
             lstDomains.Sorted = False
            Open Filename For Input As #f
                Line Input #f, strData
                
                Dim strCol() As String
                Dim g As Long
                For g = 1 To lstDomains.ColumnHeaders.count
                    ReDim strCol(lstDomains.ColumnHeaders.count - 1)
                Next
                'Print #f, "Domain, Register, Whois Server, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
                Dim i As Long
                Dim z As Long
                Dim k As Long
                 'MsgBox EOF(f)
                Do Until EOF(f)
                   
                    For k = 0 To UBound(strCol)
                        Input #f, strCol(k)
                    Next

               
                
    
                       With lstDomains.ListItems.Add(, , strCol(0))

                        For z = 1 To UBound(strCol)
                           .SubItems(z) = strCol(z)
                        Next
                     End With
  
                Loop
                 'Check if domain is already in list
                

            Close #f
            Call MruList.Add(Filename)
            Call MruList.Update(frmMain)
            lstDomains.Sorted = True
    


Exit Sub
errHandle:
Exit Sub
    MsgBox "Error: LoadCSVDomains:" & Err.Description

End Sub
Private Sub mnuFileLoadCSVData_Click()
'On Error GoTo errHandle
    CD.Filename = ""
    CD.DialogTitle = "Open CSV"
    CD.DefaultExt = "csv"
    CD.Filter = "CSV Files (*.csv)|*.csv"
    CD.Flags = cdlOFNOverwritePrompt
    CD.ShowOpen
    
    If CD.Filename <> "" Then
        Call LoadCSVDomains(CD.Filename)
    End If
Exit Sub
'On Error GoTo errHandle
'    CD.FileName = ""
'    CD.DialogTitle = "Open CSV"
'    CD.DefaultExt = "csv"
'    CD.Filter = "CSV Files (*.csv)|*.csv"
'    CD.Flags = cdlOFNOverwritePrompt
'    CD.ShowOpen
'
'    If CD.FileName <> "" Then
'           ' Dim f As Long, strData As String
'            f = FreeFile
'            Open CD.FileName For Input As #f
'                Line Input #f, strData
'                Dim d1 As String, d2 As String, d3 As String
'                Dim d4 As String, d5 As String, d6 As String
'                Dim d7 As String, d8 As String, d9 As String
'                Do While Not EOF(f)
'                    Input #f, d1, d2, d3, d4, d5, d6, d7, d8, d9
'
'                    With lstDomains
'                        .ListItems.Add , , d1
'                        .ListItems.Item(.ListItems.Count).SubItems(1) = d2
'                        .ListItems.Item(.ListItems.Count).SubItems(2) = d3
'                        .ListItems.Item(.ListItems.Count).SubItems(3) = d4
'                        .ListItems.Item(.ListItems.Count).SubItems(4) = d5
'                        .ListItems.Item(.ListItems.Count).SubItems(5) = d6
'                        .ListItems.Item(.ListItems.Count).SubItems(6) = d7
'                        .ListItems.Item(.ListItems.Count).SubItems(7) = d8
'                        .ListItems.Item(.ListItems.Count).SubItems(8) = d9
'                    End With
'
'                Loop
'
'
'            Close #f
'    End If
'Exit Sub
''errHandle:
 '   MsgBox "Error: mnuFileLoadCSVData" & Err.Description
End Sub

Private Sub mnuFileLoadDomains_Click()
    Call LoadDomains
End Sub

Private Sub mnuFileMinimizeTray_Click()
    Me.Hide
    IsInTray = True
    frmTrayIcon.mnuHideDomainManager.Caption = "Show Domain Manager Pro"
    
End Sub

Private Sub mnuFilePrint_Click()
On Error Resume Next
    CD.CancelError = True
    CD.ShowPrinter
    
    If Err.Number = cdlCancel Then
        CD.CancelError = False
            Exit Sub
    End If

    Printer.FontSize = 26
    Printer.FontBold = True
    Printer.Print "Domain Manager Pro - Domain Listing"
    Printer.FontBold = False
    Printer.FontSize = 6
    'Printer.Print "Domain, Register, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
    Dim g As Long, strData As String
    For g = 1 To Me.lstDomains.ColumnHeaders.count
        If g = Me.lstDomains.ColumnHeaders.count Then
            strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text
        Else
            strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text & ", "
        End If
    Next
    Printer.Print strData
                
    
    Dim i As Long
    'If Me.lstDomains.ListItems.Count > 0 Then
    '    For i = 1 To Me.lstDomains.ListItems.Count
    '        Printer.Print Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(3).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(4).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(5).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(6).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(7).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(8).Text & Chr$(34)
    '    Next
    'End If
    Dim data As String
    If Me.lstDomains.ListItems.count > 0 Then
        For i = 1 To Me.lstDomains.ListItems.count
            data = ""
            If Me.lstDomains.ListItems.ITem(i).ListSubItems.count <> 0 Then
               data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
               For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                   data = data & "," & Chr$(34) & lstDomains.ListItems.ITem(i).ListSubItems(g).Text & Chr$(34)
               Next g
                   Printer.Print data
            Else
                data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                    data = data & "," & Chr$(34) & Chr$(34)
                Next g
                Printer.Print data
            End If
                    
        Next
    End If
    
    Printer.EndDoc
    
    CD.CancelError = False

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpLicenseKey_Click()
    frmLicense.Show
End Sub

Private Sub mnuHelpUpdates_Click()
On Error GoTo errHandle
    Dim data As String, ver As String
    data = GetUrl("http://www.dnmanagerpro.com/update.txt")
    ver = App.Major & "." & App.Minor & "." & App.Revision

    If data = ver Then
        MsgBox "Your version is up to date!", vbInformation
    Else
        MsgBox "There is a newer version out! Your Version: " & ver & " Latest Version: " & data, vbCritical
    End If
    
Exit Sub
errHandle:
MsgBox "Error mnuHelpUpdates: " & Err.Description
End Sub
Function GetUrl(Url As String) As String
On Error GoTo errHandle:
    Dim xmlhttp As Object
    
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Indicate that page that will receive the request and the
    ' type of request being submitted
    xmlhttp.open "GET", Url, False
    
   'You need to send the request
    xmlhttp.send
    'Return the result
    GetUrl = xmlhttp.responseText

    Set xmlhttp = Nothing
Exit Function
errHandle:
    MsgBox "Error: GetUrl: - " & Err.Description
End Function

Private Sub mnuMRUFiles_Click(index As Integer)
On Error GoTo errHandle
    If FileExists(mnuMRUFiles(index).Caption) Then
        If Right$(LCase$(mnuMRUFiles(index).Caption), 4) = ".txt" Then
            Call Me.LoadDomainsByFile(mnuMRUFiles(index).Caption)
        End If
        If Right$(LCase$(mnuMRUFiles(index).Caption), 4) = ".csv" Then
            Call LoadCSVDomains(mnuMRUFiles(index).Caption)
        End If

    Else
        MsgBox "The file does not exist!", vbExclamation
    
    End If
Exit Sub
errHandle:
    MsgBox "Error mnuMRUFiles: " & Err.Number & " " & Err.Description
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
    
End Sub

Private Sub mnuPopUpAlexaInformation_Click()
On Error GoTo errHandle
    Dim i As Long
    Dim strData As String

    For i = 1 To Me.lstDomains.ListItems.count
        If lstDomains.ListItems.ITem(i).Selected = True Then
            strData = strData & lstDomains.ListItems.ITem(i).Text & vbNewLine
        End If
    Next
    txtAlexaDomains.Text = strData
    SSTab.Tab = 3
Exit Sub
errHandle:
    MsgBox "Error: mnuPopUpAlexaInformation " & Err.Number & " " & Err.Description

End Sub

Private Sub mnuPopupCopy_Click()
    Dim i As Long
    Dim strData As String

    For i = 1 To Me.lstDomains.ListItems.count
        If lstDomains.ListItems.ITem(i).Selected = True Then
            strData = strData & lstDomains.ListItems.ITem(i).Text & vbNewLine
        End If
    Next
    Clipboard.Clear
    Clipboard.SetText strData


End Sub

Private Sub mnuPopupCopyLine_Click()
    Dim i As Long
    Dim strData As String

    For i = 1 To Me.lstDomains.ListItems.count
        If lstDomains.ListItems.ITem(i).Selected = True Then
            strData = strData & lstDomains.ListItems.ITem(i).Text & "," & lstDomains.ListItems.ITem(i).SubItems(1) & "," & lstDomains.ListItems.ITem(i).SubItems(2) & "," & lstDomains.ListItems.ITem(i).SubItems(3) & "," & lstDomains.ListItems.ITem(i).SubItems(4) & "," & lstDomains.ListItems.ITem(i).SubItems(5) & "," & lstDomains.ListItems.ITem(i).SubItems(6) & "," & lstDomains.ListItems.ITem(i).SubItems(7) & "," & lstDomains.ListItems.ITem(i).SubItems(8) & vbNewLine
        End If
    Next
    Clipboard.Clear
    Clipboard.SetText strData
End Sub

Private Sub mnuPopUpCustom_Click(index As Integer)
On Error Resume Next
    Dim i As Long
    For i = 0 To UBound(CustomMenu)
        If CustomMenu(i).MenuTitle = mnuPopUpCustom(index).Caption Then
            ShellExecute Me.hWnd, vbNullString, CustomMenu(i).MenuLink, vbNullString, "C:\", SW_SHOWNORMAL
        End If
    Next
End Sub

Private Sub mnuPopupDelete_Click()
    Dim i As Long

    For i = Me.lstDomains.ListItems.count To 1 Step -1
        If lstDomains.ListItems.ITem(i).Selected = True Then
           lstDomains.ListItems.Remove (i)
        End If
    Next

End Sub

Private Sub mnuPopUpEditDomain_Click()
    If Me.lstDomains.ListItems.count = 0 Then Exit Sub
    frmEditDomain.Show
    frmEditDomain.EditDomain (ClickedItem)
End Sub

Private Sub mnuPopupRescanDomain_Click()
On Error GoTo errHandle
    Dim i As Long

    For i = 1 To Me.lstDomains.ListItems.count
        If lstDomains.ListItems.ITem(i).Selected = True Then
                Call DomainScanner(i)
        End If
    Next
Exit Sub
errHandle:
    MsgBox "Error mnuPopUpRescanDomain: " & Err.Description
End Sub

Private Sub mnuPopupResult_Click()
On Error GoTo errHandle
    Dim i As Long
    Dim strData As String

    For i = 1 To Me.lstDomains.ListItems.count
        If lstDomains.ListItems.ITem(i).Selected = True Then
            strData = strData & lstDomains.ListItems.ITem(i).Text & vbNewLine
        End If
    Next
    txtBulkResult.Text = strData
    SSTab.Tab = 1
Exit Sub
errHandle:
    MsgBox "Error: mnuPopUpResult " & Err.Number & " " & Err.Description
End Sub

Private Sub mnuPopupViewBrowser_Click()
On Error GoTo errHandle
    Dim i As Long

    For i = 1 To Me.lstDomains.ListItems.count
        If lstDomains.ListItems.ITem(i).Selected = True Then
                ShellExecute Me.hWnd, vbNullString, "http://" & lstDomains.ListItems.ITem(i).Text, vbNullString, "C:\", SW_SHOWNORMAL
        End If
    Next
Exit Sub
errHandle:
    MsgBox "Error mnuPopUpViewBrowser: " & Err.Description
End Sub

Private Sub mnuPopUpWhois_Click()
On Error GoTo errHandle

    Me.txtDomainWhois.Text = lstDomains.ListItems.ITem(ClickedItem).Text
    SSTab.Tab = 2
Exit Sub
errHandle:
    MsgBox "Error: mnuPopUpResult " & Err.Number & " " & Err.Description

End Sub

Private Sub mnuReportsAlexaReport_Click()
    If lstAlexa.ListItems.count = 0 Then
        MsgBox "You have no Alexa information to generate a report!", vbInformation
    Else
        frmReports.Caption = "Alexa Report"
        frmReports.Tag = "3"
        frmReports.Show vbModal, Me
    End If
End Sub

Private Sub mnuReportsDomainReport_Click()
    If lstDomains.ListItems.count = 0 Then
        MsgBox "You have no domains to generate a report!", vbInformation
    Else
        frmReports.Caption = "Domains Report"
        frmReports.Tag = "1"
        frmReports.Show vbModal, Me
    End If
End Sub

Private Sub mnuResultsReport_Click()
    If lstResults.ListItems.count = 0 Then
        MsgBox "You have no results to generate a report!", vbInformation
    Else
        frmReports.Caption = "Results Report"
        frmReports.Tag = "2"
        frmReports.Show vbModal, Me
    End If
End Sub

Private Sub mnuSearch_Click()
    frmSearch.Show
End Sub

Private Sub mnuSortAZ_Click()
    lstDomains.SortOrder = 1
    lstDomains.Sorted = True
    
    mnuSortAZ.Checked = lstDomains.SortOrder = 0
    mnuSortZA.Checked = mnuSortAZ.Checked = False


End Sub

Private Sub mnuSortZA_Click()

    lstDomains.SortOrder = 0
    lstDomains.Sorted = True
    
    mnuSortAZ.Checked = lstDomains.SortOrder = 0
    mnuSortZA.Checked = mnuSortAZ.Checked = False
End Sub

Private Sub tmrAutoSave_Timer()
    Call SaveDomainsBackUp
End Sub

Private Sub txtDomain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdAddDomain_Click
        txtDomain.Text = ""
    End If
    
End Sub


Private Sub txtDomainWhois_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdWhois_Click
        txtDomainWhois.Text = ""
    End If
End Sub
Private Sub LoadDomains()
On Error GoTo errHandle
    'Dim T As Single
    CD.Filename = ""
    CD.DefaultExt = "txt"
    CD.DialogTitle = "Load Domain List (One domain per line)"
    CD.Filter = "Text Files (*.txt)|*.txt"
    CD.ShowOpen
    
    If CD.Filename <> "" Then
        Call LoadDomainsByFile(CD.Filename)

    End If
 '   MsgBox CSng(Timer - T)
Exit Sub
errHandle:
    MsgBox "Error LoadDomains: " & Err.Description
End Sub
Public Sub LoadDomainsByFile(Filename As String)
On Error GoTo errHandle
    'Dim T As Single

       ' T = Timer
        Dim f As Long, strData As String
        f = FreeFile
        Open Filename For Input As #f
            strData = input(LOF(f), f)
        Close #f
        Call MruList.Add(Filename)
        Call MruList.Update(frmMain)
        Dim strTemp() As String
        strTemp = Split(strData, vbCrLf)
        Dim i As Long, g As Long, blnExist As Boolean
        With Me.lstDomains
            For g = 0 To UBound(strTemp)
                'Check if domain is already in list
               ' blnExist = False
                'For i = 1 To .ListItems.Count
               '     If .ListItems.Item(i).Text = strTemp(g) Then
                '       MsgBox "You already have that domain listed.", vbInformation
                 '       blnExist = True
                '    End If
                'Next
                'Sanity checks
                If InStr(strTemp(g), ".") = False Then
                    'MsgBox "You have entered an invalid domain name", vbInformation
                
                Else '
                    'If blnExist = False Then
                        strTemp(g) = Replace(strTemp(g), "http://www", "")
                        strTemp(g) = Replace(strTemp(g), "http://", "")
                        
                        .ListItems.Add , , strTemp(g)
                    'End If
                End If
            Next
        End With


Exit Sub
errHandle:
    MsgBox "Error LoadDomainsByFile: " & Err.Description
End Sub
Private Sub ExportDomains()
On Error Resume Next
    If Me.lstDomains.ListItems.count = 0 Then
        MsgBox "No Data to export. Run Scan Domains first", vbInformation
        Exit Sub
    End If
    CD.Filename = ""
    CD.DialogTitle = "Save CSV"
    CD.DefaultExt = "csv"
    CD.Filter = "CSV Files (*.csv)|*.csv"
    CD.Flags = cdlOFNOverwritePrompt
    CD.ShowSave
    
    If CD.Filename <> "" Then
            Dim f As Long
            f = FreeFile
   
            Open CD.Filename For Output As #f
            
                'Print #f, "Domain, Register, Whois Server, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
                'Dim i As Long
                'If Me.lstDomains.ListItems.Count > 0 Then
                '    For i = 1 To Me.lstDomains.ListItems.Count
                '        Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(2).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(3).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(4).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(5).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(6).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(7).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(8).Text & Chr$(34)
                '    Next
                'End If
                Dim g As Long, strData As String
                For g = 1 To Me.lstDomains.ColumnHeaders.count
                    If g = Me.lstDomains.ColumnHeaders.count Then
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text
                    Else
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text & ", "
                    End If
                Next
                Print #f, strData
 
                Dim i As Long, data As String
                If Me.lstDomains.ListItems.count > 0 Then
                    For i = 1 To Me.lstDomains.ListItems.count
                        data = ""
                        If Me.lstDomains.ListItems.ITem(i).ListSubItems.count <> 0 Then
                            data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                                data = data & "," & Chr$(34) & lstDomains.ListItems.ITem(i).ListSubItems(g).Text & Chr$(34)
                            Next g
                           Print #f, data
                        Else
                            data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                                data = data & "," & Chr$(34) & Chr$(34)
                            Next g
                           Print #f, data
                        End If
                    
                    Next
                End If
                

            Close #f
    End If
Exit Sub
errHandle:
    MsgBox "Error: cmdExportCSV:" & Err.Description
End Sub
Public Sub CheckDomainExpire()
On Error Resume Next
    If Me.lstDomains.ListItems.count = 0 Then
        Exit Sub
    End If
    Dim i As Long
    Dim strData As String
    Dim dCount As Long
    Dim Length As Long
    Dim tDate As String
    tDate = Date
    
    dCount = lstDomains.ListItems.count
    frmExpire.txtExpire.Text = ""
    With lstDomains.ListItems
        For i = 1 To dCount
            
            Length = DateDiff("d", tDate, Mid$(.ITem(i).SubItems(8), 1, 11))
            If Length <= ExpireDays Then
                strData = strData & .ITem(i).Text & " Expires in: " & Length & " days." & vbCrLf
    
            End If
        Next
    End With
    frmExpire.txtExpire.Text = strData
    frmExpire.Show
'DateDiff("d", Date, "03-jan-2006")

End Sub
Public Sub SaveDomains()
'On Error Resume Next
            Dim f As Long
            f = FreeFile
            Close
            Open App.Path & "\data.csv" For Output As #f
            
                'Print #f, "Domain, Register, Whois Server, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
                
                Dim g As Long, strData As String
                For g = 1 To Me.lstDomains.ColumnHeaders.count
                    If g = Me.lstDomains.ColumnHeaders.count Then
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text
                    Else
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text & ", "
                    End If
                Next
                Print #f, strData
 
                Dim i As Long, data As String
                If Me.lstDomains.ListItems.count > 0 Then
                    For i = 1 To Me.lstDomains.ListItems.count
                        data = ""
                        'If Me.lstDomains.ListItems.Item(i).ListSubItems.Count <> 0 Then
                        'Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(2).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(3).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(4).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(5).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(6).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(7).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(8).Text & Chr$(34)
                        'Else
                        '    Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34)
                        'End If
                        If Me.lstDomains.ListItems.ITem(i).ListSubItems.count <> 0 Then
                            data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                                data = data & "," & Chr$(34) & lstDomains.ListItems.ITem(i).ListSubItems(g).Text & Chr$(34)
                            Next g
                           Print #f, data
                        Else
                           '
                            data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstDomains.ColumnHeaders.count - 1
                                data = data & "," & Chr$(34) & Chr$(34)
                            Next g
                           Print #f, data
                            'Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34)
                        End If
                    
                    Next
                End If

            Close #f

Exit Sub
errHandle:
    MsgBox "Error: SaveDomains:" & Err.Description
End Sub
Public Sub LoadDomainInfo()
On Error GoTo errHandle
            Dim f As Long, T As Long
            T = Timer
            f = FreeFile
            Dim strData As String
            Open App.Path & "\data.csv" For Input As #f
                Line Input #f, strData
                
                Dim strCol() As String
                Dim g As Long

                    For g = 1 To lstDomains.ColumnHeaders.count
                        ReDim strCol(lstDomains.ColumnHeaders.count - 1)
                    Next
                    'Print #f, "Domain, Register, Whois Server, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
                    Dim i As Long
                    Dim k As Long
                    Dim z As Long
                    'Dim s As String * 1
                    's = Chr$(34)
                    'MsgBox EOF(f)
                    Do Until EOF(f)
                        'Line Input #F, strData
                        'strData = Replace(strData, s, "")
                       ' For k = 0 To UBound(strCol)
                        '   Input #f, strCol(k)
                       ' Next
                       Line Input #f, strData
                       Dim tempStringArray
                       
                       tempStringArray = Split(strData, Chr$(34) & "," & Chr$(34))
                       For i = 0 To UBound(tempStringArray)
                        tempStringArray(i) = Replace(tempStringArray(i), Chr$(34), "")
                        strCol(i) = tempStringArray(i)
                        'MsgBox tempStringArray(i)
                       Next i
                       
                        'strCol = Split(strData, ",")
                        With lstDomains.ListItems.Add(, , strCol(0))
                        
                            For z = 1 To UBound(strCol)
                               ' .ListItems.ITem(.ListItems.count).SubItems(z) = strCol(z)
                                 .SubItems(z) = strCol(z)
                            Next
                        End With
                    Loop
           
                

            Close #f

Exit Sub
errHandle:
    Exit Sub
    MsgBox "Error: LoadDomains:" & Err.Description
End Sub
Sub LoadMenus()
On Error GoTo nofile:
    ReDim CustomMenu(0)
    Dim f As Long
    Dim strData As String
    f = FreeFile
    Open App.Path & "\menu.ini" For Input As #f
        Do While Not EOF(f)
            Line Input #f, strData
            If Left$(strData, 10) = "MenuTitle=" Then
                ReDim Preserve CustomMenu(UBound(CustomMenu) + 1)
                CustomMenu(UBound(CustomMenu)).MenuTitle = Right$(strData, Len(strData) - 10)
            End If
            If Left$(strData, 9) = "MenuLink=" Then
                CustomMenu(UBound(CustomMenu)).MenuLink = Right$(strData, Len(strData) - 9)
            End If
        Loop
    Close #f
    Call ReloadMenuItems
Exit Sub
nofile:
    Call ReloadMenuItems
End Sub
Sub SaveMenus()
    Dim f As Long
    f = FreeFile
    Open App.Path & "\menu.ini" For Output As #f
        Print #f, ";Custom Menu for Domain Manager Pro"
        Dim i As Long
        For i = 0 To UBound(CustomMenu)
            If CustomMenu(i).MenuLink <> "" And CustomMenu(i).MenuTitle <> "" Then
                Print #f, "MenuTitle=" & CustomMenu(i).MenuTitle
                Print #f, "MenuLink=" & CustomMenu(i).MenuLink
            End If
        Next
    Close f
End Sub
Sub ReloadMenuItems()

    Dim i As Long
    For i = 1 To Me.mnuPopUpCustom.count - 1
       Unload Me.mnuPopUpCustom(i)
    Next
    Dim g As Long

    For i = 0 To UBound(CustomMenu)
      g = Me.mnuPopUpCustom.count - 1
        If CustomMenu(i).MenuTitle <> "" Then
            If i < g Then
                With Me.mnuPopUpCustom(i)
                    .Caption = CustomMenu(i).MenuTitle
                    .Visible = True
                End With
            Else
                Load Me.mnuPopUpCustom(g + 1)
                With Me.mnuPopUpCustom(g + 1)
                    .Caption = CustomMenu(i).MenuTitle
                    .Visible = True
                End With
            End If
        End If
        
        

    Next
End Sub
Private Sub LoadCol()
On Error GoTo nofile:
    Dim f As Long
    Dim strData As String
    f = FreeFile
    lstDomains.ColumnHeaders.Clear
    Open App.Path & "\col.ini" For Input As #f
        Do While Not EOF(f)
            Line Input #f, strData
            If Left$(strData, 9) = "ColTitle=" Then
                lstDomains.ColumnHeaders.Add , , Right$(strData, Len(strData) - 9)
            End If
        Loop
    Close #f

Exit Sub
nofile:
End Sub
Public Function GetColByName(strName As String) As Integer
    Dim i As Long
    For i = 1 To lstDomains.ColumnHeaders.count
        If lstDomains.ColumnHeaders.ITem(i).Text = strName Then
            GetColByName = (i - 1)
            Exit Function
        End If
    Next
End Function
Public Sub SaveDomainsBackUp()
On Error Resume Next
            Dim f As Long
            f = FreeFile
            Close
            Open App.Path & "\databackup.csv" For Output As #f
            
                'Print #f, "Domain, Register, Whois Server, Name Server 1, Name Server 2, Status, Updated Date, Created Date, Expiration Date"
                
                Dim g As Long, strData As String
                For g = 1 To Me.lstDomains.ColumnHeaders.count
                    If g = Me.lstDomains.ColumnHeaders.count Then
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text
                    Else
                        strData = strData & Me.lstDomains.ColumnHeaders.ITem(g).Text & ", "
                    End If
                Next
                Print #f, strData
 
                Dim i As Long, data As String
                If Me.lstDomains.ListItems.count > 0 Then
                    For i = 1 To Me.lstDomains.ListItems.count
                        data = ""
                        'If Me.lstDomains.ListItems.Item(i).ListSubItems.Count <> 0 Then
                        'Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(2).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(3).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(4).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(5).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(6).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(7).Text & Chr$(34) & "," & Chr$(34) & lstDomains.ListItems.Item(i).ListSubItems(8).Text & Chr$(34)
                        'Else
                        '    Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34)
                        'End If
                        If Me.lstDomains.ListItems.ITem(i).ListSubItems.count <> 0 Then
                            data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstDomains.ListItems.ITem(i).ListSubItems.count
                                data = data & "," & Chr$(34) & lstDomains.ListItems.ITem(i).ListSubItems(g).Text & Chr$(34)
                            Next g
                           Print #f, data
                        Else
                           '
                            data = Chr$(34) & Me.lstDomains.ListItems.ITem(i).Text & Chr$(34)
                            For g = 1 To Me.lstDomains.ColumnHeaders.count - 1
                                data = data & "," & Chr$(34) & Chr$(34)
                            Next g
                           Print #f, data
                            'Print #f, Chr$(34) & Me.lstDomains.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34)
                        End If
                    
                    Next
                End If

            Close #f

Exit Sub
errHandle:
    MsgBox "Error: SaveDomains:" & Err.Description
End Sub
Private Sub DomainScannerCache(Range As Long)
On Error GoTo errHandle:
    Dim i As Long, Min As Long, Max As Long
    Dim strDomainName As String
    cmdStop.Visible = True
    cmdAddDomain.Enabled = False
    cmdBulkAddDomain.Enabled = False
    If Range = -1 Then
        Min = 1
        Max = Me.lstDomains.ListItems.count
    Else
        Min = Range
        Max = Range
    End If
    blnStop = False
    
    For i = Min To Max
        If blnStop = True Then
            Exit Sub
        End If
        Dim strData As String
        strDomainName = lstDomains.ListItems.ITem(i).Text
        If FileExists(App.Path & "\rawwhois\" & strDomainName & ".txt") = True Then
            Dim f As Long
            f = FreeFile
            Open App.Path & "\rawwhois\" & strDomainName & ".txt" For Input As #f
                strData = input(LOF(f), f)
            Close #f
        Else
            GoTo nextdomain:
        End If

        Dim J As Long
        Dim strTemp() As String
        If Right$(LCase$(strDomainName), 4) = ".com" Or Right$(LCase$(strDomainName), 4) = ".net" Then
            strTemp = Split(strData, vbLf)
        Else
            strTemp = Split(strData, vbCrLf)
        End If
      
        Dim NS1 As Boolean, TwoNS As Boolean
        TwoNS = False
        NS1 = False
        'MsgBox UBound(strTemp)
        For J = 0 To UBound(strTemp)
            With lstDomains.ListItems.ITem(i)
                If InStr(1, strTemp(J), "Registrar:") <> 0 Then
                    .SubItems(GetColByName("Registrar")) = Trim$(Replace(strTemp(J), "Registrar:", ""))

                ElseIf InStr(1, strTemp(J), "Whois Server:") <> 0 Then
                    .SubItems(GetColByName("Whois Server")) = Trim$(Replace(strTemp(J), "Whois Server:", ""))
                ElseIf InStr(1, strTemp(J), "Name Server:") <> 0 And TwoNS = False Then
    
                    If NS1 = False Then
                       .SubItems(GetColByName("Name Server 1")) = Trim$(Replace(strTemp(J), "Name Server:", ""))
                        NS1 = True
                    Else
                        .SubItems(GetColByName("Name Server 2")) = Trim$(Replace(strTemp(J), "Name Server:", ""))
                        TwoNS = True
                    End If
                ElseIf InStr(1, strTemp(J), "Status:") <> 0 Then
                    .SubItems(GetColByName("Status")) = Trim$(Replace(strTemp(J), "Status:", ""))
                ElseIf InStr(1, strTemp(J), "Updated Date:") <> 0 Then
                    .SubItems(GetColByName("Updated Date")) = Trim$(Replace(strTemp(J), "Updated Date:", ""))
                ElseIf InStr(1, strTemp(J), "Last Updated On:") <> 0 Then
                    .SubItems(GetColByName("Updated Date")) = Trim$(Replace(strTemp(J), "Last Updated On:", ""))
                ElseIf InStr(1, strTemp(J), "Created On:") <> 0 Then
                    .SubItems(GetColByName("Created Date")) = Trim$(Replace(strTemp(J), "Created On:", ""))
                ElseIf InStr(1, strTemp(J), "Creation Date:") <> 0 Then
                    .SubItems(GetColByName("Created Date")) = Trim$(Replace(strTemp(J), "Creation Date:", ""))
                ElseIf InStr(1, strTemp(J), "Expiration Date:") <> 0 Then
                    .SubItems(GetColByName("Expiration Date")) = Trim$(Replace(strTemp(J), "Expiration Date:", ""))
                End If
            End With
            
            
            DoEvents
        Next
nextdomain:
        DoEvents
    Next
    cmdStop.Visible = False
    cmdAddDomain.Enabled = True
    cmdBulkAddDomain.Enabled = True
    'Call CheckDomainExpire
Exit Sub
errHandle:
    MsgBox "Error: DomainScannerCache " & Err.Number & " " & Err.Description
End Sub

