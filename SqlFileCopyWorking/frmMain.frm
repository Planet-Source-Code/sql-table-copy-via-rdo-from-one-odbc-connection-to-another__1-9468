VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "SQL Express"
   ClientHeight    =   7425
   ClientLeft      =   2040
   ClientTop       =   2415
   ClientWidth     =   10605
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   Begin TabDlg.SSTab tbsMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12091
      _Version        =   393216
      Tab             =   1
      TabHeight       =   529
      TabCaption(0)   =   "   Files"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTarget"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSource"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "   Fields"
      TabPicture(1)   =   "frmMain.frx":08A4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraAction"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraOptions2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "   Results"
      TabPicture(2)   =   "frmMain.frx":0E3E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCheck"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraAuto"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraOptions2 
         Caption         =   "Options"
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   5280
         Width           =   10335
         Begin VB.OptionButton optClear 
            Caption         =   "Use Delete"
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   2
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Use Delete to clear the Target File"
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optClear 
            Caption         =   "Use Truncate"
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   1
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Use truncate to clear the target file"
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optClear 
            Caption         =   "No Clear"
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   0
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Do not clear the target file"
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Target Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4815
         Left            =   6120
         TabIndex        =   30
         Top             =   360
         Width           =   4335
         Begin VB.ListBox lstTargetFields 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3630
            ItemData        =   "frmMain.frx":13D8
            Left            =   120
            List            =   "frmMain.frx":13DA
            TabIndex        =   31
            Top             =   615
            Width           =   4095
         End
         Begin VB.Label lblFile 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Field Listing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Source Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4815
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   4335
         Begin VB.ListBox lstSourceFields 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3630
            ItemData        =   "frmMain.frx":13DC
            Left            =   120
            List            =   "frmMain.frx":13DE
            TabIndex        =   28
            Top             =   615
            Width           =   4095
         End
         Begin VB.Label lblFile 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Field Listing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame fraAction 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4815
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start Transfer"
            Height          =   1095
            Left            =   240
            Picture         =   "frmMain.frx":13E0
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   ">>>>>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   ">>>>>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   3120
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4815
         Left            =   -70440
         TabIndex        =   19
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdFields 
            Caption         =   "Get Fields"
            Height          =   1095
            Left            =   240
            Picture         =   "frmMain.frx":196A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   ">>>>>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   ">>>>>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   3120
            Width           =   975
         End
      End
      Begin VB.Frame fraAuto 
         Caption         =   "Data Transfer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   -69480
         TabIndex        =   7
         Top             =   480
         Width           =   2535
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Time Started"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblTime 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Time Ended"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblTime 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblTime 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   8
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.Frame fraCheck 
         Caption         =   "Action List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   4815
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5295
         Begin VB.ListBox lstResults 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00008000&
            Height          =   4350
            ItemData        =   "frmMain.frx":1EF4
            Left            =   120
            List            =   "frmMain.frx":1EF6
            TabIndex        =   18
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame fraSource 
         Caption         =   "Source Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4815
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   4335
         Begin VB.ListBox lstTablesS 
            Height          =   2790
            ItemData        =   "frmMain.frx":1EF8
            Left            =   120
            List            =   "frmMain.frx":1EFA
            TabIndex        =   14
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdConnectS 
            Caption         =   "Connect"
            Height          =   855
            Left            =   120
            Picture         =   "frmMain.frx":1EFC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3840
            Width           =   4095
         End
         Begin VB.Label lblODBCName 
            Alignment       =   2  'Center
            Caption         =   "Not Connected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   3480
            Width           =   4095
         End
         Begin VB.Label lblList 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Table Listing"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame fraTarget 
         Caption         =   "Target Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4815
         Left            =   -68880
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         Begin VB.ListBox lstTablesT 
            Height          =   2790
            ItemData        =   "frmMain.frx":2206
            Left            =   120
            List            =   "frmMain.frx":2208
            TabIndex        =   16
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdConnectT 
            Caption         =   "Connect"
            Height          =   855
            Left            =   120
            Picture         =   "frmMain.frx":220A
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3840
            Width           =   4095
         End
         Begin VB.Label lblODBCName 
            Alignment       =   2  'Center
            Caption         =   "Not Connected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   3480
            Width           =   4095
         End
         Begin VB.Label lblList 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Table Listing"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   4095
         End
      End
   End
   Begin ComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7050
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5953
            Text            =   "System Offline"
            TextSave        =   "System Offline"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12232
            MinWidth        =   8819
            Text            =   "Ready..."
            TextSave        =   "Ready..."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgOff 
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":2514
      Top             =   6000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOn 
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":2A9E
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit Program"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' The following is for the folder browsing dialog box
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Sub cmdDir_Click()
'Opens a Treeview control that displays the directories in a computer
Dim lpIDList As Long
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
szTitle = "Hello World. Click on a directory and it's path will be displayed in a message box"
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    
    If Mid(sBuffer, Len(sBuffer), 1) = "\" Then
    Else
        sBuffer = sBuffer & "\"
    End If
    
    lblDir = sBuffer
End If
End Sub

Private Sub cmdConnectS_Click()
lstTablesS.Clear
Dim Result As String
'Result = InputBox("Please enter the ODBC Connection Name for the Source Database", "StoreProcTester", "dtaDesktop")

SourceConnectTest = SourceInitRDO(Result, frmMain, lstTablesS)

If SourceConnectTest = False Then
    MsgBox ("Could not connect to database")
    lstResults.AddItem "Could not connect to " & Result
    Exit Sub
End If

lblList(0).Visible = True
lstTablesS.Visible = True
lblODBCName(0).Caption = SourceCon.Name

End Sub
Private Sub cmdConnectT_Click()
lstTablesT.Clear
Dim Result As String
'Result = InputBox("Please enter the ODBC Connection Name for the Target Database", "StoreProcTester", "dtaDesktop")

TargetConnectTest = TargetInitRDO(Result, frmMain, lstTablesT)

If TargetConnectTest = False Then
    MsgBox ("Could not connect to database")
    lstResults.AddItem "Could not connect to " & Result
    Exit Sub
End If


lblList(1).Visible = True
lstTablesT.Visible = True
lblODBCName(1).Caption = TargetCon.Name
End Sub

Private Sub cmdExit_Click()
Unload pgb
Unload Me
End Sub

Private Sub Step2()

lstSourceFields.Clear
lstTargetFields.Clear

Dim sTime As Double
Dim bCheck As Boolean
Dim TimeLapse As Double
Dim vBuffer As Variant
Dim iRowsReturned As Long
Dim ii As Long

If SourceConnectTest = False Then
    MsgBox ("Not connected to the Source Database")
    tbsMain.Tab = 0
    Exit Sub
End If

If TargetConnectTest = False Then
    MsgBox ("Not connected to the Target Database")
    tbsMain.Tab = 0
    Exit Sub
End If

If lstTablesS.ListIndex < 0 Then
    MsgBox ("No Source Table Selected")
    tbsMain.Tab = 0
    Exit Sub
End If

If lstTablesT.ListIndex < 0 Then
    MsgBox ("No Target Table Selected")
    tbsMain.Tab = 0
    Exit Sub
End If

'Alright , now pay attention. We are going to create 3 files , so we need to
' 3 main looping routines
sTime = timeGetTime
frmMain.lblTime(0) = Time()
bCheck = GetFields(lstTablesS, frmMain, lstSourceFields, SourceCon)

bCheck = GetFields(lstTablesT, frmMain, lstTargetFields, TargetCon)


DoEvents

frmMain.lblTime(1) = Time()
frmMain.lblTime(2) = (((timeGetTime - sTime) / 1000) / 60) & " Minutes"

End Sub


Private Sub cmdFields_Click()
tbsMain.Tab = 1
End Sub

Private Sub Step3()

Dim sTime As Double
Dim bCheck As Boolean
Dim TimeLapse As Double
Dim vBuffer As Variant
Dim iRowsReturned As Long
Dim ii As Long

If SourceConnectTest = False Then
    MsgBox ("Not connected to the Source Database")
    tbsMain.Tab = 1
    Exit Sub
End If

If TargetConnectTest = False Then
    MsgBox ("Not connected to the Target Database")
    tbsMain.Tab = 1
    Exit Sub
End If

'Alright , now pay attention. We are going to create 3 files , so we need to
' 3 main looping routines
sTime = timeGetTime
frmMain.lblTime(0) = Time()


bCheck = ClearFile(lstTablesT, frmMain, TargetCon)

DoEvents
'bCheck = Transfer(lstTablesS.List(lstTablesS.ListCount), _
                                    lstTablesT.List(lstTablesT.ListCount), frmMain)
  bCheck = Transfer(lstTablesS, lstTablesT, frmMain)

'vBuffer = getResults(txtInput)
'iRowsReturned = UBound(vBuffer, 2) + 1
'
'For ii = 0 To iRowsReturned - 1
'    txtResult = txtResult & vBuffer(0, ii) & Chr(9) & vBuffer(1, ii) & Chr(9) & vBuffer(2, ii) & vbCrLf
'Next ii


frmMain.lblTime(1) = Time()
frmMain.lblTime(2) = (((timeGetTime - sTime) / 1000) / 60) & " Minutes"

End Sub

Private Sub cmdTest_Click()
Call CreateTest
End Sub

Private Sub cmdStart_Click()
tbsMain.Tab = 2
End Sub

Private Sub Form_Load()
'Center the form on the screen
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
stb.Panels(1).Text = "System Online"
tbsMain.Tab = 0
'sBuffer = "c:\"
'lblDir = sBuffer
End Sub

Private Sub fraAuto_DblClick()
ShowTest
End Sub

Private Sub lstTablesS_Click()

lblFile(0).Caption = "Fields for " & lstTablesS
Dim iCount As Integer

For iCount = 0 To lstTablesT.ListCount - 1
    If Trim(UCase(lstTablesT.List(iCount))) = Trim(UCase(lstTablesS)) Then
        lstTablesT.ListIndex = iCount
    End If
    
Next iCount

End Sub

Private Sub lstTablesT_Click()
lblFile(1).Caption = "Fields for " & lstTablesT

End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub tbsMain_Click(PreviousTab As Integer)
If tbsMain.Tab = 1 Then Call Step2
If tbsMain.Tab = 2 Then Call Step3
Call SetOn(tbsMain, tbsMain.Tab, Me)

End Sub

