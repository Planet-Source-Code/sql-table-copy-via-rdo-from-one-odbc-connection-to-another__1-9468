VERSION 5.00
Begin VB.Form pgb 
   ClientHeight    =   870
   ClientLeft      =   2325
   ClientTop       =   5565
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   430
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   120
      Width           =   6435
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   19
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   18
         Left            =   4320
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   17
         Left            =   4080
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   16
         Left            =   3840
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   15
         Left            =   3600
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   14
         Left            =   3360
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   13
         Left            =   3120
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   12
         Left            =   2880
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   11
         Left            =   2640
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   10
         Left            =   2400
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   9
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   8
         Left            =   1920
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   7
         Left            =   1680
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   6
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   5
         Left            =   1200
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   4
         Left            =   960
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   3
         Left            =   720
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   2
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   1
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label lblPerc 
      Alignment       =   2  'Center
      Caption         =   "0% Completed"
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
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "pgb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author : Renier Barnard (renier_barnard@santam.co.za)
'
' Date    : July 1999
'
' Description :
' This code will demonstrate how to make a simple but nice
' looking progress bar. It could be more simple (Using the line command)
' but this looks better. The from_click event will start the progress bar of.
' Try resizing the progress bar form. There is some code to demonstrate
' how to make something like this generic in size !
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const FLAGS = 1
Const HWND_TOPMOST = -1


Option Explicit

Public Function Progress(Value, MaxValue, Optional color As ColorConstants)
' This is the actual progress bar function.

Dim Perc
Dim bb As Integer
'Me.Show

'Get a color to do it in
If color = 0 Then color = vbBlue

'Now work out the percentage (0-100) of where we currently are
Perc = (Value / MaxValue) * 100
lblPerc.ForeColor = color
DoEvents
lblPerc.Caption = Int(Perc) & "% Completed" 'Just the Label Display
Perc = Perc / 5
Perc = Int(Perc)
Perc = Perc - 1

' Now , fill the blocks that need to be filled
For bb = 0 To 19
    DoEvents
    If bb <= Perc Then
        Picture2(bb).BackColor = color ' Done
    Else
        Picture2(bb).BackColor = vbButtonFace ' Not yet Done
    End If
Next bb

DoEvents

End Function

Private Sub Form_Click()
Dim ii As Integer

For ii = 1 To 5000 ' ii = val , 5000 = maxval
    Call Progress(ii, 5000, vbMagenta) ' Call the progressbar function
Next ii

End Sub

Private Sub Form_Load()

Const FLAGS = 1
Const HWND_TOPMOST = -1

'Me.Top = Screen.Height / 2 - Me.Height / 2
'Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = frmMain.Top + (frmMain.Height)
Me.Width = frmMain.Width
Me.Left = frmMain.Left

'Sets form on always on top.
Dim Success As Integer
'Success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
                                                ' Change the "0's" above to position the window.

'Me.Top = Screen.Height / 2 - Me.Height / 2
'Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub Form_Resize()
'Look at the following code to see how the progress bar will resize to
'fit itself , no matter what the size of the form.

Dim bb As Integer


Picture1.Width = Me.Width - 350
Picture1.Height = Me.Height - 500

For bb = 0 To 19
    Picture2(bb).Width = (Picture1.Width) / 20 - 2
    Picture2(bb).Height = (Picture1.Height) - 50
    If bb > 0 Then
        Picture2(bb).Left = Picture2(bb - 1).Left + Picture2(bb - 1).Width
    End If
    DoEvents
Next bb

lblPerc.Width = Picture1.Width
lblPerc.Top = Me.Height - 350
Picture1.Refresh

End Sub

