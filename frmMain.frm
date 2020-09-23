VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empty Directory Finder (Developed By Farhan Alam)"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCopy 
      BackColor       =   &H00BFBCB0&
      Caption         =   "Copy To Clip"
      Height          =   555
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4740
      Width           =   1965
   End
   Begin VB.CommandButton cmdDeleteAll 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Delete All Directory "
      Enabled         =   0   'False
      Height          =   555
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4740
      Width           =   1965
   End
   Begin MSComCtl2.Animation ani 
      Height          =   525
      Left            =   8775
      TabIndex        =   9
      Top             =   4845
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   926
      _Version        =   393216
      Center          =   -1  'True
      BackStyle       =   1
      BackColor       =   0
      FullWidth       =   84
      FullHeight      =   35
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   15
      TabIndex        =   7
      Top             =   495
      Width           =   3270
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   3390
      TabIndex        =   5
      Top             =   150
      Width           =   6690
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Delete Directory"
      Height          =   555
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1965
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open Parent Directory"
      Height          =   555
      Left            =   2220
      TabIndex        =   3
      Top             =   4095
      Width           =   1965
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Directory"
      Height          =   555
      Left            =   2220
      TabIndex        =   2
      Top             =   4740
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BFBCB0&
      Caption         =   "Find Empty Directories"
      Height          =   555
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4095
      Width           =   1965
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   150
      Width           =   3285
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   9345
      TabIndex        =   8
      ToolTipText     =   "Number of Empty Directories"
      Top             =   4140
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1875
      Left            =   8595
      Top             =   3885
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   5430
      Width           =   8235
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BFBCB0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   2865
      Left            =   -30
      Top             =   3885
      Width           =   8625
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'EmptyFolder Devloped By Farhan Alam
'If you need me let me know at farhanalam1@hotmail.com
Private StopIt As Boolean

Private Sub CmdCopy_Click()
Dim str As String
Dim i As Integer
  For i = 0 To List1.ListCount - 1
      DoEvents
      str = str & List1.List(i) & vbCrLf
      Label1.Caption = i
  Next i
  Clipboard.Clear
  Clipboard.SetText str
End Sub

Private Sub cmdDeleteAll_Click()
Dim i As Integer
Dim rspn As Integer
Dim cnt As Integer
 rspn = MsgBox("Are you sure you want to delete all empty folders", vbYesNo, "Developed By Farhan Alam")
If rspn = vbYes Then
   ani.Open App.Path & "\farstone.avi"
   ani.Play
A:
    For i = 0 To List1.ListCount - 1
        DoEvents
        RmDir List1.List(i)
        List1.RemoveItem i
        cnt = cnt + 1
        lblCnt.Caption = cnt
        
        GoTo A:
    Next i
End If
ani.Stop
ani.Close
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Stop Search" Then
StopIt = True
Exit Sub
End If
Reset
StopIt = False
Command1.Caption = "Stop Search"
ani.Open App.Path & "\farstone.avi"
ani.Play
Dim this As String
If Len(Dir1.Path) = 3 Then
  this = Dir1.Path
Else
this = Dir1.Path & "\"  'Left(Drive1.Drive, 2) & "\"
End If
Dim test As String
On Error Resume Next
test = Dir(this & "*", vbDirectory Or vbArchive Or vbHidden Or vbReadOnly Or vbSystem)

  If Err Then
     MsgBox "Cannot access " & this & ".", vbCritical, "Error accessing drive"
  Else
On Error GoTo 0
GoRecursion this
End If
Screen.MousePointer = 0
Command1.Caption = "Find Empty Directories"
StopIt = False
ani.Stop
ani.Close
If List1.ListCount = 0 Then MsgBox "No empty directories found.", vbInformation, "Finished": Exit Sub
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
cmdDeleteAll.Enabled = True
End Sub

Private Sub Command2_Click()
If GetSelected = "" Then MsgBox "Please select a directory.", vbInformation, "Error": Exit Sub
Shell "explorer " & GetSelected, vbNormalFocus
End Sub

Private Sub Command3_Click()
If GetSelected = "" Then MsgBox "Please select a directory.", vbInformation, "Error": Exit Sub
Dim this As String
this = FixDir(GetSelected)
this = Left(this, Len(this) - 1)
If Len(this) = 2 Then MsgBox "This is a root directory.", vbInformation, "Error": Exit Sub
this = Left(this, InStrRev(this, "\"))
Shell "explorer " & this, vbNormalFocus
End Sub

Private Sub Command4_Click()
If GetSelected = "" Then MsgBox "Please select a directory.", vbInformation, "Error": Exit Sub
If MsgBox("Are you sure you want to delete " & GetSelected & "?", vbQuestion Or vbYesNo, "Remove Directory") = vbYes Then
On Error Resume Next
RmDir GetSelected
If Err Then
MsgBox "Error deleting " & GetSelected & ".", vbCritical, "Error"
Exit Sub
End If
On Error GoTo 0
Dim this As Long
this = List1.ListIndex
If this = 0 And List1.ListCount > 1 Then List1.ListIndex = 1 Else List1.ListIndex = this - 1
List1.RemoveItem this
If List1.ListCount = 0 Then Reset
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Reset
''Drive1.Drive = "X:\"
'Dir1.Path = "X:\Works\Scan"
End Sub

Private Sub Reset()
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
List1.Clear
Label1.Caption = ""
End Sub

Private Sub List1_Click()
Label1.Caption = GetSelected
End Sub

Private Function GetSelected() As String
If List1.ListIndex = -1 Then GetSelected = "": Exit Function
GetSelected = List1.List(List1.ListIndex)
End Function

Private Sub List1_DblClick()
If Command2.Enabled = True Then Command2.Value = True
End Sub

Private Function FixDir(ByVal this As String) As String
FixDir = Trim(this)
If Right(FixDir, 1) <> "\" Then FixDir = FixDir & "\"
End Function

Private Sub GoRecursion(ByVal Directory As String)
DoEvents

Dim test As String
Dim found As Boolean
Dim ComeBack As String

Restart:
test = Dir(Directory & "*", vbDirectory Or vbArchive Or vbHidden Or vbReadOnly Or vbSystem)
Do Until test = "" Or StopIt = True
If ComeBack <> "" Then
  If test = ComeBack Then ComeBack = ""
Else
If test <> "." And test <> ".." Then
  found = True
  DoEvents
  If (GetAttr(Directory & test) And vbDirectory) = vbDirectory Then
    GoRecursion FixDir(Directory & test)
    Dir1.Path = Directory & test
    ComeBack = test
    GoTo Restart
  End If
End If
End If
test = Dir
Loop

If found = False And StopIt = False Then
  List1.AddItem Directory, 0
  lblCnt.Caption = List1.ListCount
  DoEvents
End If

End Sub
