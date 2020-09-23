VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Encryption Alogrythm"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Clear"
      Height          =   495
      Left            =   9480
      TabIndex        =   9
      Top             =   6960
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   6960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C&opy"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Character Count"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8160
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Load"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   7530
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11668
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Decrypt"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Encrypt"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bd As Long
Private Sub Command1_Click()
On Error Resume Next
Encrypt text1.Text, text1, 100, ProgressBar1
reverse text1.Text, text1, ProgressBar1
'Use the above line by simply removing the "'"
'to add a lot of extra security to it, but
'expect delays
End Sub
Private Sub Command2_Click()
On Error Resume Next
reverse text1.Text, text1, ProgressBar1
'Use the above line by simply removing the "'"
'to add a lot of extra security to it, but
'expect delays
Decrypt text1.Text, text1, 100, ProgressBar1
End Sub

Private Sub Command3_Click()
On Error GoTo d
cd.CancelError = True
cd.Filter = "Encrypted File Format|*.eff|All Files|*.*"
cd.ShowOpen
text1.LoadFile cd.filename, rtfText
d:
Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo d
cd.CancelError = True
cd.Filter = "Encrypted File Format|*.eff|All Files|*.*"
cd.ShowSave
text1.SaveFile cd.filename, rtfText
d:
Exit Sub

End Sub

Private Sub Command5_Click()
MsgBox Len(text1.Text)
End Sub

Private Sub Command6_Click()
Clipboard.SetText text1.Text
End Sub

Private Sub Command7_Click()
text1.Text = ""
End Sub

Private Sub Timer1_Timer()
bd = bd + 1
Label1.Caption = bd
End Sub
