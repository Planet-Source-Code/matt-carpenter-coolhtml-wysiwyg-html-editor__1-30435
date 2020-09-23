VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "CoolHTML"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      ExtentX         =   12091
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Navigate "about:blank"


End Sub

Private Sub Form_Resize()
WebBrowser1.Top = 40
WebBrowser1.Left = 40
WebBrowser1.Width = Me.Width - 300
WebBrowser1.Height = Me.Height / 2 - 80

RichTextBox1.Top = Me.Height / 2 + 200
RichTextBox1.Left = 40
RichTextBox1.Width = Me.Width - 300
RichTextBox1.Height = Me.Height / 2 - 950
End Sub

Private Sub mnuOpen_Click()
filetoopen = InputBox("Open what file?")
On Error GoTo errhndlr
RichTextBox1.LoadFile filetoopen
Open "C:\temp.html" For Output As #1: Print #1, RichTextBox1.Text: Close #1
WebBrowser1.Navigate "C:\temp.html"
Exit Sub
errhndlr:
MsgBox "There was an error opening the file"

End Sub

Private Sub mnuSave_Click()
filetosave = InputBox("Save to what file?")
Open filetosave For Output As #1: Print #1, RichTextBox1.Text: Close #1


End Sub

Private Sub RichTextBox1_Change()
On Error Resume Next
DoEvents
Open "C:\temp.html" For Output As #1: Print #1, RichTextBox1.Text: Close #1
DoEvents
WebBrowser1.Navigate "C:\temp.html"


End Sub
