VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebPage Test Maker"
   ClientHeight    =   6945
   ClientLeft      =   4245
   ClientTop       =   3465
   ClientWidth     =   7455
   Icon            =   "TestMaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7455
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6360
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   6000
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to view RESULTS in your Browser.  (Save .txt before viewing to see changes)"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "TestMaker.frx":0E42
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label6 
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "Where HTML is stored:"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Current .TXT file in use:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Where to send RESULTS (email):"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Title that appears on test:"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currenthtml As String
Dim email As String
Dim title As String

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Open "testmaker.ini" For Input As 1
Line Input #1, email
Text3.Text = email
Close
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim X%
X% = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit program")
If X% = vbNo Then
Cancel = 1
Exit Sub

End If
End Sub

Private Sub mnAbout_Click()
Load frmabout
frmabout.Show
End Sub

Private Sub mnExit_Click()
Unload Me
End Sub

Private Sub mnHelp_Click()
MsgBox "Open the existing example file called: Minnesota.txt "
'future help info
End Sub

Private Sub mnOpen_Click()
Dim ff
Dim strfilter, strlines, alltext As String
ff = FreeFile
strfilter = "Text Files (*.txt)|*.txt"
cd1.Filter = strfilter
cd1.ShowOpen
If cd1.filename <> "" Then
Open cd1.filename For Input As #ff
Do Until EOF(ff)
Line Input #ff, strlines
alltext = alltext & strlines & vbCrLf
Text1.Text = alltext
Loop
End If
cd1.CancelError = False
Close #ff
Label1.Caption = cd1.filename
For y = 1 To Len(cd1.filename)
   If Mid$(cd1.filename, y, 1) = "\" Then
   title = ""
   End If
   If Mid$(cd1.filename, y, 1) <> "\" Then
   title = title + Mid$(cd1.filename, y, 1)
   End If
Next y
title = Left$(title, Len(title) - 4)
Text2.Text = title
currenthtml = Left$(cd1.filename, Len(cd1.filename) - 4) & ".htm"
Label6.Caption = currenthtml
End Sub


Private Sub mnSave_Click()
Load Form2
Form2.Show

End Sub


Private Sub Command1_Click()
Call saveini

quot = Chr$(34)
currenthtm = "noname.htm"
If Text2.Text <> "" Then
currenthtm = Text2.Text & ".htm"
End If
Open currenthtm For Output As 1
Open cd1.filename For Input As 2
Print #1, "<html>"

Print #1, "<head>"
Print #1, "<meta http-equiv=" & quot & "Content-Type" & quot
Print #1, "content=" & quot & "text/html; charset=iso-8859-1" & quot & ">"
Print #1, "<meta name=" & quot & "GENERATOR" & quot & "content=" & quot & "Microsoft FrontPage Express 2.0" & quot & ">"
Print #1, "<title>" & Text2.Text & "</title>"
Print #1, "</head>"
Print #1, " "
Print #1, "<body background=" & quot & "paper.gif" & quot & ">"
Print #1, " "
Print #1, "<form method=post action=" & quot & "http://www.thalasson.com/cgi-bin/www.thalasson.com/mf.pl" & quot & ">"
Print #1, "<font face=" & quot & "Verdana, Arial, Helvetica, sans-serif" & quot & " size=" & quot & "2" & quot & ">"
Print #1, "<input type=" & quot & "hidden" & quot & " name=emailto value=" & quot & Text3.Text & quot & ">"
Print #1, "<input type=" & quot & "hidden" & quot & " name=subject value=" & quot & Text2.Text & quot & ">"
'
'
'  There are some things you will need to change here for your own use ...
'
'
'  The next two lines are -- useraddr to respond back to and username to respond back to.
'  This information is not known until the person enters it and hits submit ... so I just put
'  in a generic text here ... unless you know ahead of time who is submitting the survey.
Print #1, "<input type=" & quot & "hidden" & quot & " name=useraddr to value=" & quot & "Test/Survey" & quot & ">"
Print #1, "<input type=" & quot & "hidden" & quot & " name=username value=" & quot & "Test/Survey" & quot & ">"
'
'
'   The next line is where to go after the person hits the submit button ... a thank you page?
Print #1, "<input type=" & quot & "hidden" & quot & " name=nextpage value=" & quot & "http://www.inkandline.com/orderthanks.htm" & quot & ">"
'
'
Print #1, "</font>"
Print #1, "    <p><font size=" & quot & "6" & quot & ">" & Text2.Text & "</font></p>"
Print #1, "    <hr>"
Print #1, "    <p>Enter your name: <input type=" & quot & "text" & quot & " size=" & quot & "20" & quot
Print #1, "    name=" & quot & "UserName" & quot & "></p>"
Print #1, "    <p>Enter your email address: <input type=" & quot & "text" & quot & " size=" & quot & "20" & quot
Print #1, "    name=" & quot & "UserEmail" & quot & "></p>"

Do While EOF(2) = False
Line Input #2, lin
  If Left$(lin, 1) = "[" Then
     If Mid$(lin, 2, 1) = "r" Or Mid$(lin, 2, 1) = "R" Then
       Print #1, "<input type=" & quot & "radio" & quot & " name=" & quot & qn & quot & " value=" & quot & Mid$(lin, 4, 3) & quot & ">" & Right(lin, Len(lin) - 3) & "<br>"
     End If
     If Mid$(lin, 2, 1) = "c" Or Mid$(lin, 2, 1) = "C" Then
       Print #1, "<input type=" & quot & "checkbox" & quot & " name=" & quot & qn & ")" & Mid$(lin, 4, 3) & quot & ">" & Right(lin, Len(lin) - 3) & "<br>"
     End If
     If Mid$(lin, 2, 1) = "s" Or Mid$(lin, 2, 1) = "S" Then
       Print #1, "<input type=" & quot & "text" & quot & " size=" & quot & "40" & quot & " name=" & quot & qn & quot & "><br>"
     End If
     If Mid$(lin, 2, 1) = "l" Or Mid$(lin, 2, 1) = "L" Then
       Print #1, "<textarea name=" & quot & qn & quot & " rows=" & quot & "4" & quot & " cols=" & quot & "40" & quot & "></textarea><br>"
     End If
     If Asc(Mid$(lin, 2, 1)) > 47 And Asc(Mid$(lin, 2, 1)) < 58 Then
       qn = ""
       t = 0
       For y = 2 To 4
           If Mid$(lin, y, 1) <> "]" And t = 0 Then
           qn = qn + Mid$(lin, y, 1)
           End If
           If Mid$(lin, y, 1) = "]" Then
           t = 1
           End If
       Next y
       Print #1, "</p>"
       Print #1, "<hr>"
       Print #1, "<p>" & qn & ") " & Right$(lin, Len(lin) - 3) & "<br>"
       Print #1, "<br>"
     End If
  End If
Loop

Print #1, "<hr>"
Print #1, "<p>Complete all questions, Check your answers ...</p>"
Print #1, "<p>Click FINISHED button when test is complete.</p>"
Print #1, "<input type=" & quot & "submit" & quot & " name=" & quot & Time$ & "|" & Date$ & quot & " value=" & quot & "FINISHED" & quot & "></p>"
Print #1, "<hr>"
Print #1, "</form>"
Print #1, "</body>"
Print #1, "</html>"

Close 2
Close 1
Dim m As Long
m = ShellExecute(Me.hwnd, "save", currenthtm, "", App.Path, 1)

m = ShellExecute(Me.hwnd, "open", currenthtm, "", App.Path, 1)

End Sub
Private Sub saveini()
' Save some settings
Open "testmaker.ini" For Output As 1
Print #1, Text3.Text
Close
End Sub
