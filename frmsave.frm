VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save TEXT file"
   ClientHeight    =   1860
   ClientLeft      =   3690
   ClientTop       =   3555
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Select an option the click this button to save"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Save normally by selecting a file and folder"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Create new directory and save text file."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim directory$, directory2$
If Option1.Value = True Then
directory$ = InputBox$("Enter the full path of the directory to save your web pages", "Test Maker")
On Error GoTo ErrHand
MkDir (directory$)
prompt$ = "Please now enter the name of the new .TXT" & vbCrLf & "DON'T FORGET TO PUT THE .txt EXTENSION!!"
directory2$ = InputBox(prompt$, "Test Maker")
Dim ff
ff = FreeFile
Open directory$ & "\" & directory2$ For Output As #ff
Print #ff, Form1.Text1.Text
Close #ff
DoEvents
MsgBox "Your file has been saved!", , "Test Maker"
Unload Form2
Form1.Show
Exit Sub
ErrHand:
Call ErrHandler
End If

If Option2.Value = True Then
CancelError = False

Dim strfilter As String
Dim fff As Integer
fff = FreeFile
strfilter = "TEXT files (*.txt)|*.txt"
cd2.Filter = strfilter
cd2.InitDir = App.Path
cd2.DialogTitle = "Save your Web Page"
cd2.ShowSave
If cd2.filename <> "" Then
Open cd2.filename For Output As #fff
Print #fff, Form1.Text1.Text
Close #fff
End If

End If


End Sub

Private Sub Form_Load()
Me.Icon = Form1.Icon

End Sub
