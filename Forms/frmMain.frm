VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Virus Creations Guestbook Generator"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   7250
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Start"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "l1"
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(3)=   "Label8"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "odir"
      Tab(1).Control(5)=   "gname"
      Tab(1).Control(6)=   "dname"
      Tab(1).Control(7)=   "gbtitle"
      Tab(1).Control(8)=   "Drive1"
      Tab(1).Control(9)=   "Dir1"
      Tab(1).Control(10)=   "Command2"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Inputs"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "View Page"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Command1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "GenVIEW"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Add Page"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "GenADD"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Main"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "mpage"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.CommandButton Command2 
         Caption         =   "Save Guestbook"
         Height          =   330
         Left            =   -70800
         TabIndex        =   19
         Top             =   3495
         Width           =   1800
      End
      Begin VB.TextBox GenVIEW 
         Height          =   2745
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   660
         Width           =   5790
      End
      Begin VB.TextBox mpage 
         Height          =   2745
         Left            =   -74790
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   660
         Width           =   5790
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   -70800
         TabIndex        =   15
         Top             =   975
         Width           =   1800
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -70800
         TabIndex        =   14
         Top             =   2970
         Width           =   1800
      End
      Begin VB.TextBox GenADD 
         Height          =   2745
         Left            =   -74790
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   660
         Width           =   5790
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generate ADD page"
         Height          =   330
         Left            =   -73005
         TabIndex        =   12
         Top             =   3495
         Width           =   1800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate VIEW page"
         Height          =   330
         Left            =   1995
         TabIndex        =   11
         Top             =   3495
         Width           =   1800
      End
      Begin VB.TextBox Text1 
         Height          =   2640
         Left            =   -74895
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "frmMain.frx":00A8
         Top             =   450
         Width           =   6000
      End
      Begin VB.TextBox gbtitle 
         Height          =   285
         Left            =   -74895
         TabIndex        =   7
         Text            =   "Guestbook"
         Top             =   2865
         Width           =   3270
      End
      Begin VB.TextBox dname 
         Height          =   285
         Left            =   -74895
         TabIndex        =   5
         Text            =   "default.asp"
         Top             =   1920
         Width           =   3270
      End
      Begin VB.TextBox gname 
         Height          =   285
         Left            =   -74895
         TabIndex        =   3
         Text            =   "guestbook.txt"
         Top             =   975
         Width           =   3270
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "You can't change much at the moment but it does include Unlimited field support using a small scripting engine"
         Height          =   645
         Left            =   -74895
         TabIndex        =   22
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "This will  create a simple but working Ready-To-Use ASP Guestbook. "
         Height          =   540
         Left            =   -74895
         TabIndex        =   21
         Top             =   1290
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "This program is under development. About 50% ready and 100% working. NOT bug free. Freeware"
         Height          =   540
         Left            =   -74895
         TabIndex        =   20
         Top             =   3390
         Width           =   3690
      End
      Begin VB.Label odir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -74895
         TabIndex        =   17
         Top             =   3600
         Width           =   45
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the output directory"
         Height          =   225
         Left            =   -74895
         TabIndex        =   16
         Top             =   3285
         Width           =   3165
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":00D8
         Height          =   720
         Left            =   -74895
         TabIndex        =   9
         Top             =   3180
         Width           =   5385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Give the title for the Guestbook's View Page"
         Height          =   195
         Left            =   -74895
         TabIndex        =   6
         Top             =   2445
         Width           =   3120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Give the location of the default file:"
         Height          =   195
         Left            =   -74895
         TabIndex        =   4
         Top             =   1500
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Give the filename of the file that will contain the Guestbook inputs:"
         Height          =   195
         Left            =   -74895
         TabIndex        =   2
         Top             =   555
         Width           =   4665
      End
      Begin VB.Label l1 
         BackStyle       =   0  'Transparent
         Height          =   540
         Left            =   -74895
         TabIndex        =   1
         Top             =   555
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
GenVIEW.Text = GenViewT
End Sub

Private Sub Command2_Click()
StartSave
End Sub

Private Sub Command3_Click()
'GenADD.Text = GenAddT
ProcessCommands Text1.Text, GenADD

End Sub

Private Sub Dir1_Change()
odir.Caption = Dir1.Path

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
PutLabels
End Sub


Public Function GenViewT() As String
Dim x As String
x = "<html><head><title>" + gbtitle.Text + "</title></head><body bgcolor=SILVER><center>" + vbCrLf
x = x + "<font size=6 color=BLACK>T-Virus Creations</font><br><br>" + vbCrLf
x = x + "</center><%" + vbCrLf
x = x + "Const ForReading = 1, ForWriting = 2, ForAppending = 8" + vbCrLf
x = x + "Dim ThisLine" + vbCrLf + "Dim PrintLine" + vbCrLf + "Set FileObject = Server.CreateObject(""Scripting.FileSystemObject"")" + vbCrLf
x = x + "GuestBookFile = Server.MapPath(""" + gname.Text + """)" + vbCrLf
x = x + "Set InputStream = FileObject.OpenTextFile (GuestBookFile, ForReading, False)" + vbCrLf
x = x + "Do While not InputStream.AtEndOfStream" + vbCrLf
x = x + "ThisLine = InputStream.ReadLine" + vbCrLf
x = x + "PrintLine = PrintLine + ThisLine + ""<br>""" + vbCrLf
x = x + "Loop" + vbCrLf
x = x + "InputStream.Close" + vbCrLf
x = x + "Set OutputStream = Nothing" + vbCrLf
x = x + "Set FileObject = Nothing" + vbCrLf
x = x + "Response.Write PrintLine" + vbCrLf
x = x + "%>" + vbCrLf
x = x + "<center>" + vbCrLf
x = x + "<br><br>" + vbCrLf
x = x + "<a href=""" + dname.Text + """>Main</a><br><br>" + vbCrLf
x = x + "</center></body></html>"
GenViewT = x

End Function


Sub StartSave()
On Error GoTo 10
Command1_Click
Command3_Click
Dim x As String
Dim t As String
x = odir.Caption
If Right$(x, 1) <> "\" Then x = x + "\"
t = GenVIEW.Text

Open x + "view.asp" For Output As #1

Print #1, t
Close #1

Open x + "add.asp" For Output As #1

t = GenADD.Text
Print #1, t
Close #1

Open x + dname.Text For Output As #1

t = mpage.Text
Print #1, t
Close #1
Exit Sub
10
MsgBox "Error Occured", vbCritical, "Error"

End Sub

