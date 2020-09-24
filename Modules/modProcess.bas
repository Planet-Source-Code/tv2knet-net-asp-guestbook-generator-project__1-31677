Attribute VB_Name = "modProcess"
Public RunsApp As Boolean
Public InputT As Boolean
Public InputStrIn As String

Public Function ProcessCommands(Commands As String, Out As Object)
Dim x As String
Dim Y As String
Dim ps As Long

x = Commands
Out.Text = GetHeadADD
frmMain.mpage.Text = GetFore

For i = 1 To Len(x)
GoToLast Out

ps = 0





Y = Mid$(x, i, 7) '= "#MAIN"
If UCase$(Y) = "#INPUT(" Then
For p = i To Len(x)
Y = Mid$(x, p, 1) '= "#MAIN"
ps = ps + 1
If Y = ")" Then
PrintT Mid$(x, i + 1 + 6, ps - 8), Out
GoTo 10
End If


Next
10

i = i + 6
End If


GoToLast Out
Next
'CallEnd
Out.Text = Out.Text + vbCrLf + GetFootADD
frmMain.mpage.Text = frmMain.mpage.Text + GetBack

End Function

Sub GoToLast(Out As Object)
Out.SelStart = Len(Out.Text)
End Sub
Sub CallErrorT(ErrorNmr As Long)
If ErrorNmr = 1 Then MsgBox "Wrong syntax with #MAIN"
If ErrorNmr = 2 Then MsgBox "Wrong syntax with #END"
If ErrorNmr = 3 Then MsgBox "Wrong syntax with PRINT "

End Sub
Function GetFore() As String
Dim x As String
x = "<html>" + vbCrLf
x = x + vbCrLf
x = x + "<head>" + vbCrLf
x = x + "<meta http-equiv=""Content-Language"" content=""af"">" + vbCrLf
x = x + "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" + vbCrLf
x = x + "<meta name=""GENERATOR"" content=""Microsoft FrontPage 4.0"">" + vbCrLf
x = x + "<meta name=""ProgId"" content=""FrontPage.Editor.Document"">" + vbCrLf
x = x + "<title>" + frmMain.gbtitle + "</title>" + vbCrLf
x = x + "</head>" + vbCrLf
x = x + "<body>" + vbCrLf
x = x + "<form method=""POST"" action=""add.asp"">" + vbCrLf
GetFore = x
End Function

Function GetBack() As String
Dim x As String

x = x + "<p><input type=""hidden"" name=""add"" value=""1"" size=""20"">" + vbCrLf
x = x + "<p><input type=""submit"" value=""Send"" name=""B1""><input type=""reset"" value=""Reset"" name=""B2""></p>" + vbCrLf
x = x + "</form>" + vbCrLf
x = x + "<p align=""left"">&nbsp;</p>" + vbCrLf
x = x + "<a href=""view.asp"">View</a>" + vbCrLf
x = x + "</body>" + vbCrLf
x = x + "</html>" + vbCrLf
GetBack = x
End Function
Sub PrintT(Data As String, Box As Object)
Dim x As String
x = "OutputStream.WriteLine """ + LCase(Data) + " :" + """ +  request.form(""" + Data + """)"
Box.Text = Box.Text + x + vbCrLf
x = "<p>" + Data + ": <input type=""text"" name=""" + Data + """ size=""50""></p>"
frmMain.mpage = frmMain.mpage + vbCrLf + x

End Sub


Function GetFileContent(FileLoc As String) As String
Dim x As String
Dim t As String
t = ""
On Error GoTo 10
Open FileLoc For Input As #1
While EOF(1) = False
Input #1, x
If t = "" Then
t = x
GoTo 2
End If
t = t + vbCrLf + x
2
Wend
Close #1
GetFileContent = t
10
End Function


Sub ClearT(Out As Object)
Out.Text = ""


End Sub


Public Function GetHeadADD() As String
Dim x As String
x = x + "<html><body>" + vbCrLf
x = x + "<%" + vbCrLf + "if request.Form(""Add"")=""1"" then" + vbCrLf
x = x + "Const ForReading = 1, ForWriting = 2, ForAppending = 8" + vbCrLf
x = x + "Set FileObject = Server.CreateObject(""Scripting.FileSystemObject"")" + vbCrLf
x = x + "GuestBookFile = Server.MapPath(""" + frmMain.gname.Text + """)" + vbCrLf
x = x + "Set OutputStream = FileObject.OpenTextFile(GuestBookFile, ForAppending, True)" + vbCrLf

GetHeadADD = x

End Function

Public Function GetFootADD() As String
Dim x As String
x = x + "OutputStream.WriteLine "" """ + vbCrLf
x = x + "OutputStream.Close" + vbCrLf
x = x + "Set OutputStream = Nothing" + vbCrLf
x = x + "Set FileObject = Nothing" + vbCrLf
x = x + "end if" + vbCrLf
x = x + "%>"
x = x + "<a href=""" + frmMain.dname.Text + """>Main</a>" + vbCrLf
x = x + "<a href=""view.asp"">View</a>" + vbCrLf
x = x + "</body></html>"

GetFootADD = x

End Function

