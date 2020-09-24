<html><head><title>Guestbook</title></head><body bgcolor=SILVER><center>
<font size=6 color=BLACK>T-Virus Creations</font><br><br>
</center><%
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim ThisLine
Dim PrintLine
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
GuestBookFile = Server.MapPath("guestbook.txt")
Set InputStream = FileObject.OpenTextFile (GuestBookFile, ForReading, False)
Do While not InputStream.AtEndOfStream
ThisLine = InputStream.ReadLine
PrintLine = PrintLine + ThisLine + "<br>"
Loop
InputStream.Close
Set OutputStream = Nothing
Set FileObject = Nothing
Response.Write PrintLine
%>
<center>
<br><br>
<a href="default.asp">Main</a><br><br>
</center></body></html>
