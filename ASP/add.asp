<html><body>
<%
if request.Form("Add")="1" then
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
GuestBookFile = Server.MapPath("guestbook.txt")
Set OutputStream = FileObject.OpenTextFile(GuestBookFile, ForAppending, True)
OutputStream.WriteLine "name :" +  request.form("name")
OutputStream.WriteLine "email :" +  request.form("email")
OutputStream.WriteLine "comments :" +  request.form("comments")

OutputStream.WriteLine " "
OutputStream.Close
Set OutputStream = Nothing
Set FileObject = Nothing
end if
%><a href="default.asp">Main</a>
<a href="view.asp">View</a>
</body></html>
