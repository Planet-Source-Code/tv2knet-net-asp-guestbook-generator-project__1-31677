Attribute VB_Name = "modLanguage"
Dim label(1 To 1000) As String
Public Sub PutLabels()
LabelOn
With frmMain
.l1.Caption = label(1)
.l1.AutoSize = True



End With
End Sub

Sub LabelOn()
label(1) = "T-Virus Guestbook Generator" + vbCrLf
label(1) = label(1) + "Version 1.0 Freeware"


End Sub
