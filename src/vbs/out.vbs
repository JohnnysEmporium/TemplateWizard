Dim fname, prio, incNo, stat

fname = WScript.Arguments(0)
prio = WScript.Arguments(1)
incNo = WScript.Arguments(2)
stat = WScript.Arguments(3)

srcDir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")-4)
messagesDir = srcDir & "Messages\"
docFile = srcDir & "output.docx"
msgFile = messagesDir & fname

Set oWord = CreateObject("Word.Application")
Set oDoc = oWord.Documents.Open(docFile)
oWord.visible = False
oDoc.tables(1).Range.Copy

Set oOutlook = CreateObject("Outlook.Application")
Set oMsg = oOutlook.CreateItemFromTemplate(msgFile)

with oMsg
	.Subject = "PRIORITY " & prio & " - SNOW REF " & incNo & " - " & stat & " NOTIFICATION"
	Set olInsp = .GetInspector
	Set wdDoc = olInsp.WordEditor
	
	with wdDoc.Range
		.Delete
		.Paste
	end with
	
	.BodyFormat = 3
	olInsp.Close(0)
	.SaveAs messagesDir & "output.msg"
	.Close(1)
	
	Set olInsp = Nothing
	Set wdDoc = Nothing 
end with

oDoc.Close(wdDoNotSaveChanges)
oWord.Quit

Set oDoc = Nothing
Set oWord = Nothing

msgFile = messagesDir & "output.msg"
Set oMsg = oOutlook.CreateItemFromTemplate(msgFile)

with oMsg
	.Display
end with

set oMsg = Nothing