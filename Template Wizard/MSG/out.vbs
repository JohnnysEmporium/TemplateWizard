Dim fname, prio, incNo, stat

fname = WScript.Arguments(0)
prio = WScript.Arguments(1)
incNo = WScript.Arguments(2)
stat = WScript.Arguments(3)

currentdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
srcdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")-4)
docFile = srcdir & "output.docx"
msgFile = currentdir & fname
WScript.Echo docFile

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
end with

oDoc.Close(wdDoNotSaveChanges)
oWord.Quit

oMsg.Close olSave
oMsg.Display