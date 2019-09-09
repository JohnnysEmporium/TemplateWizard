Dim msgFile

srcdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\") - 5)
currentdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
msgFile = "\Messages\" & "P3_INC1467340.msg"

Set oWord = CreateObject("Word.Application")
Set oDoc = oWord.Documents.Add()
oWord.visible = False

Set oOutlook = CreateObject("Outlook.Application")
Set oMsg = oOutlook.CreateItemFromTemplate(srcdir & msgFile)

with oMsg
	Set olInsp = .GetInspector
	Set wdDoc = olInsp.WordEditor
	wdDoc.tables(wdDoc.tables.Count).Range.Copy
end with
with oDoc.Range
	.Paste
end with	

oMsg.Close olDiscard
oMsg.Delete
oDoc.SaveAS(currentdir & "temp.docx")
oWord.Quit












