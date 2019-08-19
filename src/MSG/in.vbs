Dim msgFile

currentdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
msgFile = WScript.Arguments(0)

Set oWord = CreateObject("Word.Application")
Set oDoc = oWord.Documents.Add()
oWord.visible = False

Set oOutlook = CreateObject("Outlook.Application")
Set oMsg = oOutlook.CreateItemFromTemplate(currentdir & msgFile)

with oMsg
	Set olInsp = .GetInspector
	Set wdDoc = olInsp.WordEditor
	wdDoc.tables(1).Range.Copy
end with
with oDoc.Range
	.Paste
end with	

oMsg.Close olDiscard
oMsg.Delete
oDoc.SaveAS(currentdir & "temp.docx")
oWord.Quit












