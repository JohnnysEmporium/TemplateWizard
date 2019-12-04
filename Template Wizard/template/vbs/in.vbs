Dim msgFile

msgFile = WScript.Arguments(0)

On Error Resume Next

srcdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\") - 14)
currentdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
msgFile = "\Messages\" & msgFile

Set oWord = CreateObject("Word.Application")
Set oDoc = oWord.Documents.Add()
oWord.visible = False

Set oOutlook = CreateObject("Outlook.Application")
Set oMsg = oOutlook.CreateItemFromTemplate(srcdir & msgFile)

If Err.Number <> 0 Then
	WScript.Echo "Error has occured while accessing the outlook *.msg file. Check if the file isn't open in any other program and try again"
	oDoc.Display
Else
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
	oWord.Quit(0)
End If 