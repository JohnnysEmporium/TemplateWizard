Dim fname, prio, incNo, stat

fname = WScript.Arguments(0)
prio = WScript.Arguments(1)
incNo = WScript.Arguments(2)
stat = WScript.Arguments(3)

'On Error Resume Next

srcDir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")- 13)
messagesDir = srcDir & "Messages\"
docFile = srcDir & "output.docx"
msgFile = messagesDir & fname

Set oWord = CreateObject("Word.Application")

Set oDoc = oWord.Documents.Open(docFile)
oWord.visible = False
oDoc.tables(1).Range.Copy

Set oOutlook = CreateObject("Outlook.Application")
Set ns = oOutlook.session
Set targetFolder = ns.Folders("Desk.CIM@arcelormittal.com").Folders("Drafts")
Set oMsg = oOutlook.CreateItemFromTemplate(msgFile, targetFolder)

If Err.Number <> 0 Then
	WScript.Echo "Error has occured while accessing the outlook *.msg file. Check if the file isn't open in any other program and try again"
	oWord.Quit(0)
Else
	with oMsg
		.Subject = "PRIORITY " & prio & " - SNOW REF " & incNo & " - " & stat & " NOTIFICATION"
		Set olInsp = .GetInspector
		Set wdDoc = olInsp.WordEditor
		
		with wdDoc.Range
			.Delete
			.Paste
		end with
		
		.BodyFormat = 2
		olInsp.Close(0)
		.SaveAs messagesDir & "output.msg"
		.Close(1)
		
		Set olInsp = Nothing
		Set wdDoc = Nothing 
	end with

	oWord.Quit(0)

	Set oDoc = Nothing
	Set oMsg = Nothing
	Set oWord = Nothing
End If