'////////////////////////////////////////////////////////////
'	HTML ENCRYPTOR MADE BY
'	SARFRAZ AHMED CHANDIO
'////////////////////////////////////////////////////////////

Set FSO = CreateObject ("Scripting.FileSystemObject")
InFile = InputBox ("Enter the path of HTML file you want to encrypt","Specify HTML File")

If Len (Trim (InFile)) > 0 Then
	Set OutFile = FSO.CreateTextFile (FSO.GetParentFolderName (FSO.GetSpecialFolder (0)) & 	"Output.htm")
	OutFile.Write Encrypt (InFile)
	OutFile.Close
 	Set OutFile = Nothing
	MsgBox "The file has been saved to:" & vbNewLine & FSO.GetParentFolderName 		(FSO.GetSpecialFolder (0)) & "Output.htm",vbInformation
End If

'This function encrypts a given HTML file
Function Encrypt (InputFile)
Dim FSO
Dim File
Dim Contents

If Len (Trim (InputFile)) > 0 Then
	Set FSO = CreateObject ("Scripting.FileSystemObject")
	Set File = FSO.OpenTextFile (InputFile,1,False,0)
	Contents = File.ReadAll
	File.Close
End If

Set FSO = Nothing
Set File = Nothing

'Escape encodes it all automatically
Encrypt = "<script>" & vbNewLine & "<!--" & vbNewLine & "document.write (unescape (" & Chr (34) & Escape (Contents) & Chr (34) & "));" & vbNewLine & "//-->" & vbNewLine & "</script>"
End Function