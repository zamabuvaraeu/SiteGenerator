Option Explicit

Const ConfigurationFileName = "Generator.xml"

Function ReadBinaryFile(FileName)
	Dim bArr
	Dim FileStream
	Set FileStream = CreateObject("ADODB.Stream")
	FileStream.Type = 1 ' adTypeBinary
	FileStream.Open
	FileStream.LoadFromFile FileName
	bArr = FileStream.Read
	FileStream.Close
	ReadBinaryFile = bArr
	Set FileStream = Nothing
End Function

Function ReadFile(FileName)
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Dim Stream
	Set Stream = FSO.OpenTextFile(FileName, 1, 0)
	
	ReadFile = Stream.ReadAll()
	
	Set Stream = Nothing
	Set FSO = Nothing
End Function

Sub TestSub
	Dim ReadResult
	ReadResult = ReadBinaryFile("Generator.vbs")
	' ReadResult = ReadFile("Generator.vbs")
	MsgBox VarType(ReadResult)
	
	
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
	
	Dim xmlParser
	Set xmlParser = CreateObject("Msxml2.DOMDocument")
	xmlParser.async = False
	xmlParser.Load ConfigurationFileName
	
	If xmlParser.parseError.errorCode Then
		MsgBox xmlParser.parseError, "Error"
	End If
	
	' MsgBox xmlParser.xml
	Dim colNodes
	Set colNodes = xmlParser.selectSingleNode("/WebSiteGenerator/configuration/wwwRoot")
	MsgBox colNodes.text
	
	
	Set colNodes = Nothing
	Set xmlParser = Nothing			
	
	Set WshShell = Nothing
	Set FSO = Nothing
End Sub
