Option Explicit

Class BinaryFileOptions
	Public UrlPath
	Public FileName
	Public ContentType
	Public ContentLanguage
End Class

Class TextFileOptions
	Public UrlPath
	Public FileName
	Public ContentType
	Public ContentLanguage
End Class

Class GenerationFileOptions
	Public UrlPath
	Public FileName
	Public ContentType
	Public ContentLanguage
	Public YamlFileName
End Class

Class ConfigurationOptions
	Public Url
	Public SourceFolder
	Public PandocFolder
	Public OneLineFolder
	Public ArchivatorFolder
	Public HttpPutFolder
	Public IpBindAddress
	Public UserName
	Public Password
End Class

Dim xmlParser
Dim FSO
Dim WshShell

Function ReadBinaryFile(FileName)
	Dim FileStream
	Set FileStream = CreateObject("ADODB.Stream")
	
	FileStream.Type = 1 ' adTypeBinary
	FileStream.Open
	FileStream.LoadFromFile FileName
	
	Dim bArr
	bArr = FileStream.Read
	
	FileStream.Close
	ReadBinaryFile = bArr
	
	Set FileStream = Nothing
End Function

Function ReadTextFile(FileName)
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Dim Stream
	Set Stream = FSO.OpenTextFile(FileName, 1, 0)
	
	ReadTextFile = Stream.ReadAll()
	
	Set Stream = Nothing
	Set FSO = Nothing
End Function

Function HttpPutData(WwwUrlPath, DataBytes, ContentType, ContentLanguage, UserName, Password)
	Dim Request
	Set Request = CreateObject("Microsoft.XmlHttp")
	
	Request.Open "PUT", WwwUrlPath, False, UserName, Password
	
	Request.SetRequestHeader "Content-Type", ContentType
	If Len(ContentLanguage) > 0 Then
		Request.SetRequestHeader "Content-Language", ContentLanguage
	End If
	
	Request.send DataBytes
	
	HttpPutData = "HTTP/1.1 " & Request.Status & " " & Request.StatusText & vbCrLf '&  Request.getAllResponseHeaders
	
	Set Request = Nothing
End Function

Sub SetTextBoxValue(TextBoxName, Value)
	Dim TextBox
	Set TextBox = document.getElementById(TextBoxName)
	TextBox.value = Value
	Set TextBox = Nothing
End Sub

Function GetTextBoxValue(TextBoxName)
	Dim TextBox
	Set TextBox = document.getElementById(TextBoxName)
	GetTextBoxValue = TextBox.value
	Set TextBox = Nothing
End Function

Function GetTextBoxValueByIndex(NodesCollection, Index)
	Dim TextBox
	Set TextBox = NodesCollection(Index)
	GetTextBoxValueByIndex = TextBox.value
	Set TextBox = Nothing
End Function

Function QuoteString(Value)
	QuoteString = """" & Value & """"
End Function

Function GetHttpPutRequestString(UrlPath)
	GetHttpPutRequestString = "PUT " & UrlPath & " HTTP/1.1" & vbCrLf
End Function

Sub ClearFilePlaceholder()
	Dim PlaceHolderFileList
	Set PlaceHolderFileList = document.getElementById("PlaceHolderFileList")
	
	Dim divFileList
	Set divFileList = document.getElementById("divFileList")
	
	PlaceHolderFileList.RemoveChild(divFileList)
	Set divFileList = Nothing
	
	Set PlaceHolderFileList = Nothing
End Sub

Sub AppendTextElement(Container, ElementName, innerText)
	Dim node
	Set node = Document.CreateElement(ElementName)
	node.innerText = innerText
	Container.appendChild(node)
	Set node = Nothing
End Sub

Function GetSectionHiddenValue(node, Index)
	GetSectionHiddenValue = GetTextBoxValueByIndex(node.parentNode.parentNode.childNodes, Index)
End Function

Function GetConfigurationFileName()
	GetConfigurationFileName = GetTextBoxValue("txtConfigurationFileName")
End Function

Function TextFileToOneLine(FullFileName)
	Dim OneLineUtils
	OneLineUtils = GetTextBoxValue("txtOnelineFolder")
	
	Dim Commanda
	Commanda = QuoteString(OneLineUtils) & " " & QuoteString(FullFileName)
	
	TextFileToOneLine = WshShell.Run(Commanda, 0, True)
	
End Function

Function ArchiveFile(FullFileNameGzip, FullFileNameTxtUtf8)
	Dim ArchUtils
	ArchUtils = GetTextBoxValue("txt7zipFolder")
	
	Dim Commanda
	Commanda = QuoteString(ArchUtils) & " a " & QuoteString(FullFileNameGzip) & " " & QuoteString(FullFileNameTxtUtf8) & " -mx9"
	
	ArchiveFile = WshShell.Run(Commanda, 0, True)
	
End Function

Function GenerateFile(YamlFullFileName)
	Dim PandocUtils
	PandocUtils = GetTextBoxValue("txtPandocFolder")
	
	Dim Commanda
	Commanda = QuoteString(PandocUtils) & " -d " & QuoteString(YamlFullFileName)
	
	GenerateFile = WshShell.Run(Commanda, 0, True)
	
End Function

Function GetConfigurationOptions(xmlParser)
	Set GetConfigurationOptions = New ConfigurationOptions
	
	Dim ConfigurationNode
	Set ConfigurationNode = xmlParser.selectSingleNode("/WebSiteGenerator/configuration")
	
	Dim ParameterNode
	For Each ParameterNode In ConfigurationNode.childNodes
		
		Select Case ParameterNode.nodeName
			Case "url"
				GetConfigurationOptions.Url = ParameterNode.Text
			Case "sourceFolder"
				GetConfigurationOptions.SourceFolder = ParameterNode.Text
			Case "pandocFolder"
				GetConfigurationOptions.PandocFolder = ParameterNode.Text
			Case "oneLineFolder"
				GetConfigurationOptions.OneLineFolder = ParameterNode.Text
			Case "archivatorFolder"
				GetConfigurationOptions.ArchivatorFolder = ParameterNode.Text
			Case "httpPutFolder"
				GetConfigurationOptions.HttpPutFolder = ParameterNode.Text
			Case "ipBindAddress"
				GetConfigurationOptions.IpBindAddress = ParameterNode.Text
			Case "userName"
				GetConfigurationOptions.UserName = ParameterNode.Text
			Case "password"
				GetConfigurationOptions.Password = ParameterNode.Text
		End Select
		
	Next
	
	Set ConfigurationNode = Nothing
	
End Function

Function GetBinaryFileOptions(xmlNode)
	Set GetBinaryFileOptions = New BinaryFileOptions
	
	Dim subNode
	For Each subNode In xmlNode.childNodes
		Select Case subNode.nodeName
			Case "urlPath"
				GetBinaryFileOptions.UrlPath = subNode.Text
			Case "file"
				GetBinaryFileOptions.FileName = subNode.Text
			Case "contentType"
				GetBinaryFileOptions.ContentType = subNode.Text
			Case "contentLanguage"
				GetBinaryFileOptions.ContentLanguage = subNode.Text
		End Select
	Next
	
End Function

Function GetTextFileOptions(xmlNode)
	Set GetTextFileOptions = New TextFileOptions
	
	Dim subNode
	For Each subNode In xmlNode.childNodes
		Select Case subNode.nodeName
			Case "urlPath"
				GetTextFileOptions.UrlPath = subNode.Text
			Case "file"
				GetTextFileOptions.FileName = subNode.Text
			Case "contentType"
				GetTextFileOptions.ContentType = subNode.Text
			Case "contentLanguage"
				GetTextFileOptions.ContentLanguage = subNode.Text
		End Select
	Next
	
End Function

Function GetGenerationFileOptions(xmlNode)
	Set GetGenerationFileOptions = New GenerationFileOptions
	
	Dim subNode
	For Each subNode In xmlNode.childNodes
		Select Case subNode.nodeName
			Case "urlPath"
				GetGenerationFileOptions.UrlPath = subNode.Text
			Case "file"
				GetGenerationFileOptions.FileName = subNode.Text
			Case "contentType"
				GetGenerationFileOptions.ContentType = subNode.Text
			Case "contentLanguage"
				GetGenerationFileOptions.ContentLanguage = subNode.Text
			Case "yaml"
				GetGenerationFileOptions.YamlFileName = subNode.Text
		End Select
	Next
	
End Function

Sub SetConfigurationTextBoxes(oConfig)
	SetTextBoxValue "txtUrl", oConfig.Url
	SetTextBoxValue "txtWwwRoot", oConfig.SourceFolder
	SetTextBoxValue "txtUserName", oConfig.UserName
	SetTextBoxValue "txtPassword", oConfig.Password
	SetTextBoxValue "txtPandocFolder", oConfig.PandocFolder
	SetTextBoxValue "txtOnelineFolder", oConfig.OneLineFolder
	SetTextBoxValue "txt7zipFolder", oConfig.ArchivatorFolder
End Sub

Sub CreateBinaryFileSection(divFileList, oFile)
	Dim Container
	Set Container = Document.CreateElement("div")
	' Container.class = "binaryfilecontainer"
	
	Dim txtHiddenUrlPath
	Set txtHiddenUrlPath = Document.CreateElement("input")
	txtHiddenUrlPath.type = "hidden"
	txtHiddenUrlPath.value = oFile.UrlPath
	Container.appendChild(txtHiddenUrlPath)
	Set txtHiddenUrlPath = Nothing
	
	Dim txtHiddenFile
	Set txtHiddenFile = Document.CreateElement("input")
	txtHiddenFile.type = "hidden"
	txtHiddenFile.value = oFile.FileName
	Container.appendChild(txtHiddenFile)
	Set txtHiddenFile = Nothing
	
	Dim txtHiddenContentType
	Set txtHiddenContentType = Document.CreateElement("input")
	txtHiddenContentType.type = "hidden"
	txtHiddenContentType.value = oFile.ContentType
	Container.appendChild(txtHiddenContentType)
	Set txtHiddenContentType = Nothing
	
	Dim txtHiddenContentLanguage
	Set txtHiddenContentLanguage = Document.CreateElement("input")
	txtHiddenContentLanguage.type = "hidden"
	txtHiddenContentLanguage.value = oFile.ContentLanguage
	Container.appendChild(txtHiddenContentLanguage)
	Set txtHiddenContentLanguage = Nothing
	
	AppendTextElement Container, "h4", oFile.UrlPath
	AppendTextElement Container, "p", "Имя файла: " & oFile.FileName
	AppendTextElement Container, "p", "Тип файла: " & oFile.ContentType
	AppendTextElement Container, "p", "Язык файла: " & oFile.ContentLanguage
	
	Dim ButtonContainer
	Set ButtonContainer = Document.CreateElement("p")
	
	Dim cmdSendFile
	Set cmdSendFile = Document.CreateElement("input")
	cmdSendFile.type = "button"
	cmdSendFile.value = "Отправить"
	Set cmdSendFile.onClick = GetRef("SendBinaryFile_Click")
	ButtonContainer.appendChild(cmdSendFile)
	Set cmdSendFile = Nothing
	
	Container.appendChild(ButtonContainer)
	Set ButtonContainer = Nothing
	
	divFileList.appendChild(Container)
	Set Container = Nothing
End Sub

Sub CreateTextFileSection(divFileList, oFile)
	Dim Container
	Set Container = Document.CreateElement("div")
	' Container.class = "binaryfilecontainer"
	
	Dim txtHiddenUrlPath
	Set txtHiddenUrlPath = Document.CreateElement("input")
	txtHiddenUrlPath.type = "hidden"
	txtHiddenUrlPath.value = oFile.UrlPath
	Container.appendChild(txtHiddenUrlPath)
	Set txtHiddenUrlPath = Nothing
	
	Dim txtHiddenFile
	Set txtHiddenFile = Document.CreateElement("input")
	txtHiddenFile.type = "hidden"
	txtHiddenFile.value = oFile.FileName
	Container.appendChild(txtHiddenFile)
	Set txtHiddenFile = Nothing
	
	Dim txtHiddenContentType
	Set txtHiddenContentType = Document.CreateElement("input")
	txtHiddenContentType.type = "hidden"
	txtHiddenContentType.value = oFile.ContentType
	Container.appendChild(txtHiddenContentType)
	Set txtHiddenContentType = Nothing
	
	Dim txtHiddenContentLanguage
	Set txtHiddenContentLanguage = Document.CreateElement("input")
	txtHiddenContentLanguage.type = "hidden"
	txtHiddenContentLanguage.value = oFile.ContentLanguage
	Container.appendChild(txtHiddenContentLanguage)
	Set txtHiddenContentLanguage = Nothing
	
	AppendTextElement Container, "h4", oFile.UrlPath
	AppendTextElement Container, "p", "Имя файла: " & oFile.FileName
	AppendTextElement Container, "p", "Тип файла: " & oFile.ContentType
	AppendTextElement Container, "p", "Язык файла: " & oFile.ContentLanguage
	
	Dim ButtonContainer
	Set ButtonContainer = Document.CreateElement("p")
	
	Dim cmdSendFile
	Set cmdSendFile = Document.CreateElement("input")
	cmdSendFile.type = "button"
	cmdSendFile.value = "Упростить и отправить"
	Set cmdSendFile.onClick = GetRef("SendTextFile_Click")
	ButtonContainer.appendChild(cmdSendFile)
	Set cmdSendFile = Nothing
	
	Container.appendChild(ButtonContainer)
	Set ButtonContainer = Nothing
	
	divFileList.appendChild(Container)
	Set Container = Nothing
End Sub

Sub CreateGenerationFileSection(divFileList, oFile)
	Dim Container
	Set Container = Document.CreateElement("div")
	' Container.class = "binaryfilecontainer"
	
	Dim txtHiddenUrlPath
	Set txtHiddenUrlPath = Document.CreateElement("input")
	txtHiddenUrlPath.type = "hidden"
	txtHiddenUrlPath.value = oFile.UrlPath
	Container.appendChild(txtHiddenUrlPath)
	Set txtHiddenUrlPath = Nothing
	
	Dim txtHiddenFile
	Set txtHiddenFile = Document.CreateElement("input")
	txtHiddenFile.type = "hidden"
	txtHiddenFile.value = oFile.FileName
	Container.appendChild(txtHiddenFile)
	Set txtHiddenFile = Nothing
	
	Dim txtHiddenContentType
	Set txtHiddenContentType = Document.CreateElement("input")
	txtHiddenContentType.type = "hidden"
	txtHiddenContentType.value = oFile.ContentType
	Container.appendChild(txtHiddenContentType)
	Set txtHiddenContentType = Nothing
	
	Dim txtHiddenContentLanguage
	Set txtHiddenContentLanguage = Document.CreateElement("input")
	txtHiddenContentLanguage.type = "hidden"
	txtHiddenContentLanguage.value = oFile.ContentLanguage
	Container.appendChild(txtHiddenContentLanguage)
	Set txtHiddenContentLanguage = Nothing
	
	Dim txtHiddenYamlFileName
	Set txtHiddenYamlFileName = Document.CreateElement("input")
	txtHiddenYamlFileName.type = "hidden"
	txtHiddenYamlFileName.value = oFile.YamlFileName
	Container.appendChild(txtHiddenYamlFileName)
	Set txtHiddenYamlFileName = Nothing
	
	AppendTextElement Container, "h4", oFile.UrlPath
	AppendTextElement Container, "p", "Имя файла: " & oFile.FileName
	AppendTextElement Container, "p", "Тип файла: " & oFile.ContentType
	AppendTextElement Container, "p", "Язык файла: " & oFile.ContentLanguage
	AppendTextElement Container, "p", "YamlFileName: " & oFile.YamlFileName
	
	Dim ButtonContainer
	Set ButtonContainer = Document.CreateElement("p")
	
	Dim cmdSendFile
	Set cmdSendFile = Document.CreateElement("input")
	cmdSendFile.type = "button"
	cmdSendFile.value = "Генерировать, упростить и отправить"
	Set cmdSendFile.onClick = GetRef("SendGenerationFile_Click")
	ButtonContainer.appendChild(cmdSendFile)
	Set cmdSendFile = Nothing
	
	Container.appendChild(ButtonContainer)
	Set ButtonContainer = Nothing
	
	divFileList.appendChild(Container)
	Set Container = Nothing
End Sub

Sub BodyLoad()
	Set xmlParser = CreateObject("Msxml2.DOMDocument")
	xmlParser.async = False
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
End Sub

Sub LoadConfiguration()
	xmlParser.Load GetConfigurationFileName()
	
	Dim oConfig
	Set oConfig = GetConfigurationOptions(xmlParser)
	SetConfigurationTextBoxes oConfig
	Set oConfig = Nothing
	
	ClearFilePlaceholder
	
	Dim divFileList
	Set divFileList = Document.CreateElement("div")
	divFileList.id = "divFileList"
	
	Dim colNodes, node
	
	AppendTextElement divFileList, "h3", "Двоичные файлы"
	Set colNodes = xmlParser.selectNodes("/WebSiteGenerator/binaryFiles")
	For Each node In colNodes
		Dim oBinaryFile
		Set oBinaryFile = GetBinaryFileOptions(node)
		CreateBinaryFileSection divFileList, oBinaryFile
		Set oBinaryFile = Nothing
	Next
	Set colNodes = Nothing
	
	AppendTextElement divFileList, "h3", "Текстовые файлы"
	Set colNodes = xmlParser.selectNodes("/WebSiteGenerator/textFiles")
	For Each node In colNodes
		Dim oTextFile
		Set oTextFile = GetTextFileOptions(node)
		CreateTextFileSection divFileList, oTextFile
		Set oTextFile = Nothing
	Next
	Set colNodes = Nothing
	
	AppendTextElement divFileList, "h3", "Генерируемые файлы"
	Set colNodes = xmlParser.selectNodes("/WebSiteGenerator/generationFiles")
	For Each node In colNodes
		Dim oGenerationFile
		Set oGenerationFile = GetGenerationFileOptions(node)
		CreateGenerationFileSection divFileList, oGenerationFile
		Set oGenerationFile = Nothing
	Next
	Set colNodes = Nothing
	
	Dim PlaceHolderFileList
	Set PlaceHolderFileList = document.getElementById("PlaceHolderFileList")
	PlaceHolderFileList.appendChild(divFileList)
	Set PlaceHolderFileList = Nothing
	Set divFileList = Nothing
	
End Sub

Function SendFile(UrlPath, FileName, ContentType, ContentLanguage)
	Dim SourceFolder, FullFileName
	SourceFolder = GetTextBoxValue("txtWwwRoot")
	FullFileName = FSO.BuildPath(SourceFolder, FileName)
	
	Dim wwwUrl, wwwUrlPath
	wwwUrl = GetTextBoxValue("txtUrl")
	wwwUrlPath = FSO.BuildPath(wwwUrl, UrlPath)
	
	Dim UserName
	UserName = GetTextBoxValue("txtUserName")
	
	Dim Password
	Password = GetTextBoxValue("txtPassword")
	
	Dim FileData
	FileData = ReadBinaryFile(FullFileName)
	
	SendFile = HttpPutData(wwwUrlPath, FileData, ContentType, ContentLanguage, UserName, Password)
	
End Function

Sub SendAllBinaryFiles_Click()
	Dim colNodes, node
	
	Dim Output
	Output = ""
	
	Set colNodes = xmlParser.selectNodes("/WebSiteGenerator/binaryFiles")
	For Each node In colNodes
		Dim oFile
		Set oFile = GetBinaryFileOptions(node)
		
		Output = Output & GetHttpPutRequestString(oFile.UrlPath)
		Output = Output & SendFile(oFile.UrlPath, oFile.FileName, oFile.ContentType, oFile.ContentLanguage)
		
		Set oFile = Nothing
	Next
	Set colNodes = Nothing
	
	SetTextBoxValue "txtOutput", Output

End Sub

Sub SendBinaryFile_Click()
	Dim oFile
	Set oFile = New BinaryFileOptions
	
	oFile.UrlPath = GetSectionHiddenValue(me, 0)
	oFile.FileName = GetSectionHiddenValue(me, 1)
	oFile.ContentType = GetSectionHiddenValue(me, 2)
	oFile.ContentLanguage = GetSectionHiddenValue(me, 3)
	
	Dim Output
	Output = GetHttpPutRequestString(oFile)
	
	Output = Output & SendFile(oFile.UrlPath, oFile.FileName, oFile.ContentType, oFile.ContentLanguage)
	MsgBox Output
	
	Set oFile = Nothing
	
End Sub

Sub SendTextFile_Click()
	Dim oFile
	Set oFile = New TextFileOptions
	
	oFile.UrlPath = GetSectionHiddenValue(me, 0)
	oFile.FileName = GetSectionHiddenValue(me, 1)
	oFile.ContentType = GetSectionHiddenValue(me, 2)
	oFile.ContentLanguage = GetSectionHiddenValue(me, 3)
	
	Dim FullFileName, FullFileNameTxt, FullFileNameGzip, FullFileNameTxtUtf8
	FullFileName = FSO.BuildPath(GetTextBoxValue("txtWwwRoot"), oFile.FileName)
	FullFileNameTxt = FullFileName & ".txt"
	FullFileNameTxtUtf8 = FullFileName & ".utf-8.txt"
	FullFileNameGzip = FullFileName & ".gz"
	
	Dim Output
	Output = "Минификатор файла завершился кодом: " & CStr(TextFileToOneLine(FullFileName)) & vbCrLf
	
	Output = Output & GetHttpPutRequestString(oFile.UrlPath)
	Output = Output & SendFile(oFile.UrlPath, oFile.FileName & ".txt", oFile.ContentType, oFile.ContentLanguage)
	
	Output = Output & "Архиватор файла завершился кодом: " & CStr(ArchiveFile(FullFileNameGzip, FullFileNameTxtUtf8)) & vbCrLf
	
	Output = Output & GetHttpPutRequestString(oFile.UrlPath & ".gz")
	Output = Output & SendFile(oFile.UrlPath & ".gz", oFile.FileName & ".gz", "application/x-gzip", oFile.ContentLanguage)
	
	Set oFile = Nothing
	
	MsgBox Output
	
	FSO.DeleteFile FullFileNameTxt
	FSO.DeleteFile FullFileNameGzip
	FSO.DeleteFile FullFileNameTxtUtf8
	
End Sub

Sub SendGenerationFile_Click()
	Dim oFile
	Set oFile = New GenerationFileOptions
	
	oFile.UrlPath = GetSectionHiddenValue(me, 0)
	oFile.FileName = GetSectionHiddenValue(me, 1)
	oFile.ContentType = GetSectionHiddenValue(me, 2)
	oFile.ContentLanguage = GetSectionHiddenValue(me, 3)
	oFile.YamlFileName = GetSectionHiddenValue(me, 4)
	
	Dim FullFileName, FullYamlFileName, FullFileNameTxt, FullFileNameGzip, FullFileNameTxtUtf8
	FullFileName = FSO.BuildPath(GetTextBoxValue("txtWwwRoot"), oFile.FileName)
	FullYamlFileName = FSO.BuildPath(GetTextBoxValue("txtWwwRoot"), oFile.YamlFileName)
	FullFileNameTxt = FullFileName & ".txt"
	FullFileNameTxtUtf8 = FullFileName & ".utf-8.txt"
	FullFileNameGzip = FullFileName & ".gz"
	
	Dim OldCurrentDirectory
	OldCurrentDirectory = WshShell.CurrentDirectory
	
	Dim Output
	WshShell.CurrentDirectory = GetTextBoxValue("txtWwwRoot")
	Output = "Генератор файла завершился кодом: " & CStr(GenerateFile(oFile.YamlFileName)) & vbCrLf
	WshShell.CurrentDirectory = OldCurrentDirectory
	
	Output = Output & "Минификатор файла завершился кодом: " & CStr(TextFileToOneLine(FullFileName)) & vbCrLf
	Output = Output & GetHttpPutRequestString(oFile.UrlPath)
	Output = Output & SendFile(oFile.UrlPath, oFile.FileName & ".txt", oFile.ContentType, oFile.ContentLanguage)
	
	Output = Output & "Архиватор файла завершился кодом: " & CStr(ArchiveFile(FullFileNameGzip, FullFileNameTxtUtf8)) & vbCrLf
	
	Output = Output & GetHttpPutRequestString(oFile.UrlPath & ".gz")
	Output = Output & SendFile(oFile.UrlPath & ".gz", oFile.FileName & ".gz", "application/x-gzip", oFile.ContentLanguage)
	
	Set oFile = Nothing
	
	MsgBox Output
	
	FSO.DeleteFile FullFileNameTxt
	FSO.DeleteFile FullFileNameGzip
	FSO.DeleteFile FullFileNameTxtUtf8
	
End Sub
