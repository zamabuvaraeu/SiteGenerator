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

Sub ClearFilePlaceholder()
	Dim PlaceHolderFileList
	Set PlaceHolderFileList = document.getElementById("PlaceHolderFileList")
	
	Dim divFileList
	Set divFileList = document.getElementById("divFileList")
	
	PlaceHolderFileList.RemoveChild(divFileList)
	Set divFileList = Nothing
	
	Set PlaceHolderFileList = Nothing
End Sub

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

Sub CreateConfigurationSection(oConfig)
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
	Set cmdSendFile.onClick = GetRef("SendBinaryFile")
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
	cmdSendFile.value = "Отправить"
	Set cmdSendFile.onClick = GetRef("SendTextFile")
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
	CreateConfigurationSection oConfig
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
		' Set oFile = GetGenerationFileOptions(node)
		' CreateGenerationFileSection divFileList, oFile
		Set oGenerationFile = Nothing
	Next
	Set colNodes = Nothing
	
	Dim PlaceHolderFileList
	Set PlaceHolderFileList = document.getElementById("PlaceHolderFileList")
	PlaceHolderFileList.appendChild(divFileList)
	Set PlaceHolderFileList = Nothing
	Set divFileList = Nothing
	
End Sub

Sub SendFile(UrlPath, FileName, ContentType, ContentLanguage)
	' MsgBox UrlPath & vbCrLf & FileName & vbCrLf & ContentType & vbCrLf & ContentLanguage
	
	Dim wwwRoot, wwwRootPath
	wwwRoot = GetTextBoxValue("txtWwwRoot")
	wwwRootPath = FSO.BuildPath(wwwRoot, FileName)
	MsgBox wwwRootPath
	
	' Прочитать двоично файл wwwRootPath
	Dim FileData
	FileData = ReadBinaryFile(wwwRootPath)
	' MsgBox VarTypeToString(FileData)
	' MsgBox TypeName(FileData)
	
	Dim wwwUrl, wwwUrlPath
	wwwUrl = GetTextBoxValue("txtUrl")
	wwwUrlPath = FSO.BuildPath(wwwUrl, UrlPath)
	MsgBox wwwUrlPath
	
	Dim UserName
	UserName = GetTextBoxValue("txtUserName")
	
	Dim Password
	Password = GetTextBoxValue("txtPassword")
	
	Dim Request
	Set Request = CreateObject("Microsoft.XmlHttp")
	Request.Open "PUT", wwwUrlPath, False, UserName, Password
	Request.SetRequestHeader "Content-Type", ContentType
	If Len(ContentLanguage) > 0 Then
		Request.SetRequestHeader "Content-Language", ContentLanguage
	End If
	Request.send FileData
	
	MsgBox "HTTP/1.1 " & Request.Status & " " & Request.StatusText & vbCrLf &  Request.getAllResponseHeaders
	
	Set Request = Nothing
End Sub

Sub SendBinaryFile()
	Dim oFile
	Set oFile = New BinaryFileOptions
	
	oFile.UrlPath = GetSectionHiddenValue(me, 0)
	oFile.FileName = GetSectionHiddenValue(me, 1)
	oFile.ContentType = GetSectionHiddenValue(me, 2)
	oFile.ContentLanguage = GetSectionHiddenValue(me, 3)
	
	SendFile oFile.UrlPath, oFile.FileName, oFile.ContentType, oFile.ContentLanguage
	
	Set oFile = Nothing
	
End Sub

Sub SendTextFile()
	Dim oFile
	Set oFile = New TextFileOptions
	
	oFile.UrlPath = GetSectionHiddenValue(me, 0)
	oFile.FileName = GetSectionHiddenValue(me, 1)
	oFile.ContentType = GetSectionHiddenValue(me, 2)
	oFile.ContentLanguage = GetSectionHiddenValue(me, 3)
	
	' Сделать одной строкой
	Dim OneLineUtils
	OneLineUtils = GetTextBoxValue("txtOnelineFolder")
	
	Dim FullFileName, FullFileNameTxt, FullFileNameGzip, FullFileNameTxtUtf8
	FullFileName = FSO.BuildPath(GetTextBoxValue("txtWwwRoot"), oFile.FileName)
	FullFileNameTxt = FullFileName & ".txt"
	FullFileNameGzip = FullFileName & ".gz"
	FullFileNameTxtUtf8 = FullFileName & ".utf-8.txt"
	
	Dim Commanda
	Commanda = """" & OneLineUtils & """" & " " & """" & FullFileName & """"
	' MsgBox Commanda
	
	Dim RetCode
	RetCode = WshShell.Run(Commanda, 0, True)
	
	SendFile oFile.UrlPath, oFile.FileName & ".txt", oFile.ContentType, oFile.ContentLanguage
	
	' Архивировать
	Dim ArchUtils
	ArchUtils = GetTextBoxValue("txt7zipFolder")
	'"%ProgramFiles%\7-Zip\7z.exe" a %FileNameToSendGZip% %FileNameOneLineUTF8woBOM% -mx9
	Commanda = """" & ArchUtils & """" & " a " & """" & FullFileNameGzip & """" & " " & """" & FullFileNameTxtUtf8 & """ -mx9"
	' MsgBox Commanda
	
	RetCode = WshShell.Run(Commanda, 0, True)
	
	SendFile oFile.UrlPath & ".gz", oFile.FileName & ".gz", "application/x-gzip", oFile.ContentLanguage
	
	Set oFile = Nothing
	
End Sub
