<%
Option Explicit
'On Error Resume Next

If Request.QueryString("upload") = "" Then
	Session.CodePage = 65001
Else
	Session.CodePage = 1252
End If

''
' Scripts name
''
Dim arPath, strScript
arPath = Split(Request.ServerVariables("SCRIPT_NAME"), "/")
strScript = arPath(Ubound(arPath))

''
' List of encodings for file editting
'
' ({@link http://msdn.microsoft.com/en-us/library/ms526296%28v=exchg.10%29.aspx Source})
''
Dim arEncodings
arEncodings = Array( _
	"ISO-8859-1", _
	"BIG5", _
	"EUC-JP", _
	"EUC-KR", _
	"GB2312", _
	"ISO-2022-JP", _
	"ISO-2022-KR", _
	"ISO-8859-2", _
	"ISO-8859-3", _
	"ISO-8859-4", _
	"ISO-8859-5", _
	"ISO-8859-6", _
	"ISO-8859-7", _
	"ISO-8859-8", _
	"ISO-8859-9", _
	"KOI8-R", _
	"SHIFT-JIS", _
	"US-ASCII", _
	"UTF-8", _
	"UNICODE" _
)

''
' File and folder attributes collection
''
Dim dAttributes
Set dAttributes = Server.CreateObject("Scripting.Dictionary")
dAttributes.Add "n", Array(0, "Normal", False)
dAttributes.Add "r", Array(1, "Read Only", True)
dAttributes.Add "h", Array(2, "Hidden", True)
dAttributes.Add "s", Array(4, "System", True)
dAttributes.Add "v", Array(8, "Volume", False)
dAttributes.Add "f", Array(16, "Directory", False)
dAttributes.Add "a", Array(32, "Archive", True)
dAttributes.Add "l", Array(1024, "Alias", False)
dAttributes.Add "c", Array(2048, "Compressed", False)

''
' Processes file for downloading
''
If Not Request.QueryString("download") = "" Then
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(Request.QueryString("download")) Then
		Set objFile = objFSO.GetFile(Request.QueryString("download"))

		' ({@link http://nolovelust.com/post/classic-asp-large-file-download-code Source})
		Dim intChunkSize, objStream, intStreamSize
		intChunkSize = 2048
		Server.ScriptTimeout = 900

		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open()
		objStream.Type = 1
		objStream.LoadFromFile objFile.Path
		intStreamSize = objStream.Size

		Response.ContentType = "application/octet-stream"
		'Response.AddHeader "Content-Length", intStreamSize
		Response.AddHeader "Content-Disposition", "attachment;filename=""" & objFile.Name & """;"
		Response.Buffer = False

		For i = 1 To intStreamSize \ intChunkSize
			If Not Response.IsClientConnected Then Exit For
			Response.BinaryWrite objStream.Read(intChunkSize)
		Next

		If intStreamSize Mod intChunkSize > 0 Then
			If Response.IsClientConnected Then
				Response.BinaryWrite objStream.Read(intStreamSize Mod intChunkSize)
			End If
		End If

		objStream.Close
		Set objStream = Nothing
	Else
		Response.Status = "404 Not Found"
		Response.Write "File Not Found"
	End If

	Response.End
End If

''
' Processes file for downloading
''
If Not Request.QueryString("list") = "" Then

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	objStartFolder = Request.QueryString("list")
	strFile = ""

	If Request.QueryString("level") = "" Then
		intMaxLevel = -1
	Else
		intMaxLevel = Int(Request.QueryString("level"))
	End If

	Response.Buffer = False
	Response.ContentType = "text/plain; charset=""UTF-8"""

	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		Response.Write vbCRLF & objFolder.Path & "\\" & objFile.Name
	Next

	ShowSubfolders objFSO.GetFolder(objStartFolder), 0

	Response.End
End If
%>
<!DOCTYPE html>
<head>
	<title>ASP File Browser</title>
	<meta http-equiv='Content-Type' content='text/html;charset=utf-8' />
	<style>
		body, input, select, table {font-size: 13px; font-family: Courier New; white-space: nowrap;}
		table td, table th {padding: 5px;}
		table tr:nth-child(even) {background: #F0F0F0;}
		table tr:nth-child(odd) {background: #FFFFFF;}
	</style>
</head>
<body>
<%
''
'
' FILE UPLOADING
'
''

If Not Request.QueryString("upload") = "" Then
	Dim strDestination
	strDestination = Request.QueryString("upload")

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

		Dim UploadRequest
		Dim byteCount, RequestBin
		Dim sFullFilePath, sPathEnd
		Dim sContentType, sFilePathName, sFileName, sValue
		Dim oFile, oFSO
		Dim i

		Response.Expires = 0
		Response.Buffer = TRUE

		byteCount = Request.TotalBytes
		RequestBin = Request.BinaryRead(byteCount)
		Set UploadRequest = Server.CreateObject("Scripting.Dictionary")

		BuildUploadRequest RequestBin

		' This will place the uploaded file into the root directory of the web site -
		' Modify this path as needed.
		If Not Right(strDestination, 1) = "\" Then
			strDestination = strDestination & "\"
		End If

		sContentType = UploadRequest.Item("blob").Item("ContentType")
		sFilePathName = UploadRequest.Item("blob").Item("FileName")
		sFileName = Right(sFilePathName, Len(sFilePathName) - InstrRev(sFilePathName, "\"))
		sValue = UploadRequest.Item("blob").Item("Value")

		sFullFilePath = strDestination & sFileName

		'Create FileSytemObject Component
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")

		'Create and Write to a File
		sPathEnd = Len(Server.mappath(Request.ServerVariables("PATH_INFO"))) - 14

		Set oFile = oFSO.CreateTextFile(sFullFilePath, True)

		For i = 1 to LenB(sValue)
			oFile.Write Chr(AscB(MidB(sValue,i,1)))
		Next

		oFile.Close

		Set oFile = Nothing
		Set oFSO = Nothing

		With Response
			.Write("<b>Uploaded File: </b>" & sFullFilePath & "<BR>")
			.Write("<b>Content Type: </b>" & sContentType & "<BR>")
		End With

		Set UploadRequest = Nothing
	End If

%>
	<form method="post" enctype="multipart/form-data" action="">
		<p>Select File : <input type="file" name="blob"></p>
		<p><input type="submit" name="btnSubmit" value="Upload"></p>
	</form>
<%
''
'
' FILE/FOLDER'S ATTRIBUTES
'
''

ElseIf Not Request.QueryString("attributes") = "" Then

	Dim objAttributes
	Dim objItem
	Dim strItem, strAttribute, colKeys, strKey

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	strItem = Trim(Request.QueryString("attributes"))
	If Right(strItem, 1) = "\" Then
		Set objItem = objFSO.GetFolder(strItem)
	Else
		Set objItem = objFSO.GetFile(strItem)
	End If

	strAttribute = fsAttributes(objItem.Attributes)
	colKeys = dAttributes.Keys

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

		For Each strKey In colKeys
			If dAttributes.Item(strKey)(2) = True Then
				If Not Request.Form("attribute_" & strKey) = "" Then
					If InStr(strAttribute, strKey) = 0 Then
						objItem.Attributes = objItem.Attributes + dAttributes.Item(strKey)(0)
					End If
				Else
					If InStr(strAttribute, strKey) > 0 Then
						objItem.Attributes = objItem.Attributes - dAttributes.Item(strKey)(0)
					End If
				End If
			End If
		Next

		If Not Request.Form("date") = "" Then
			fileDateLastModified strItem, Request.Form("date")
		End If

		strAttribute = fsAttributes(objItem.Attributes)
	End If

%>
	<form method='post' action=''>
		<table border='1'>
			<tr>
				<td rowspan='4'><strong>Attributes</strong></td>
<%
	For Each strKey In colKeys
		If dAttributes.Item(strKey)(2) = True Then
			If InStr(strAttribute, strKey) > 0 Then
				Response.Write Tab(4) & "<td style='text-align: right;'><input type='checkbox' name='attribute_" & strKey & "' checked='checked' value='" & strKey & "'></td>" & vbCRLF
				Response.Write Tab(4) & "<td>" & dAttributes.Item(strKey)(1) & "</td>" & vbCRLF
			Else
				Response.Write Tab(4) & "<td style='text-align: right;'><input type='checkbox' name='attribute_" & strKey & "' value='" & strKey & "'></td>" & vbCRLF
				Response.Write Tab(4) & "<td>" & dAttributes.Item(strKey)(1) & "</td>" & vbCRLF
			End If
			Response.Write Tab(3) & "</tr>" & vbCRLF
			Response.Write Tab(3) & "<tr>" & vbCRLF
		End If
	Next

%>
				<td>
					<strong>Last Modified Date</strong>
				</td>
				<td colspan='2'>
					<input name='date' size='30' value='<%=objItem.DateLastModified%>'>
				</td>
			</tr>
			<tr>
				<td colspan='3'>
					<input type='submit' value='Change'>
				</td>
			</tr>
		</table>
	</form>
<%

''
'
' FILE/FOLDER'S PROPERTIES
'
''

ElseIf Not Request.QueryString("properties") = "" Then

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	strItem = Trim(Request.QueryString("properties"))

	If Right(strItem, 1) = "\" Then
		Set objItem = objFSO.GetFolder(strItem)
	Else
		Set objItem = objFSO.GetFile(strItem)
	End If

	Dim strAttributeName
	strAttributeName = ""
	strAttribute = fsAttributes(objItem.Attributes)
	colKeys = dAttributes.Keys

	Dim dProperties
	Set dProperties = Server.CreateObject("Scripting.Dictionary")
	dProperties.Add "Name", objItem.Name
	dProperties.Add "Full Path", objItem.Path
	dProperties.Add "Size", convertSize(objItem.Size)
	dProperties.Add "Size (Bytes)", objItem.Size
	dProperties.Add "Type", objItem.Type
	dProperties.Add "Date Created", objItem.DateCreated
	dProperties.Add "Date Last Accessed", objItem.DateLastAccessed
	dProperties.Add "Date Last Modified", objItem.DateLastModified

	For Each strKey In colKeys
		If InStr(strAttribute, strKey) > 0 Then
			strAttributeName = strAttributeName & dAttributes.Item(strKey)(1) & " - "
		End If
	Next
	dProperties.Add "Attributes", strAttributeName

	dProperties.Add "Short Name", objItem.ShortName
	dProperties.Add "Short Path", objItem.ShortPath
	dProperties.Add "Parent Folder", objItem.ParentFolder
	dProperties.Add "Drive", objItem.Drive
	
%>
	<table border='1'>

<%
	colKeys = dProperties.Keys
	For Each strKey In colKeys
		Response.Write Tab(2) & "<tr>" & vbCRLF
		Response.Write Tab(3) & "<td style='font-weight: bolder; text-align: right;'>" & strKey & "</td>" & vbCRLF
		Response.Write Tab(3) & "<td>" & dProperties.Item(strKey) & "</td>" & vbCRLF
		Response.Write Tab(2) & "</tr>" & vbCRLF
	Next
%>

	</table>
<%

''
'
' FILE EDITTING
'
''

ElseIf Not Request.QueryString("edit") = "" Then
	Dim arSearch, strEncoding, strData, strCurrentEncoding
	arSearch = Filter(arEncodings, Request.QueryString("encoding"))
	If Ubound(arSearch) = 0 Then
		strEncoding = Request.QueryString("encoding")
	Else
		strEncoding = arEncodings(0)
	End If

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		fileWriteText Request.QueryString("edit"), Request.Form("contents"), strEncoding
	End If

	strData = strConvertHTML(fileReadText(Request.QueryString("edit"), strEncoding))

	If Err.Number = 0 Then
%>
	<form method='post' action=''>
		<input type='text' name='name' value='<%=Request.QueryString("edit")%>' size='50'>
		<span>Change File Encoding</span>
		<select onchange='this.options[this.selectedIndex].value && (window.location = scriptName() + "?edit=" + document.getElementsByName("name")[0].value + "&encoding=" + this.options[this.selectedIndex].value);'>
<%
		For Each strCurrentEncoding In arEncodings
			If strCurrentEncoding = strEncoding Then
				Response.Write Tab(3) & "<option value='" & strCurrentEncoding & "' selected='selected'>" & strCurrentEncoding & "</option>" & vbCRLF
			Else
				Response.Write Tab(3) & "<option value='" & strCurrentEncoding & "'>" & strCurrentEncoding & "</option>" & vbCRLF
			End If
		Next
%>
		</select>
		<div style="margin:5px 0;">
			<textarea style='width:100%;height:80%' name='contents'><%=strData%></textarea>
		</div>
		<input type='submit'>
	</form>
<%
	End If

''
'
' SERVER VARIABLES
'
''

ElseIf Request.QueryString("server") = "variables" Then

	Dim strVariable
	Response.Write Tab(1) & "<table border='1'>" & vbCRLF

	For Each i In Request.ServerVariables
		strVariable = Replace(Request.ServerVariables(i), vbLF, "<br />")
		strVariable = Replace(strVariable, vbCR, "")

		Response.Write Tab(2) & "<tr>" & vbCRLF
		Response.Write Tab(3) & "<td><strong>" & i & "</strong></td>" & vbCRLF
		Response.Write Tab(3) & "<td>" & strVariable & "</td>" & vbCRLF
		Response.Write Tab(2) & "</tr>" & vbCRLF
	Next

	Response.Write Tab(1) & "</table>" & vbCRLF

''
'
' FILE BROWSING
'
''

Else
	Dim strFolder
	Dim objFSO, objFolder
	If Request.QueryString("browse") = "" Then
		strFolder = Request.ServerVariables("APPL_PHYSICAL_PATH")
		If Len(strFolder) = 0 Then strFolder = "."
	Else
		strFolder = Trim(CStr(Request.QueryString("browse")))
	End If

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(strFolder)

	If Err.Number = 0 Then
	%>
	<form method='post' action='' name='form'>
		<table border='1'>
			<tr>
				<th><input type='checkbox' onclick='toggle(this)' /></th>
				<th>Type</th>
				<th>Name</th>
				<th>Size</th>
				<th>Date Created</th>
				<th>Date Last Accessed</th>
				<th>Date Last Modified</th>
				<th>Attributes</th>
			</tr>

<%
		If Not Request.Form("create") = "" Then
			Dim strItemName
			strItemName = Trim(Request.Form("name"))
			If Not strItemName = "" Then
				fsCreate Request.Form("cwd_") & "\" & strValidFilename(strItemName), Request.Form("new")
			End If
		End If

		If Not Request.Form("do_action") = "" Then
			If Request.Form("items").Count > 0 Then
				For Each i In Request.Form("items")
					Select Case Request.Form("action")
						Case "delete"
							fsDelete Request.Form("cwd_") & "\" & i
						Case "copy"
							fsCopy i, i, Request.Form("cwd_"), Request.Form("action_"), False
						Case "copyo"
							fsCopy i, i, Request.Form("cwd_"), Request.Form("action_"), True
						Case "move"
							fsMove i, i, Request.Form("cwd_"), Request.Form("action_")
						Case "rename"
							fsRename i, Request.Form("action_"), Request.Form("cwd_")
						Case "zip"
							Dim strZipFile
							strZipFile = Left(i, Len(i) - 1) & ".zip"
							zipAdd Request.Form("action_") & "\" & strZipFile, Request.Form("cwd_") & "\" & i
						Case "unzip"
							Dim strExtractFolder
							strExtractFolder = Left(i, InStrRev(i, ".") - 1)
							zipExtract Request.Form("cwd_") & "\" & i, Request.Form("action_") & "\" & strExtractFolder
					End Select
				Next
			End If
		End If

		Dim colFiles, colSubfolders
		Dim strCWD, strParent
		Dim objSubFolder, objFile
		Dim objDrive, strDriveType

		Set colFiles = objFolder.Files
		Set colSubfolders = objFolder.SubFolders
		strCWD = objFolder.Path

		Set strParent = objFolder.ParentFolder
		If Not strParent Is Nothing Then
			strParent = CStr(strParent)
			With Response
				.Write Tab(3) & "<tr>" & vbCRLF
				.Write Tab(4) & "<td>&nbsp;</td>" & vbCRLF
				.Write Tab(4) & "<td>&nbsp;</td>" & vbCRLF
				.Write Tab(4) & "<td>" & vbCRLF
				.Write Tab(5) & "<a href='" & strScript & "?browse=" & strParent & "'>..</a>" & vbCRLF
				.Write Tab(4) & "</td>" & vbCRLF
				.Write Tab(4) & "<td colspan='5'>&nbsp;</td>" & vbCRLF
				.Write Tab(3) & "</tr>" & vbCRLF
			End With
		End If

		If colSubfolders.Count > 0 Then
			For Each objSubFolder In colSubfolders
				With Response
					.Write Tab(3) & "<tr>" & vbCRLF
					.Write Tab(4) & "<td><input type='checkbox' name='items' value='" & objSubFolder.Name & "\'></td>" & vbCRLF
					.Write Tab(4) & "<td>[" & UCase(objSubFolder.Type) & "]</td>" & vbCRLF
					.Write Tab(4) & "<td>" & vbCRLF
					.Write Tab(5) & "<a href='" & strScript & "?browse=" & objSubFolder.Path & "\'>" & objSubFolder.Name & "\</a>" & vbCRLF
					.Write Tab(4) & "</td>" & vbCRLF
				End With

				objAttributes = objSubFolder.Attributes
				Err.Clear
				If Err.Number <> 0 Then
					Response.Write Tab(4) & "<td colspan='4'>&nbsp;</td>" & vbCRLF
				Else
					'Response.Write Tab(4) & "<td>" & convertSize(objSubFolder.Size) & "</td>" & vbCRLF
					Response.Write Tab(4) & "<td>-</td>" & vbCRLF
					Response.Write Tab(4) & "<td>" & CStr(objSubFolder.DateCreated) & "</td>" & vbCRLF
					Response.Write Tab(4) & "<td>" & CStr(objSubFolder.DateLastAccessed) & "</td>" & vbCRLF
					Response.Write Tab(4) & "<td>" & CStr(objSubFolder.DateLastModified) & "</td>" & vbCRLF
					Response.Write Tab(4) & "<td>" & fsAttributes(objAttributes) & "</td>" & vbCRLF
				End If

				Response.Write Tab(3) & "</tr>" & vbCRLF
			Next
		End If

		If colFiles.Count > 0 Then
			For Each objFile In colFiles
				Response.Write Tab(3) & "<tr>" & vbCRLF
				Response.Write Tab(4) & "<td><input type='checkbox' name='items' value='" & objFile.Name & "'></td>" & vbCRLF
				Response.Write Tab(4) & "<td>[" & UCase(objFile.Type) & "]</td>" & vbCRLF
				Response.Write Tab(4) & "<td>" & objFile.Name & "</td>" & vbCRLF

				objAttributes = objFile.Attributes
				Err.Clear
				If Err.Number <> 0 Then
					Response.Write Tab(4) & "<td colspan='4'>&nbsp;</td>" & vbCRLF
				Else
					With Response
						.Write Tab(4) & "<td>" & convertSize(objFile.Size) & "</td>" & vbCRLF
						.Write Tab(4) & "<td>" & CStr(objFile.DateCreated) & "</td>" & vbCRLF
						.Write Tab(4) & "<td>" & CStr(objFile.DateLastAccessed) & "</td>" & vbCRLF
						.Write Tab(4) & "<td>" & CStr(objFile.DateLastModified) & "</td>" & vbCRLF
						.Write Tab(4) & "<td>" & fsAttributes(objAttributes) & "</td>" & vbCRLF
					End With
				End If

				Response.Write Tab(3) & "</tr>" & vbCRLF
			Next
		End If
%>
			<tr>
				<td><input type='checkbox' onclick='toggle(this)' /></td>
				<td colspan='7' style='text-align: right;'>Showing <%=colFiles.Count%> files &amp; <%=colSubfolders.Count%> subfolders</td>
			</tr>
			<tr>
				<td colspan='8'><span>Selected File(s) / Folder(s)</span>
					<select name='action'>
						<option value=''>-- Select an Action --</option>
						<option value='attributes'>Attributes</option>
						<option value='copy'>Copy</option>
						<option value='copyo'>Copy (Overwrite)</option>
						<option value='edit'>Edit</option>
						<option value='delete'>Delete</option>
						<option value='download'>Download</option>
						<option value='move'>Move</option>
						<option value='properties'>Properties</option>
						<option value='rename'>Rename</option>
						<option value='unzip'>Unzip</option>
						<option value='zip'>Zip (Folder)</option>
					</select>
					<input type='hidden' name='action_' value=''>
					<input type='submit' name='do_action' value='Submit' onclick='return formSubmit();'>
				</td>
			</tr>
			<tr>
				<td colspan='8'>
					<span>Enter Name</span>
					<input type='text' name='name' value=''>
					<input type='radio' name='new' value='file'> File
					<input type='radio' name='new' value='folder'> Folder
					<input type='submit' name='create' value='Create New'> or
					<input type='button' onclick='window.open(scriptName() + "?upload=" + encodeURIComponent(document.getElementsByName("cwd_")[0].value))' value='Upload File'>
				</td>
			</tr>
			<tr>
				<td colspan='8'><span>Current Working Directory</span>
				<input type='text' name='cwd' value='<%=strCWD%>'>
				<input type='hidden' name='cwd_' value='<%=strCWD%>'>
				<input type='button' value='Change' onclick='chdir()'></td>
			</tr>
			<tr>
				<td colspan='8'>
					<span>Change Drive</span>
					<select onchange='this.options[this.selectedIndex].value && (window.location = this.options[this.selectedIndex].value);'>
						<option>-- Select a Drive --</option>

<%
		For Each objDrive in objFSO.Drives
			Select Case objDrive.DriveType
				Case 1
					strDriveType = "No Root Directory"
				Case 2
					strDriveType = "Removable Drive"
				Case 3
					strDriveType = "Local Hard Disk"
				Case 4
					strDriveType = "Network Disk"	
				Case 5
					strDriveType = "Compact Disk"	
				Case 6
					strDriveType = "RAM Disk"
				Case Else
					strDriveType = "Unknown"
			End Select
			Response.Write Tab(6) & "<option value='" & strScript & "?browse=" & objDrive.DriveLetter & ":\'>[" & UCase(strDriveType) & "] " & objDrive.DriveLetter & ":\</option>" & vbCRLF
		Next
%>
					</select>
					<span>(<a href='#' onclick='window.open(scriptName() + "?server=variables");'>Server Variables</a>)</span>
				</td>
			</tr>
		</table>
	</form>
<%
	End If
End If

If Err.Number <> 0 Then
	Response.Write "<span>Error #: " & CStr(Err.Number) & "<br />" & vbcrLF
	Response.Write "Description: " & Err.Description & "<br />" & vbcrLF
	Response.Write "Source: " & Err.Source & "</span><br />" & vbCRLF
End If
%>
	<script language='JavaScript'>
		/*
		* Gets script's name
		*
		* @link http://stackoverflow.com/questions/2196606/getting-the-current-script-executing-filename-in-javascript Source
		* @return Returns executing script's name
		*/
		function scriptName()
		{
			var url = window.location.pathname;
			var lastUri = url.substring(url.lastIndexOf("/") + 1);
			if(lastUri.indexOf("?") != -1)
			{
				return lastUri.substring(0, lastUri.indexOf("?"));
			} else
			{
				return lastUri;
			}
		}

		/*
		* Changes current script's working directory
		*/
		function chdir()
		{
			var cwd = document.getElementsByName("cwd")[0];
			if (cwd != "")
			{
				window.location = scriptName() + "?" + "?browse=" + cwd.value
			}
		}

		/*
		* Submits main program's form
		*/
		function formSubmit()
		{
			var actions = document.getElementsByName("action")[0];
			var action = actions.options[actions.selectedIndex].value;
			var actionInput = document.getElementsByName("action_")[0];
			var cwd = document.getElementsByName("cwd_")[0].value;

			switch (action)
			{
				case "copy":
				case "copyo":
				case "move":
				case "zip":
				case "unzip":
					var destination = prompt("Enter Path to Destination Folder", "");
					if (destination)
					{
						actionInput.value = destination;
						return true;
					}
					return false;
				case "properties":
				case "attributes":
					var checkboxes = document.getElementsByName("items");
					for(var i = 0, n = checkboxes.length; i < n; i++)
					{
						if(checkboxes[i].checked)
						{
							window.open(scriptName() + "?" + action + "=" + cwd + "\\" + checkboxes[i].value);
							return false;
						}
					}
					return false;
				case "edit":
				case "download":
					var checkboxes = document.getElementsByName("items");
					for(var i = 0, n = checkboxes.length; i < n; i++)
					{
						if(checkboxes[i].checked && checkboxes[i].value.slice(-1) != "\\")
						{
							window.open(scriptName() + "?" + action + "=" + encodeURIComponent(cwd + "\\" + checkboxes[i].value));
							return false;
						}
					}
					return false;
				case "delete":
					var reassert = confirm("Confirm Delete?");
					if (reassert)
					{
						return true;
					}
					return false;
				case "rename":
					var newName = prompt("Enter a New Name", "");
					if (newName)
					{
						actionInput.value = newName;
						return true;
					}
					return false;
				default:
					return false;
			}
		}

		/**
		 * Toggles checkboxes
		 *
		 * @param object source
		 * @link http://stackoverflow.com/questions/386281/how-to-implement-select-all-check-box-in-html Source
		 */
		function toggle(source)
		{
			var checkboxes = document.getElementsByName("items");
			for(var i = 0, n = checkboxes.length; i < n; i++)
			{
				checkboxes[i].checked = source.checked;
			}
		}
	</script>
</body>
</html>
<%

''
' Create a new blank ZIP file
'
' @link http://www.techcoil.com/blog/handy-vbscript-functions-for-dealing-with-zip-files-and-folders/ Source
' @param string strZipFile Path to the ZIP file
''
Sub zipCreate(strZipFile)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Dim objTextFile
	Set objTextFile = objFSO.CreateTextFile(strZipFile)

	objTextFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)

	objTextFile.Close
	Set objFSO = Nothing
	Set objTextFile = Nothing

	'Wscript.Sleep 500
End Sub

''
' Add a folders contents to a ZIP file
'
' @link http://www.techcoil.com/blog/handy-vbscript-functions-for-dealing-with-zip-files-and-folders/ Source
' @param string strZipFile Path to the ZIP file
' @param string strSource Source folder
''
Sub zipAdd(strZipFile, strSource)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	strZipFile = objFSO.GetAbsolutePathName(strZipFile)
	strSource = objFSO.GetAbsolutePathName(strSource)

	If objFSO.FileExists(strZipFile) Then
		objFSO.DeleteFile strZipFile
	End If

	If Not objFSO.FolderExists(strSource) Then
		Exit Sub
	End If

	zipCreate strZipFile

	dim objShell
	set objShell = CreateObject("Shell.Application")

	Dim objZipFolder
	Set objZipFolder = objShell.NameSpace(strZipFile)
	Dim objFolder
	Set objFolder = objShell.NameSpace(strSource)

	' Look at http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx
	' for more information about the CopyHere function
	objZipFolder.CopyHere objFolder.Items, 4

'	Do Until objFolder.Items.Count <= objZipFolder.Items.Count
'		Wscript.Sleep(200)
'	Loop
End Sub

''
' Extract a ZIP files contents to a folder
'
' @link http://www.techcoil.com/blog/handy-vbscript-functions-for-dealing-with-zip-files-and-folders/ Source
' @param string strZipFile Path to the ZIP file
' @param string strDestination Destination folder
''
Sub zipExtract(strZipFile, strDestination)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strZipFile = objFSO.GetAbsolutePathName(strZipFile)
	strDestination = objFSO.GetAbsolutePathName(strDestination)

	If (Not objFSO.FileExists(strZipFile)) Then
		Exit Sub
	End If

	If Not objFSO.FolderExists(strDestination) Then
		objFSO.CreateFolder(strDestination)
	End If

	dim objShell
	set objShell = CreateObject("Shell.Application")

	Dim objZipFolder
	Set objZipFolder = objShell.NameSpace(strZipFile)

	Dim objFolder
	Set objFolder = objShell.NameSpace(strDestination)

	' Look at http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx
	' for more information about the CopyHere function
	objFolder.CopyHere objZipFolder.Items, 1024

'	Do Until objZipFolder.Items.Count <= objFolder.Items.Count
'		Wscript.Sleep(200)
'	Loop
End Sub

''
' Processes file upload
'
' @param string RequestBin Received binary data from the request
' @link http://www.cymbala.com/Greg/HowToUpload.htm Source
''
Sub BuildUploadRequest(RequestBin)
	Dim UploadControl
	Dim ContentType
	Dim boundary, boundaryPos
	Dim PosFile, Pos, PosBeg, PosEnd, PosBound, FileName
	Dim Name, Value

	' Get the boundary
	PosBeg = 1
	PosEnd = InstrB(PosBeg, RequestBin, strToByte(Chr(13)))
	boundary = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
	boundaryPos = InstrB(1, RequestBin, boundary)
	' Get all data inside the boundaries

	Do Until (boundaryPos = InstrB(RequestBin, boundary & strToByte("--")))

		' Members variable of objects are put in a dictionary object
		Set UploadControl = Server.CreateObject("Scripting.Dictionary")

		'Get an object name
		Pos = InstrB(BoundaryPos, RequestBin, strToByte("Content-Disposition"))
		Pos = InstrB(Pos, RequestBin, strToByte("name="))
		PosBeg = Pos + 6
		PosEnd = InstrB(PosBeg, RequestBin, strToByte(Chr(34)))
		Name = byteToString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
		PosFile = InstrB(BoundaryPos, RequestBin, strToByte("filename="))
		PosBound = InstrB(PosEnd, RequestBin, boundary)

		' Test if object is of file type
		If PosFile <> 0 AND (PosFile<PosBound) Then

			' Get filename, Content-Type and contents of file
			PosBeg = PosFile + 10
			PosEnd = InstrB(PosBeg, RequestBin, strToByte(Chr(34)))
			FileName = byteToString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))

			' Add filename to dictionary object
			UploadControl.Add "FileName", FileName
			Pos = InstrB(PosEnd, RequestBin, strToByte("Content-Type:"))
			PosBeg = Pos + 14
			PosEnd = InstrB(PosBeg, RequestBin, strToByte(Chr(13)))

			' Add Content-Type to dictionary object
			ContentType = byteToString(MidB(RequestBin, PosBeg, PosEnd-PosBeg))
			UploadControl.Add "ContentType", ContentType

			' Get content of object
			PosBeg = PosEnd + 4
			PosEnd = InstrB(PosBeg, RequestBin, boundary) - 2
			Value = MidB(RequestBin, PosBeg, PosEnd - PosBeg)

		Else

			'Get content of object
			Pos = InstrB(Pos, RequestBin, strToByte(Chr(13)))
			PosBeg = Pos + 4
			PosEnd = InstrB(PosBeg, RequestBin, boundary) - 2
			Value = byteToString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
		End If

		' Add content to dictionary object
		UploadControl.Add "Value" , Value

		' Add dictionary object to main dictionary
		UploadRequest.Add name, UploadControl

		' Loop to next object
		BoundaryPos = InstrB(BoundaryPos + LenB(boundary), RequestBin, boundary)
	Loop

End Sub

''
' Converts string to byte
'
' @param string strString Input string
' @link http://www.cymbala.com/Greg/HowToUpload.htm Source
''
Function strToByte(strString)
	Dim strChar, i
	For i = 1 to Len(strString)
		strChar = Mid(strString, i, 1)
		strToByte = strToByte & ChrB(AscB(strChar))
	Next
End Function

''
' Converts byte to string
'
' @param string StringBin
' @link http://www.cymbala.com/Greg/HowToUpload.htm Source
''
Function byteToString(StringBin)
	Dim j
	byteToString = ""
	For j = 1 to LenB(StringBin)
		byteToString = byteToString & Chr(AscB(MidB(StringBin,j,1)))
	Next
End Function

''
' Converts size in bytes to another unit
'
' @param int intSize Input file size
' @return string Returns converted file size with its unit
''
Function convertSize(intSize)
	If intSize <= 1024 Then
		convertSize = intSize & " Bytes"
	ElseIf intSize <= 1048576 Then
		convertSize = Round(intSize / 1024, 2) & " KB"
	ElseIf intSize <= 1073741824 Then
		convertSize = Round(intSize / 1048576, 2) & " MB"
	ElseIf intSize <= 1099511627776 Then
		convertSize = Round(intSize / 1073741824, 2) & " GB"
	Else
		convertSize = Round(intSize / 1099511627776, 2) & " TB"
	End If
End Function

''
' Reads a files contents into string
'
' @param string strFile Path to the file
' @param string strCharset Encoding for reading the file
' @return string Returns the files contents
''
Function fileReadText(strFile, strCharset)
	Dim objFSO, objFile, objStream
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(strFile)

	Const adTypeText = 2
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.CharSet = strCharset
	objStream.Type = adTypeText
	objStream.Open
	objStream.LoadFromFile objFile.Path
	fileReadText = objStream.ReadText

	Set objFSO = Nothing
	Set objFile = Nothing
	Set objStream = Nothing
End Function

''
' Writes a string into a file
'
' @param string strFile Path to the file
' @param string strData Data to be written
' @param string strCharset Encoding for writing the file
''
Function fileWriteText(strFile, strData, strCharset)
	Dim objFSO, objFile, objStream
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(strFile)

	Const adTypeText = 2
	Const adSaveCreateOverWrite = 2

	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.CharSet = strCharset
	objStream.Type = adTypeText
	objStream.Open
	objStream.Position = 0
	objStream.WriteText strData
	objStream.SaveToFile objFile.Path, adSaveCreateOverWrite

	Set objFSO = Nothing
	Set objFile = Nothing
	Set objStream = Nothing
End Function

''
' Changes a files last modified date
'
' @param string strFile Path to the file
' @param string strDate New files last modified date
' @param bool Returns TRUE on success
''
Function fileDateLastModified(strFile, strDate)
	If Right(strFile, 1) = "\" Then Exit Function
	If Not IsDate(strDate) Then Exit Function

	Dim objFSO, objShell, objFolder, objFolderItem
	Dim strParent, strFilename

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	strParent = objFSO.GetParentFolderName(strFile)
	strFilename = objFSO.GetFileName(strFile)

	Set objShell = Server.CreateObject("Shell.Application")
	Set objFolder = objShell.NameSpace(strParent)

	Set objFolderItem = objFolder.ParseName(strFileName)
	If Not objFolderItem Is Nothing Then
		objFolderItem.ModifyDate = strDate
		fileDateLastModified = True
	End If
End Function

''
' Parses a files or a folders attributes
'
' @param object objAttributes The attribute object from FileSystemObject
' @param string Returns a string represent the attributes of the file or folder
''
Function fsAttributes(objAttributes)
	Dim strAttributeValue, colKeys, strKey
	strAttributeValue = ""

	colKeys = dAttributes.Keys
	For Each strKey In colKeys
		If objAttributes And dAttributes.Item(strKey)(0) Then
			strAttributeValue = strAttributeValue & strKey
		Else
			strAttributeValue = strAttributeValue & "-"
		End If
	Next

	fsAttributes = strAttributeValue
End Function

''
' Creates a new file or folder
'
' @param string strPath Path to the new file
' @param string strNew "file" or "folder"
''
Function fsCreate(strPath, strNew)
response.write strPath
	Dim objFSO, objTextFile, objFolder
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If strNew = "file" Then
		If Not objFSO.FileExists(strPath) Then
			Set objTextFile = objFSO.CreateTextFile(strPath)
		End If
	ElseIf strNew = "folder" Then
		If Not objFSO.FolderExists(strPath) Then
			Set objFolder = objFSO.CreateFolder(strPath)
		End If
	End If

	Set objFSO = Nothing
End Function

''
' Copy a file or folder
'
' @param string strItem Input file or folder
' @param string strNewName New file name
' @param string strSource Source folder of strItem
' @param string strDestination Destination where the file or folder is to be copied. Wildcard characters are not allowed.
' @param bool bOverwrite Boolean value that is True (default) if existing files or folders are to be overwritten; False if they are not.
''
Function fsCopy(strItem, strNewName, strSource, strDestination, bOverwrite)
	Dim objFSO, objItem
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If Right(strItem, 1) = "\" Then
		strItem = strSource & "\" & strItem
		Set objItem = objFSO.GetFolder(strItem)
	Else
		strDestination = strDestination & "\" & strNewName
		strItem = strSource & "\" & strItem
		Set objItem = objFSO.GetFile(strItem)
	End If

	If bOverwrite = True Then
		objItem.Copy strDestination, True
	Else
		objItem.Copy strDestination, False
	End If

	Set objFSO = Nothing
	Set objItem = Nothing
End Function

''
' Deletes a file or folder
'
' @param string strItem Input file or folder
''
Function fsDelete(strItem)
	Dim objFSO, objItem
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If Right(strItem, 1) = "\" Then
		Set objItem = objFSO.GetFolder(strItem)
	Else
		Set objItem = objFSO.GetFile(strItem)
	End If

	objItem.Delete
	Set objFSO = Nothing
	Set objItem = Nothing
End Function

''
' Moves a file or folder
'
' @param string strItem Input file or folder
' @param string strNewName New file name
' @param string strSource Source folder of strItem
' @param string strDestination Destination where the file or folder is to be moved. Wildcard characters are not allowed.
''
Function fsMove(strItem, strNewName, strSource, strDestination)
	Dim objFSO, objItem
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If Right(strItem, 1) = "\" Then
		strItem = strSource & "\" & strItem
		Set objItem = objFSO.GetFolder(strItem)
	Else
		strDestination = strDestination & "\" & strNewName
		strItem = strSource & "\" & strItem
		Set objItem = objFSO.GetFile(strItem)
	End If

	objItem.Move strDestination

	Set objFSO = Nothing
	Set objItem = Nothing
End Function

''
' Renames a file or folder
'
' @param string strItem Input file or folder
' @param string strNewName New file name
' @param string strSource Source folder of strItem
''
Function fsRename(strItem, strNewName, strSource)
	Dim objFSO, objItem
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	strNewName = strValidFilename(strNewName)
	strItem = strSource & "\" & strItem
	If Right(strItem, 1) = "\" Then
		Set objItem = objFSO.GetFolder(strItem)
	Else
		Set objItem = objFSO.GetFile(strItem)
	End If

	objItem.Move strSource & "\" & strNewName

	Set objFSO = Nothing
	Set objItem = Nothing
End Function

''
' Generates tabs
'
' @param int intCount Number of tabs
''
Function Tab(intCount)
	If intCount > 0 Then
		Dim arTmp()
		ReDim arTmp(intCount)
		Tab = Join(arTmp, vbTab)
	End If
End Function

''
' Escapes backslash in a string
'
' @param string strString Input string
''
Function escapeBackslash(ByVal strString)
	If ((Not IsNull(strString)) And (strString <> "")) Then
		strString = Replace(strString, "\", "\\")
	End If

	escapeBackslash = strString
End Function

''
' Replaces HTML special characters
'
' @param string strString Input string
' @return string Returns replaced string
''
Function strConvertHTML(ByVal strString)
	If ((Not IsNull(strString)) And (strString <> "")) Then
		strString = Replace(strString, "&", "&amp;")
		strString = Replace(strString, "<", "&lt;")
		strString = Replace(strString, ">", "&gt;")
		strString = Replace(strString, """", "&quot;")
		strString = Replace(strString, "'", "&apos;")
	End If

    strConvertHTML = strString
End Function

Function strValidFilename(strFilename)
	If ((Not IsNull(strFilename)) And (strFilename <> "")) Then
		strFilename = Replace(strFilename, "\", "_")
		strFilename = Replace(strFilename, "/", "_")
		strFilename = Replace(strFilename, ":", "_")
		strFilename = Replace(strFilename, "*", "_")
		strFilename = Replace(strFilename, "?", "_")
		strFilename = Replace(strFilename, """", "_")
		strFilename = Replace(strFilename, "<", "_")
		strFilename = Replace(strFilename, ">", "_")
		strFilename = Replace(strFilename, "|", "_")
	End If

	strValidFilename = strFilename
End Function

''
' Recursively lists contents of a folder
'
' @param object objFolder The folder object from FileSystemObject
' @param int intCurrentLevel Current crawling depth
''
Sub ShowSubFolders(objFolder, intCurrentLevel)
	If intCurrentLevel < intMaxLevel Or intMaxLevel = -1 Then
		For Each Subfolder In objFolder.SubFolders
			Set objSubFolder = objFSO.GetFolder(Subfolder.Path)
			Set colFiles = objSubFolder.Files
			For Each objFile In colFiles
				Response.Write vbCRLF + Subfolder.Path + "\" + objFile.Name
			Next
			ShowSubFolders Subfolder, (intCurrentLevel + 1)
		Next
	End If
End Sub
%>