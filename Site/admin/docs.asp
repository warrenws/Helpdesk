<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 2/13/12
'Last Updated 6/16/14

'This is the docs page.

Option Explicit

On Error Resume Next

Dim objNetwork, strUserAgent, strSQL, objNameCheckSet, strRole, strPath, intInputSize, objFSO
Dim objFolder, objFile, objSubFolder, strFileType, strTitle, intLayer, strUp, objUp, intIndex
Dim intSize, strSize, Upload, strNewFolder, strMessage, bolShowLogout, strUser

'Redirect the user the SSL version if required
If Application("ForceSSL") Then
   If Request.ServerVariables("SERVER_PORT")=80 Then
      If Request.ServerVariables("QUERY_STRING") = "" Then
         Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
      Else
         Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
      End If
   End If
End If
'If the database and the website are not the same version then let them know
If Application("VersionError") Then
   VersionProblem
End If

'Get the users logon name
Set objNetwork = CreateObject("WSCRIPT.Network")   
strUser = objNetwork.UserName
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

'Check and see if anonymous access is enabled
If LCase(Left(strUser,4)) = "iusr" Then
   strUser = GetUser
   bolShowLogout = True
Else
   bolShowLogout = False
End If

'Find the current users role
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"

Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)

'See if the user has the rights to visit this page
If objNameCheckSet(2) Then

   'An error would be generated if the user has NTFS rights to get in but is not found
   'in the database.  In this case the user is denied access.
   If Err Then
      AccessDenied
   Else
      If objNameCheckSet(6) <> "Deny" Then
         AccessGranted
      Else
         AccessDenied
      End If
   End If
Else
   AccessDenied
End If%>

<%Sub AccessGranted 

   On Error Resume Next

   If InStr(strUserAgent,"Android") Or InStr(strUserAgent,"Silk") Then
      intInputSize = 65
   Else
      intInputSize = 75
   End If

   strPath = Request.QueryString("Path")
   intLayer = Request.QueryString("Layer")
   
   If InStr(strPath,"..") Then
      strPath = ""
   End If
   
   If intLayer = "" Or intLayer = 0 Then
      intLayer = 1
   Else   
      objUp = Split(strPath,"/")
      For intIndex = 0 to (intLayer - 3)
         strUp = strUp & objUp(intIndex) & "/"
      Next
      If Right(strUp,1) = "/" Then
         strUp = Left(strUp,(Len(strUp)-1))
      End If
   End If
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   
   strTitle = "Documentation/" & strPath
   
   Set Upload = New FreeASPUpload
   Upload.Save(Application("FileLocation"))
   
   Select Case Upload.Form("cmdSubmit")
      Case "New Folder"
         strNewFolder = Upload.Form("NewFolder")
         If strPath = "" Then
            objFSO.CreateFolder Application("DocLocation") & "\" & strPath & strNewFolder
         Else
            objFSO.CreateFolder Application("DocLocation") & "\" & strPath & "\" & strNewFolder
         End If
         If Err Then
            strMessage = Err.Description 
            Err.Clear
         Else 
            strMessage = "Folder Created"
         End If   
      Case "Upload"
         Upload.Save(Application("DocLocation") & "\" & strPath)
         strMessage = "File Uploaded"
      Case "Yes"
         Select Case Upload.Form("Type")
            Case "File"
               objFSO.DeleteFile Upload.Form("FilePath")
            Case "Folder"
               objFSO.DeleteFolder Upload.Form("FilePath")
         End Select
         strMessage = Upload.Form("FileName") & " Deleted..."
   End Select
   
%>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 9") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.99, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 6") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.53, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"GT-N5110") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.77, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <meta name="theme-color" content="#<%=Application("AndroidBarColor")%>">
   </head>

   <body>
   <div class="header">
      <%=Application("SchoolName")%> Help Desk
   </div>
   
   <div class="version">
      Version <%=Application("Version")%>
   </div>
   
   <hr class="admintopbar" />
   <div class="admintopbar">
      <ul class="topbar">
			<li class="topbar"><a href="index.asp">Home</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="view.asp?Filter=AllOpenTickets">Open Tickets</a><font class="separator"> | </font></li>
      <% If strRole <> "Data Viewer" Then %>
         <li class="topbar"><a href="view.asp?Filter=MyOpenTickets">Your Tickets</a><font class="separator"> | </font></li>  
      <% End If %>
      <% If Application("UseTaskList") Then %>
         <li class="topbar"><a href="tasklist.asp">Tasks</a><font class="separator"> | </font></li>
      <% End If %>
		<% If Application("UseStats") Then %>
			<li class="topbar"><a class="linkbar" href="stats.asp">Stats</a><font class="separator"> | </font></li> 
      <% End If %>
      <% If Application("UseDocs") And objNameCheckSet(6) <> "Deny" Then %>
         <li class="topbar">Docs<font class="separator"> | </font></li>
      <% End If %>
         <li class="topbar"><a href="settings.asp">Settings</a>
      <% If objNameCheckSet(1) = "Administrator" Then %> 
         <font class="separator"> | </font></li>
         <li class="topbar"><a class="linkbar" href="setup.asp">Admin Mode</a>
      <% Else %>
         </li>
      <% End If %>
      <% If bolShowLogout Then %>
         <font class="separator"> | </font></li>
         <li class="topbar"><a class="linkbar" href="login.asp?action=logout">Log Out</a></li>
      <% Else %>
         </li>
      <% End If %>
		</ul>
   </div>
   
<% If InStr(strUserAgent,"MSIE") Then %>
      <hr class="adminbottombarIE"/>
<% Else %>   
      <hr class="adminbottombar"/>
<% End If %>

   <center>
   <table width="750">
      <tr><th colspan="5"><%=strTitle%></th></tr>
      <tr><th colspan="5"><hr/></th></tr>
   <% If Upload.Form("cmdSubmit") = "Delete" Then %>   
         <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="docs.asp?path=<%=strPath%>&layer=<%=intLayer%>">
         <tr>
            <td colspan="5" align="center">
               Are you sure you want to delete the <%=LCase(Upload.Form("Type"))%>&nbsp;"<%=Upload.Form("FileName")%>"?
               &nbsp;&nbsp;<input type="submit" value="Yes" name="cmdSubmit">&nbsp;<input type="submit" value="No" name="cmdSubmit">
               <input type="hidden" value="<%=Upload.Form("FilePath")%>" name="FilePath"/>
               <input type="hidden" value="<%=Upload.Form("FileName")%>" name="FileName"/>
               <input type="hidden" value="<%=Upload.Form("Type")%>" name="Type"/>
            </td>
         </tr>
         </form>
   <% Else %>
      <% If objNameCheckSet(6) = "Full" Then %>
            <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="docs.asp?path=<%=strPath%>&layer=<%=intLayer%>">
            <tr>
               <td colspan="5" align="center">
                  Create a Folder: <input type="text" Name="NewFolder"> <input type="submit" value="New Folder" name="cmdSubmit">
               </td>
            </tr>
            
            <tr>
               <td colspan="5" align="center">
               <% If inStr(strUserAgent,"iPad") = False And inStr(strUserAgent,"iPhone") = False Then
                     If InStr(strUserAgent,"Chrome") or InStr(strUserAgent,"Safari") Then %>
                        Upload a File: <input class="fileuploadchrome" type="file" name="Attachment" size="50">
                     <%Else%>
                        Upload a File: <input class="fileupload" type="file" name="Attachment" size="50">
                     <%End If%>
                     <input type="submit" value="Upload" name="cmdSubmit">
                  <%End If%>
               </td>
            </tr>
         <% If strMessage <> "" Then %>
               <tr>
                  <td colspan="5" align="center">
               <% If strMessage = "Folder Created" or strMessage = "File Uploaded" or InStr(strMessage,"Deleted...")Then %>
                     <font class="information"><%=strMessage%></font>
               <% Else %>
                     <font class="missing"><%=strMessage%></font>
               <% End If %>
                  </td>
               </tr>
         <% End If %>
            </form>
      <% End If%>
   <% End If %>
      <tr><th colspan="5"><hr/></th></tr>
      <tr>
         <th>&nbsp;</th>
         <th>File Name</th>
         <th>Date Modified</th>
         <th>Size</th>
      </tr>
   <% If strPath <> "" Then %>
         <tr>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <td class="showborders" width="1%">
                  <a href="docs.asp?path=<%=strUp%>&Layer=<%=intLayer - 1%>">
                     <img border="0" src="../themes/<%=Application("Theme")%>/images/filetypes/upfolder.png">
                  </a>
               </td>
            <% Else %>
               <td class="showborders" width="1%">
                  <a href="docs.asp?path=<%=strUp%>&Layer=<%=intLayer - 1%>">
                     <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/filetypes/upfolder.png">
                  </a>
               </td>
            <% End If %>
            <td class="showborders"><a href="docs.asp?path=<%=strUp%>&Layer=<%=intLayer - 1%>">.. Up One Level</a></td>
            <td class="showborders">&nbsp;</td>
            <td class="showborders">&nbsp;</td>
            <td class="showborders">&nbsp;</td>
         </tr>
   <% End If %>
<%    Set objFolder = objFSO.GetFolder(Application("DocLocation") & "\" & strPath)
      For Each objSubFolder in objFolder.SubFolders %>
         <tr>
         <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
            <td class="showborders" width="1%">
            <% If strPath = "" Then %>
               <a href="docs.asp?path=<%=objSubFolder.Name%>&Layer=<%=intLayer + 1%>">
            <% Else %>
               <a href="docs.asp?path=<%=strPath & "/"%><%=objSubFolder.Name%>&Layer=<%=intLayer + 1%>">
            <% End If %>
                  <img border="0" src="../themes/<%=Application("Theme")%>/images/filetypes/folder.png">
               </a>
            </td>
         <% Else %>
            <td class="showborders" width="1%">
            <% If strPath = "" Then %>
               <a href="docs.asp?path=<%=objSubFolder.Name%>&Layer=<%=intLayer + 1%>">
            <% Else %>
               <a href="docs.asp?path=<%=strPath & "/"%><%=objSubFolder.Name%>&Layer=<%=intLayer + 1%>">
            <% End If %>   
                  <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/filetypes/folder.png">
               </a>
            </td>
         <% End If %>
            <% If strPath = "" Then %>
               <td class="showborders"><a href="docs.asp?path=<%=objSubFolder.Name%>&Layer=<%=intLayer + 1%>"><%=objSubFolder.Name%></a></td>
            <% Else %>
               <td class="showborders"><a href="docs.asp?path=<%=strPath & "/"%><%=objSubFolder.Name%>&Layer=<%=intLayer + 1%>"><%=objSubFolder.Name%></a></td>
            <% End If %>
            <td class="showborders"><%=objSubFolder.DateLastModified%></td>
            <td class="showborders">&nbsp;</td>
            <% If objNameCheckSet(6) = "Full" Then %>
                  <td class="showborders">
                     <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="docs.asp?path=<%=strPath%>&layer=<%=intLayer%>">
                        <input type="hidden" value="<%=objSubFolder.Path%>" name="FilePath"/>
                        <input type="hidden" value="<%=objSubFolder.Name%>" name="FileName"/>
                        <input type="hidden" value="Folder" name="Type"/>
                        <input type="submit" value="Delete" name="cmdSubmit" />
                     </form>
                  </td>
            <% End If %>
         </tr>
   <% Next
      For Each objFile in objFolder.Files
         If UCase(objFile.Name) <> "THUMBS.DB" And UCase(objFile.Name) <> ".DS_STORE" Then 
            Select Case LCase(Right(objFile.Name,(Len(objFile.Name) - InStr(objFile.Name,"."))))
               Case "bat"
                  strFileType = "bat.png"
               Case "bmp"
                  strFileType = "bmp.png"
               Case "doc"
                  strFileType = "doc.png"
               Case "docx"
                  strFileType = "docx.png"
               Case "gif"
                  strFileType = "gif.png"
               Case "html"
                  strFileType = "html.png"
               Case "jpg"
                  strFileType = "jpg.png"
               Case "log"
                  strFileType = "log.png"
               Case "mp3"
                  strFileType = "mp3.png"
               Case "pdf"
                  strFileType = "pdf.png"
               Case "png"
                  strFileType = "png.png"
               Case "ppt"
                  strFileType = "ppt.png"
               Case "pptx"
                  strFileType = "pptx.png"
               Case "rar"
                  strFileType = "rar.png"
               Case "rtf"
                  strFileType = "rtf.png"
               Case "txt"
                  strFileType = "txt.png"
               Case "xls"
                  strFileType = "xls.png"
               Case "xlsx", "xlsm"
                  strFileType = "xlsx.png"
               Case "zip"
                  strFileType = "zip.png"
               Case Else
                  strFileType = "file.png"
            End Select
            %>
            <tr>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <td class="showborders">
               <% If strPath = "" Then %>
                     <a href="docs/<%=objFile.Name%>">
               <% Else %>
                     <a href="docs/<%=strPath%>/<%=objFile.Name%>">
               <% End If %>
                     <img border="0" src="../themes/<%=Application("Theme")%>/images/filetypes/<%=strFileType%>">
                  </a>
               </td>         
            <% Else %>
               <td class="showborders">
               <% If strPath = "" Then %>
                     <a href="docs/<%=objFile.Name%>">
               <% Else %>
                     <a href="docs/<%=strPath%>/<%=objFile.Name%>">
               <% End If %>
                     <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/filetypes/<%=strFileType%>">
                  </a>
               </td>
            <% End If %>
            <% If strPath = "" Then %>
               <td class="showborders"><a href="download.asp?source=docs&file=<%=objFile.Name%>"><%=objFile.Name%></a></td>
            <% Else %>
               <td class="showborders"><a href="download.asp?source=docs&folder=<%=strPath%>&file=<%=objFile.Name%>"><%=objFile.Name%></a></td>
            <% End If %>
               <td class="showborders"><%=objFile.DateLastModified%></td>
			<% If objFile.Size > 1048576 Then
               intSize = Round(objFile.Size/1048576,2)
               strSize = intSize & " MB"
            ElseIf objFile.Size > 1024 Then
               intSize = Round(objFile.Size/1024,2)
               strSize = intSize & " KB"               
            Else
               strSize = Round(objFile.Size,2) & " Bytes"
			   End If %>
               <td class="showborders"><%=strSize%></td>
            <% If objNameCheckSet(6) = "Full" Then %>
                  <td class="showborders">
                     <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="docs.asp?path=<%=strPath%>&layer=<%=intLayer%>">
                        <input type="hidden" value="<%=objFile.Path%>" name="FilePath"/>
                        <input type="hidden" value="<%=objFile.Name%>" name="FileName"/>
                        <input type="hidden" value="File" name="Type"/>
                        <input type="submit" value="Delete" name="cmdSubmit" />
                     </form>
                  </td>
            <% End If %>
            </tr>
      <% End If
      Next%>

<%End Sub %>

<%Sub AccessDenied 

   If bolShowLogout Then
      Response.Redirect("login.asp?action=logout")
   Else
   %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
      <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
      <meta name="viewport" content="width=device-width" />
   </head>
   <body>
      <center><b>Access Denied</b></center>
   </body>
   </html>
   
<% End If

End Sub%>

<%
'  For examples, documentation, and your own free copy, go to:
'  http://www.freeaspupload.net
'  Note: You can copy and use this script for free and you can make changes
'  to the code, but you cannot remove the above comment.

'Changes:
'Aug 2, 2005: Add support for checkboxes and other input elements with multiple values
'Jan 6, 2009: Lars added ASP_CHUNK_SIZE
'Sep 3, 2010: Enforce UTF-8 everywhere; new function to convert byte array to unicode string

const DEFAULT_ASP_CHUNK_SIZE = 200000

const adModeReadWrite = 3
const adTypeBinary = 1
const adTypeText = 2
const adSaveCreateOverWrite = 2

Class FreeASPUpload
	Public UploadedFiles
	Public FormElements

	Private VarArrayBinRequest
	Private StreamRequest
	Private uploadedYet
	Private internalChunkSize

	Private Sub Class_Initialize()
		Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
		Set FormElements = Server.CreateObject("Scripting.Dictionary")
		Set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = adTypeText
		StreamRequest.Open
		uploadedYet = false
		internalChunkSize = DEFAULT_ASP_CHUNK_SIZE
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(UploadedFiles) Then
			UploadedFiles.RemoveAll()
			Set UploadedFiles = Nothing
		End If
		If IsObject(FormElements) Then
			FormElements.RemoveAll()
			Set FormElements = Nothing
		End If
		StreamRequest.Close
		Set StreamRequest = Nothing
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If FormElements.Exists(LCase(sIndex)) Then Form = FormElements.Item(LCase(sIndex))
	End Property

	Public Property Get Files()
		Files = UploadedFiles.Items
	End Property
	
    Public Property Get Exists(sIndex)
            Exists = false
            If FormElements.Exists(LCase(sIndex)) Then Exists = true
    End Property
        
    Public Property Get FileExists(sIndex)
        FileExists = false
            if UploadedFiles.Exists(LCase(sIndex)) then FileExists = true
    End Property
        
    Public Property Get chunkSize()
		chunkSize = internalChunkSize
	End Property

	Public Property Let chunkSize(sz)
		internalChunkSize = sz
	End Property

	'Calls Upload to extract the data from the binary request and then saves the uploaded files
	Public Sub Save(path)
		Dim streamFile, fileItem, filePath

		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload

		For Each fileItem In UploadedFiles.Items
			filePath = path & fileItem.FileName
			Set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = adTypeBinary
			streamFile.Open
			StreamRequest.Position=fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile filePath, adSaveCreateOverWrite
			streamFile.close
			Set streamFile = Nothing
			fileItem.Path = filePath
		 Next
	End Sub
	
	public sub SaveOne(path, num, byref outFileName, byref outLocalFileName)
		Dim streamFile, fileItems, fileItem, fs

        set fs = Server.CreateObject("Scripting.FileSystemObject")
		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload
		if UploadedFiles.Count > 0 then
			fileItems = UploadedFiles.Items
			set fileItem = fileItems(num)
		
			outFileName = fileItem.FileName
			outLocalFileName = GetFileName(path, outFileName)
		
			Set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = adTypeBinary
			streamFile.Open
			StreamRequest.Position = fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile path & outLocalFileName, adSaveCreateOverWrite
			streamFile.close
			Set streamFile = Nothing
			fileItem.Path = path & filename
		end if
	end sub

	Public Function SaveBinRequest(path) ' For debugging purposes
		StreamRequest.SaveToFile path & "\debugStream.bin", 2
	End Function

	Public Sub DumpData() 'only works if files are plain text
		Dim i, aKeys, f
		response.write "Form Items:<br>"
		aKeys = FormElements.Keys
		For i = 0 To FormElements.Count -1 ' Iterate the array
			response.write aKeys(i) & " = " & FormElements.Item(aKeys(i)) & "<BR>"
		Next
		response.write "Uploaded Files:<br>"
		For Each f In UploadedFiles.Items
			response.write "Name: " & f.FileName & "<br>"
			response.write "Type: " & f.ContentType & "<br>"
			response.write "Start: " & f.Start & "<br>"
			response.write "Size: " & f.Length & "<br>"
		 Next
   	End Sub

	Public Sub Upload()
		Dim nCurPos, nDataBoundPos, nLastSepPos
		Dim nPosFile, nPosBound
		Dim sFieldName, osPathSep, auxStr
		Dim readBytes, readLoop, tmpBinRequest
		
		'RFC1867 Tokens
		Dim vDataSep
		Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
		tNewLine = String2Byte(Chr(13))
		tDoubleQuotes = String2Byte(Chr(34))
		tTerm = String2Byte("--")
		tFilename = String2Byte("filename=""")
		tName = String2Byte("name=""")
		tContentDisp = String2Byte("Content-Disposition")
		tContentType = String2Byte("Content-Type:")

		uploadedYet = true

		On Error resume next
			' Copy binary request to a byte array, on which functions like InstrB and others can be used to search for separation tokens
			readBytes = internalChunkSize
			VarArrayBinRequest = Request.BinaryRead(readBytes)
			VarArrayBinRequest = midb(VarArrayBinRequest, 1, lenb(VarArrayBinRequest))
			Do Until readBytes < 1
				tmpBinRequest = Request.BinaryRead(readBytes)
				if readBytes > 0 then
					VarArrayBinRequest = VarArrayBinRequest & midb(tmpBinRequest, 1, lenb(tmpBinRequest))
				end if
			Loop
			StreamRequest.WriteText(VarArrayBinRequest)
			StreamRequest.Flush()
			if Err.Number <> 0 then 
				response.write "<br><br><B>System reported this error:</B><p>"
				response.write Err.Description & "<p>"
				response.write "The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
				Exit Sub
			end if
		On Error goto 0 'reset error handling

		nCurPos = FindToken(tNewLine,1) 'Note: nCurPos is 1-based (and so is InstrB, MidB, etc)

		If nCurPos <= 1  Then Exit Sub
		 
		'vDataSep is a separator like -----------------------------21763138716045
		vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)

		'Start of current separator
		nDataBoundPos = 1

		'Beginning of last line
		nLastSepPos = FindToken(vDataSep & tTerm, 1)

		Do Until nDataBoundPos = nLastSepPos
			
			nCurPos = SkipToken(tContentDisp, nDataBoundPos)
			nCurPos = SkipToken(tName, nCurPos)
			sFieldName = ExtractField(tDoubleQuotes, nCurPos)

			nPosFile = FindToken(tFilename, nCurPos)
			nPosBound = FindToken(vDataSep, nCurPos)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile
				Set oUploadFile = New UploadedFile
				
				nCurPos = SkipToken(tFilename, nCurPos)
				auxStr = ExtractField(tDoubleQuotes, nCurPos)
                ' We are interested only in the name of the file, not the whole path
                ' Path separator is \ in windows, / in UNIX
                ' While IE seems to put the whole pathname in the stream, Mozilla seem to 
                ' only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
				oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))

				if (Len(oUploadFile.FileName) > 0) then 'File field not left empty
					nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
                    ' NN on UNIX puts things like this in the stream:
                    '    ?? python py type=?? python application/x-python
					oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
					nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
					
					oUploadFile.Start = nCurPos+1
					oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
					
					If oUploadFile.Length > 0 Then UploadedFiles.Add LCase(sFieldName), oUploadFile
				End If
			Else
				Dim nEndOfData, fieldValueUniStr
				nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				fieldValueuniStr = ConvertUtf8BytesToString(nCurPos, nEndOfData-nCurPos)
				If Not FormElements.Exists(LCase(sFieldName)) Then 
					FormElements.Add LCase(sFieldName), fieldValueuniStr
				else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & fieldValueuniStr
                end if 

			End If

			'Advance to next separator
			nDataBoundPos = FindToken(vDataSep, nCurPos)
		Loop
	End Sub

	Private Function SkipToken(sToken, nStart)
		SkipToken = InstrB(nStart, VarArrayBinRequest, sToken)
		If SkipToken = 0 then
			Response.write "Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
			Response.End
		end if
		SkipToken = SkipToken + LenB(sToken)
	End Function

	Private Function FindToken(sToken, nStart)
		FindToken = InstrB(nStart, VarArrayBinRequest, sToken)
	End Function

	Private Function ExtractField(sToken, nStart)
		Dim nEnd
		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		If nEnd = 0 then
			Response.write "Error in parsing uploaded binary request."
			Response.End
		end if
		ExtractField = ConvertUtf8BytesToString(nStart, nEnd-nStart)
	End Function

	'String to byte string conversion
	Private Function String2Byte(sString)
		Dim i
		For i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		Next
	End Function

	Private Function ConvertUtf8BytesToString(start, length)	
		StreamRequest.Position = 0
	
	    Dim objStream
	    Dim strTmp
	    
	    ' init stream
	    Set objStream = Server.CreateObject("ADODB.Stream")
	    objStream.Charset = "utf-8"
	    objStream.Mode = adModeReadWrite
	    objStream.Type = adTypeBinary
	    objStream.Open
	    
	    ' write bytes into stream
	    StreamRequest.Position = start+1
	    StreamRequest.CopyTo objStream, length
	    objStream.Flush
	    
	    ' rewind stream and read text
	    objStream.Position = 0
	    objStream.Type = adTypeText
	    strTmp = objStream.ReadText
	    
	    ' close up and return
	    objStream.Close
	    Set objStream = Nothing
	    ConvertUtf8BytesToString = strTmp	
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public Start
	Public Length
	Public Path
	Private nameOfFile

    ' Need to remove characters that are valid in UNIX, but not in Windows
    Public Property Let FileName(fN)
        nameOfFile = fN
        nameOfFile = SubstNoReg(nameOfFile, "\", "_")
        nameOfFile = SubstNoReg(nameOfFile, "/", "_")
        nameOfFile = SubstNoReg(nameOfFile, ":", "_")
        nameOfFile = SubstNoReg(nameOfFile, "*", "_")
        nameOfFile = SubstNoReg(nameOfFile, "?", "_")
        nameOfFile = SubstNoReg(nameOfFile, """", "_")
        nameOfFile = SubstNoReg(nameOfFile, "<", "_")
        nameOfFile = SubstNoReg(nameOfFile, ">", "_")
        nameOfFile = SubstNoReg(nameOfFile, "|", "_")
    End Property

    Public Property Get FileName()
        FileName = nameOfFile
    End Property

    'Public Property Get FileN()ame
End Class


' Does not depend on RegEx, which is not available on older VBScript
' Is not recursive, which means it will not run out of stack space
Function SubstNoReg(initialStr, oldStr, newStr)
    Dim currentPos, oldStrPos, skip
    If IsNull(initialStr) Or Len(initialStr) = 0 Then
        SubstNoReg = ""
    ElseIf IsNull(oldStr) Or Len(oldStr) = 0 Then
        SubstNoReg = initialStr
    Else
        If IsNull(newStr) Then newStr = ""
        currentPos = 1
        oldStrPos = 0
        SubstNoReg = ""
        skip = Len(oldStr)
        Do While currentPos <= Len(initialStr)
            oldStrPos = InStr(currentPos, initialStr, oldStr)
            If oldStrPos = 0 Then
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, Len(initialStr) - currentPos + 1)
                currentPos = Len(initialStr) + 1
            Else
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, oldStrPos - currentPos) & newStr
                currentPos = oldStrPos + skip
            End If
        Loop
    End If
End Function

Function GetFileName(strSaveToPath, FileName)
'This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
'It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
'It keeps going until it returns a filename that does not exist.
'You could just create a filename from the ID field but that means writing the record - and it still might exist!
'N.B. Requires strSaveToPath variable to be available - and containing the path to save to
    Dim Counter
    Dim Flag
    Dim strTempFileName
    Dim FileExt
    Dim NewFullPath
    dim objFSO, p
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    p = instrrev(FileName, ".")
    FileExt = mid(FileName, p+1)
    strTempFileName = left(FileName, p-1)
    NewFullPath = strSaveToPath & "\" & FileName
    Flag = False
    
    Do Until Flag = True
        If objFSO.FileExists(NewFullPath) = False Then
            Flag = True
            GetFileName = Mid(NewFullPath, InstrRev(NewFullPath, "\") + 1)
        Else
            Counter = Counter + 1
            NewFullPath = strSaveToPath & "\" & strTempFileName & Counter & "." & FileExt
        End If
    Loop
End Function 
%>

<%
Function BuildReturnLink(bolIncludeID)

   Dim strLinkID, strLinkLocation, strLinkStatus, strLinkTech, strLinkCategory, strLinkUser, strLinkFilter, strLinkProblem
   Dim strLinkNotes, strLinkEMail, strLinkSort, strLinkDays, strLinkBack, strLinkViewed

   'Build the return link
   If bolIncludeID Then
      strLinkID = Request.QueryString("ID")
   End If
   strLinkLocation = Request.QueryString("Location")
   strLinkStatus = Request.QueryString("Status")
   strLinkTech = Request.QueryString("Tech")
   strLinkCategory = Request.QueryString("Category")
   strLinkUser = Request.QueryString("User")
   strLinkFilter = Request.QueryString("Filter")
   strLinkProblem = Request.QueryString("Problem")
   strLinkNotes = Request.QueryString("Notes")
   strLinkEMail = Request.QueryString("EMail")
   strLinkSort = Request.QueryString("Sort")
   strLinkDays = Request.QueryString("Days")
   strLinkBack = Request.QueryString("Back")
   strLinkViewed = Request.QueryString("Viewed")

   If strLinkID <> "" Then
      BuildReturnLink = BuildReturnLink & "&ID=" & Replace(strLinkID," ","%20")
   End If   
   If strLinkLocation <> "" Then
      BuildReturnLink = BuildReturnLink & "&Location=" & Replace(strLinkLocation," ","%20")
   End If
   If strLinkStatus <> "" Then
      BuildReturnLink = BuildReturnLink & "&Status=" & Replace(strLinkStatus," ","%20")
   End If
   If strLinkTech <> "" Then
      BuildReturnLink = BuildReturnLink & "&Tech=" & Replace(strLinkTech," ","%20")
   End If
   If strLinkCategory <> "" Then
      BuildReturnLink = BuildReturnLink & "&Category=" & Replace(strLinkCategory," ","%20")
   End If
   If strLinkUser <> "" Then
      BuildReturnLink = BuildReturnLink & "&User=" & Replace(strLinkUser," ","%20")
   End If
   If strLinkFilter <> "" Then
      BuildReturnLink = BuildReturnLink & "&Filter=" & Replace(strLinkFilter," ","%20")
   End If
   If strLinkProblem <> "" Then
      BuildReturnLink = BuildReturnLink & "&Problem=" & Replace(strLinkProblem," ","%20")
   End If
   If strLinkNotes <> "" Then
      BuildReturnLink = BuildReturnLink & "&Notes=" & Replace(strLinkNotes," ","%20")
   End If
   If strLinkEMail <> "" Then
      BuildReturnLink = BuildReturnLink & "&EMail=" & Replace(strLinkEMail," ","%20")
   End If
   If strLinkSort <> "" Then
      BuildReturnLink = BuildReturnLink & "&Sort=" & Replace(strLinkSort," ","%20")
   End If
   If strLinkBack <> "" Then
      BuildReturnLink = BuildReturnLink & "&Back=Yes"
   End If
   If strLinkDays <> "" Then
      BuildReturnLink = BuildReturnLink & "&Days=" & strLinkDays
   End If
   If strLinkViewed <> "" Then
      BuildReturnLink = BuildReturnLink & "&Viewed=" & strLinkViewed
   End If
   
   If BuildReturnLink <> "" Then
      BuildReturnLink = "?" & Right(BuildReturnLink,(Len(BuildReturnLink)-1))
   End If

End Function
%>

<%
Function GetUser

   Const USERNAME = 1

   Dim strUserAgent, strSessionID, objSessionLookup
   
   'Get some needed data
   strSessionID = Request.Cookies("SessionID")
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
   
   'Send them to the logon screen if they don't have a Session ID
   If strSessionID = "" Then
      SendToLogonScreen

   'Get the username from the database
   Else
   
      strSQL = "SELECT ID,UserName,SessionID,IPAddress,UserAgent,ExpirationDate FROM Sessions "
      strSQL = strSQL & "WHERE UserAgent='" & Left(Replace(strUserAgent,"'","''"),250) & "' And SessionID='" & Replace(strSessionID,"'","''") & "'"
      strSQL = strSQL & " And ExpirationDate > Date()"
      Set objSessionLookup = Application("Connection").Execute(strSQL)
      
      'If a session isn't found for then kick them out
      If objSessionLookup.EOF Then
         SendToLogonScreen
      Else
         GetUser = objSessionLookup(USERNAME)
      End If
   End If  
   
End Function
%>

<%
Sub SendToLogonScreen

   Dim strReturnLink, strSourcePage
      
   'Build the return link before sending them away.
   strReturnLink = BuildReturnLink(True)
   strSourcePage = Request.ServerVariables("SCRIPT_NAME")
   strSourcePage = Right(strSourcePage,Len(strSourcePage) - InStrRev(strSourcePage,"/"))
   If strReturnLink = "" Then
      strReturnLink = "?SourcePage=" & strSourcePage
   Else
      strReturnLink = strReturnLink & "&SourcePage=" & strSourcePage
   End If
   
   Response.Redirect("login.asp" & strReturnLink)
   
End Sub 
%>

<%Sub VersionProblem %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
      <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
      <meta name="viewport" content="width=device-width" />
   </head>
   <body>
      <center><b>Web and Database versions don't match</b></center>
   </body>
   </html>

<%
   Response.End

End Sub%>