<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 1/6/12
'Last Updated 6/16/14

'This page allows you to post a message to the help desk site.

Option Explicit

On Error Resume Next

Dim objNetwork, strUser, strSQL, strRole, objNameCheckSet, strUserAgent, objFSO
Dim objThemesFolder, colThemes, objFolder, strTheme, objUserSettings, strCMD
Dim strNewTheme, strMessage, strNewMobileVersion, bolMobileVersion, bolShowLogout
Dim objSessions, strSelectedSessions, objSelectedSessions, intZoom

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

'Set the zoom level
If Request.Cookies("ZoomLevel") = "ZoomIn" Then
   If InStr(strUserAgent,"Silk") Then
      intZoom = 1.4
   Else
      intZoom = 1.9
   End If
End If

'Build the SQL string, this will check the userlevel of the user.
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"
Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)
bolMobileVersion = objNameCheckSet(4)

'See if the user has the rights to visit this page
If objNameCheckSet(2) Then

   'An error would be generated if the user has NTFS rights to get in but is not found
   'in the database.  In this case the user is denied access.
   If Err Then
      AccessDenied
   Else
      AccessGranted
   End If
Else
   AccessDenied
End If

Sub AccessGranted  

   Dim strSession

   strCMD = Request.Form("cmdSave")
   strNewTheme = Request.Form("Theme")
   strNewMobileVersion = Request.Form("MobileVersion")
   strSelectedSessions = Request.Form("chkSession")

   objSelectedSessions = Split(strSelectedSessions,",")
   
   Select Case strCMD
      Case "Save"
         strSQL = "UPDATE Tech" & vbCRLF
         strSQL = strSQL & "SET Theme='" & strNewTheme & "',MobileVersion=" & strNewMobileVersion & vbCRLF
         strSQL = strSQL & "WHERE UserName='" & strUser & "'"     
         Application("Connection").Execute(strSQL)
         strMessage = "Settings Updated"
      
      Case "Delete Selected"
         For Each strSession in objSelectedSessions
            strSQL = "DELETE FROM Sessions" & vbCRLF
            strSQL = strSQL & "Where ID = " & Trim(strSession)
            Application("Connection").Execute(strSQL)
         Next   
         GetUser
         
      Case "Delete All"
         strSQL = "DELETE FROM Sessions" & vbCRLF
         strSQL = strSQL & "Where Username = '" & strUser & "'"
         Application("Connection").Execute(strSQL) 
         GetUser

   End Select

   'Get the current theme
   strSQL = "SELECT Theme FROM Tech WHERE UserName='" & strUser & "'"
   Set objUserSettings = Application("Connection").Execute(strSQL)
   strTheme = objUserSettings(0)
   
   If IsNull(strTheme) Then
      strTheme = ""
   End If

   'Get the current version (Mobile or full)
   strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"
   Set objNameCheckSet = Application("Connection").Execute(strSQL)
   bolMobileVersion = objNameCheckSet(4)
   
   'Remove all old sessions
   strSQL = "DELETE FROM Sessions WHERE Date() > ExpirationDate"
   Application("Connection").Execute(strSQL)
   
   'Get the list of active sessions
   strSQL = "SELECT ID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate "
   strSQL = strSQL & "FROM Sessions WHERE Username='" & strUser & "' ORDER BY LoginDate DESC, LoginTime DESC"
   Set objSessions = Application("Connection").Execute(strSQL)
   
   If IsMobile Then
      MobileVersion
   Else
      MainVersion
   End If   
   
End Sub%>

<%
Function IsMobile

   'It's not mobile if the user is requesting the full site
   Select Case LCase(Request.QueryString("Site"))
      Case "full"
         IsMobile = False
         Response.Cookies("SiteVersion") = "Full"
         Exit Function
      Case "mobile"
         IsMobile = True
         Response.Cookies("SiteVersion") = "Mobile"
         Exit Function
   End Select
   
   'Choose the site based on the cookie
   Select Case LCase(Request.Cookies("SiteVersion"))
      Case "full"
         IsMobile = False
         Exit Function
      Case "mobile"
         IsMobile = True
         Exit Function
   End Select
   
   'It's not mobile if the mobile version is turned off.
   If Not bolMobileVersion Then
      IsMobile = False
      Exit Function
   End If
   
   'Check the user agent for signs they are on a mobile browser
   If InStr(strUserAgent,"iPhone") Then
      IsMobile = True
   ElseIf Instr(strUserAgent,"Android") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"Windows Phone") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"BlackBerry") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"Nintendo") Then
      IsMobile = True   
   ElseIf InStr(strUserAgent,"PlayStation Vita") Then
      IsMobile = True
   Else
      IsMobile = False
   End If 

End Function 
%>

<%Sub MainVersion %>

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
      <% If Application("UseTaskList") And objNameCheckSet(5) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="tasklist.asp">Tasks</a><font class="separator"> | </font></li>
      <% End If %>
		<% If Application("UseStats") Then %>
			<li class="topbar"><a class="linkbar" href="stats.asp">Stats</a><font class="separator"> | </font></li> 
      <% End If %>
      <% If Application("UseDocs") And objNameCheckSet(6) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="docs.asp">Docs</a><font class="separator"> | </font></li>
      <% End If %>
         <li class="topbar">Settings
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
  <div class="mainarea">
<% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
      <td><img class="mainareaimage"" src="../themes/<%=Application("Theme")%>/images/admin.gif"></td>
<% Else %>
      <td><img class="mainareaimage" src="../themes/<%=objNameCheckSet(3)%>/images/admin.gif"></td>
<% End If %>
      <div class="mainareatext">
   <center>
   <table width="100%">
      <form method="POST" action="settings.asp">
   		<tr>
   			<td>Pick your theme:</td>
   			<td>
               <select size="1" name="Theme">
                  <option value="">~~System Default~~</option>
         <%       'Populate the color scheme pulldown list
                  Set objFSO = CreateObject("Scripting.FileSystemObject")
                  Set objThemesFolder = objFSO.GetFolder(Application("ThemeLocation"))
                  Set colThemes = objThemesFolder.Subfolders

                  For Each objFolder in colThemes
                     If Trim(Ucase(objFolder.Name)) <> Trim(Ucase(strTheme)) Then %>
                        <option value="<%=objFolder.Name%>"><%=Replace(objFolder.Name,",,,","...")%></option>
         <%          Else %>
                        <option value="<%=objFolder.Name%>" selected="selected"><%=Replace(objFolder.Name,",,,","...")%></option>
         <%          End If
                  Next%>   
               </select>
            </td>
   			<td>
               
            </td>
         </tr>
         <tr><td colspan="3"><hr /></td></tr>
         <tr>
            <td colspan="2">
               <table width="100%">
                  <tr>
                     <td>
                        Use mobile version?
                     </td>
                     <td>
                        <select size="1" name="MobileVersion">
                        <% If objNameCheckSet(4) Then %>
                           <option value="True" selected="selected">Yes</option> 
                           <option value="False">No</option> 
                        <% Else %>
                           <option value="True">Yes</option> 
                           <option value="False" selected="selected">No</option> 
                        <% End If %>
                        </select>
                     </td>
                     <td>
                        <input type="submit" value="Save" name="cmdSave" style="float: right">
                     </td>
                  </tr>
               </table>
            </td>
         </tr>
         <tr><td colspan="3"><hr /></td></tr>
   <% If Not objSessions.EOF Then %>
         <tr><td colspan="3" align="center">Active Sessions</td></tr>
         <tr><td colspan="3" align="center">
            <table>
               <tr>
                  <th>&nbsp;</th>
                  <th>&nbsp;Date&nbsp;</th>
                  <th>&nbsp;Time&nbsp;</th>
                  <th>&nbsp;IP Address&nbsp;</th>
                  <th>&nbsp;Device Type&nbsp;</th>
               </tr>
         <% Do Until objSessions.EOF 'ID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate
         %>
               <tr>
                  <td class="showborders" align="center" width="1%">
                     <input type="checkbox" name="chkSession" value="<%=objSessions(0)%>">
                  </td>
                  <td class="showborders">&nbsp;<%=Left(objSessions(3), Len(objSessions(3)) - 5)%>&nbsp;</td>
                  <td class="showborders">&nbsp;<%=Left(objSessions(4), Len(objSessions(4)) - 6) & Right(objSessions(4),2)%>&nbsp;</td>
                  <td class="showborders">&nbsp;<%=objSessions(1)%>&nbsp;</td>
                  <td class="showborders">&nbsp;<%=GetDeviceType(objSessions(2))%>&nbsp;</td>
               </tr>
         <%    objSessions.MoveNext
            Loop %>
               <tr><td colspan="5">
                  <input type="submit" value="Delete Selected" name="cmdSave" style="float: right">
                  <input type="submit" value="Delete All" name="cmdSave" style="float: right">
               </td></tr>
            </table>
         </td></tr>
   <% End If %>
   	</form>	
<%    'If a task was performed display the results
      If strMessage <> "" Then %>
         <tr><td colspan="3">
            <font class="information"><%=strMessage%></font></center>
         </td></tr>
<%    End If%>
      
   </table>
   </body>
   </html>

<%End Sub%>

<%Sub MobileVersion %>
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
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>" />
   <% If Request.Cookies("ZoomLevel") = "ZoomIn" Then %>
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %>
   <meta name="theme-color" content="#<%=Application("AndroidBarColor")%>">
      
   </head>
   <body>
      <center><b><%=Application("SchoolName")%> Help Desk Admin</b></center>
      <center>
      <table align="center">
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>               
         <form method="Post" action="view.asp">
         <tr>
            <td colspan="2">
               <div align="center">
                  <input type="submit" value="Home" name="home">
                  <input type="submit" value="Open Tickets" name="filter">
            <% If strRole <> "Data Viewer" Then %>   
                  <input type="submit" value="Your Tickets" name="filter">
            <% End If %>
               </div>
            </td>
         </tr>
         </form>
         <tr><td colspan="2"><hr /></td></tr>


         <form method="POST" action="settings.asp">
   		<tr>
   			<td>Pick your theme:</td>
   			<td>
               <select size="1" name="Theme">
                  <option value="">~~System Default~~</option>
         <%       'Populate the color scheme pulldown list
                  Set objFSO = CreateObject("Scripting.FileSystemObject")
                  Set objThemesFolder = objFSO.GetFolder(Application("ThemeLocation"))
                  Set colThemes = objThemesFolder.Subfolders

                  For Each objFolder in colThemes
                     If Trim(Ucase(objFolder.Name)) <> Trim(Ucase(strTheme)) Then %>
                        <option value="<%=objFolder.Name%>"><%=Replace(objFolder.Name,",,,","...")%></option>
         <%          Else %>
                        <option value="<%=objFolder.Name%>" selected="selected"><%=Replace(objFolder.Name,",,,","...")%></option>
         <%          End If
                  Next%>   
               </select>
            </td>
   			<td>
               
            </td>
         </tr>
         <tr><td colspan="2"><hr /></td></tr>
         <tr>
            <td>
               Use mobile version?
            </td>
            <td>
               <select size="1" name="MobileVersion">
               <% If objNameCheckSet(4) Then %>
                  <option value="True" selected="selected">Yes</option> 
                  <option value="False">No</option> 
               <% Else %>
                  <option value="True">Yes</option> 
                  <option value="False" selected="selected">No</option> 
               <% End If %>
               </select>
            </td>
         </tr>
         <tr><td colspan="2"><hr /></td></tr>
         <tr>
            <td colspan="2">
   <%    'If a task was performed display the results
         If strMessage <> "" Then %>
               <font class="information"><%=strMessage%></font></center>
   <%    End If%>
               <input type="submit" value="Save" name="cmdSave" style="float: right">
            </td>
         </tr>
   <% If Not objSessions.EOF Then %>
         <tr><td colspan="3" align="center"><hr /></td></tr>
         <tr><td colspan="3" align="center">Active Sessions</td></tr>
         <tr><td colspan="3" align="center">
            <table>
               <tr>
                  <th>&nbsp;</th>
                  <th>&nbsp;Date&nbsp;</th>
                  <th>&nbsp;Time&nbsp;</th>
                  <th>&nbsp;Device Type&nbsp;</th>
               </tr>
         <% Do Until objSessions.EOF 'ID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate
         %>
               <tr>
                  <td class="showborders" align="center" width="1%">
                     <input type="checkbox" name="chkSession" value="<%=objSessions(0)%>">
                  </td>
                  <td class="showborders">&nbsp;<%=Left(objSessions(3), Len(objSessions(3)) - 5)%>&nbsp;</td>
                  <td class="showborders">&nbsp;<%=Left(objSessions(4), Len(objSessions(4)) - 6) & Right(objSessions(4),2)%>&nbsp;</td>
                  <td class="showborders">&nbsp;<%=GetDeviceType(objSessions(2))%>&nbsp;</td>
               </tr>
         <%    objSessions.MoveNext
            Loop %>
               <tr><td colspan="5">
                  <input type="submit" value="Delete Selected" name="cmdSave" style="float: right">
                  <input type="submit" value="Delete All" name="cmdSave" style="float: right">
               </td></tr>
            </table>
         </td></tr>
   <% End If %>

         </form>
      </table>
   </body>
   </html>
      
<%End Sub%>

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
Function GetDeviceType(strDevice)
   
   'Check the user agent to determine device type
   
   'iOS
   If InStr(strDevice,"iPhone") > 0 And InStr(strDevice,"CriOS") > 0 Then
      GetDeviceType = "iPhone - Chrome"
   ElseIf InStr(strDevice,"iPhone") > 0 And InStr(strDevice,"CriOS") > 0 Then
      GetDeviceType = "iPhone - Safari"
   ElseIf InStr(strDevice,"iPhone") Then
      GetDeviceType = "iPhone"
   ElseIf InStr(strDevice,"iPad") > 0 And InStr(strDevice,"CriOS") > 0 Then
      GetDeviceType = "iPad - Chrome"
   ElseIf InStr(strDevice,"iPad") > 0 And InStr(strDevice,"Safari") > 0 Then
      GetDeviceType = "iPad - Safari"
   ElseIf InStr(strDevice,"iPad") Then
      GetDeviceType = "iPad"
   
   'Android
   ElseIf InStr(strDevice,"Nexus 5") Then
      GetDeviceType = "Nexus 5"
   ElseIf InStr(strDevice,"Nexus 6") Then
      GetDeviceType = "Nexus 6"
   ElseIf InStr(strDevice,"Nexus 7") Then
      GetDeviceType = "Nexus 7"
   ElseIf InStr(strDevice,"Nexus 9") Then
      GetDeviceType = "Nexus 9"
   ElseIf InStr(strDevice,"Kindle") Then
      GetDeviceType = "Kindle"
   ElseIf InStr(strDevice,"Silk") Then
      GetDeviceType = "Amazon Fire"
   ElseIf InStr(strDevice,"GT-N5110") Then
      GetDeviceType = "Galaxy Note 8"
   ElseIf Instr(strDevice,"Android") > 0 And InStr(strDevice,"Watch") > 0 Then
      GetDeviceType = "Android Watch"   
   ElseIf Instr(strDevice,"Android") Then
      GetDeviceType = "Android Device"
   
   'Windows 7
   ElseIf InStr(strDevice, "Windows NT 6.1") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Win 7/2008 R2 - Chrome"
   ElseIf InStr(strDevice, "Windows NT 6.1") > 0 And InStr(strDevice, "Xbox") > 0 Then
      GetDeviceType = "Xbox 360"
   ElseIf InStr(strDevice, "Windows NT 6.1") > 0 And InStr(strDevice, "MSIE") > 0 Then
      GetDeviceType = "Win 7/2008 R2 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.1") > 0 And InStr(strDevice, "Trident") > 0 Then
      GetDeviceType = "Win 7/2008 R2 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.1") Then
      GetDeviceType = "Win 7"
   
   'Windows 8.1
   ElseIf InStr(strDevice, "Windows NT 6.3") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Win 8.1/2012 R2 - Chrome"
   ElseIf InStr(strDevice, "Windows NT 6.3") > 0 And InStr(strDevice, "MSIE") > 0 Then
      GetDeviceType = "Win 8.1/2012 R2 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.3") > 0 And InStr(strDevice, "Trident") > 0 Then
      GetDeviceType = "Win 8.1/2012 R2 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.3") Then
      GetDeviceType = "Win 8.1/2012 R2"   

   'Windows XP   
   ElseIf InStr(strDevice, "Windows NT 5.1") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Win XP/2003 - Chrome"
   ElseIf InStr(strDevice, "Windows NT 5.1") > 0 And InStr(strDevice, "MSIE") > 0 Then
      GetDeviceType = "Win XP/2003 - IE"
   ElseIf InStr(strDevice, "Windows NT 5.1") > 0 And InStr(strDevice, "Trident") > 0 Then
      GetDeviceType = "Win XP/2003 - IE"
   ElseIf InStr(strDevice, "Windows NT 5.1") Then
      GetDeviceType = "Win XP/2003"
   ElseIf InStr(strDevice, "Windows NT 5.2") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Win XP/2003 - Chrome"
   ElseIf InStr(strDevice, "Windows NT 5.2") > 0 And InStr(strDevice, "MSIE") > 0 Then
      GetDeviceType = "Win XP/2003 - IE"
   ElseIf InStr(strDevice, "Windows NT 5.2") > 0 And InStr(strDevice, "Trident") > 0 Then
      GetDeviceType = "Win XP/2003 - IE"
   ElseIf InStr(strDevice, "Windows NT 5.2") Then
      GetDeviceType = "Win XP/2003" 
      
   'Windows 8
   ElseIf InStr(strDevice, "Windows NT 6.2") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Win 8/2012 - Chrome"
   ElseIf InStr(strDevice, "Windows NT 6.2") > 0 And InStr(strDevice, "MSIE") > 0 Then
      GetDeviceType = "Win 8/2012 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.2") > 0 And InStr(strDevice, "Trident") > 0 Then
      GetDeviceType = "Win 8/2012 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.2") Then
      GetDeviceType = "Win 8/2012"
   
   'Windows Vista
   ElseIf InStr(strDevice, "Windows NT 6.0") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Win Vista/2008 - Chrome"
   ElseIf InStr(strDevice, "Windows NT 6.0") > 0 And InStr(strDevice, "MSIE") > 0 Then
      GetDeviceType = "Win Vista/2008 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.0") > 0 And InStr(strDevice, "Trident") > 0 Then
      GetDeviceType = "Win Vista/2008 - IE"
   ElseIf InStr(strDevice, "Windows NT 6.0") Then
      GetDeviceType = "Win Vista/2008"
   
   'Other Windows
   ElseIf InStr(strDevice,"Windows Phone 8") Then
      GetDeviceType = "Windows Phone 8"
   ElseIf InStr(strDevice,"Windows Phone 7") Then
      GetDeviceType = "Windows Phone 7"
   ElseIf InStr(strDevice,"Windows Phone") Then
      GetDeviceType = "Windows Phone"
   ElseIf Instr(strDevice, "Windows") Then
      GetDeviceType = "Windows PC"
   
   'Macintosh
   ElseIf Instr(strDevice, "Macintosh") > 0 And InStr(strDevice, "Chrome") > 0 Then
      GetDeviceType = "Mac - Chrome"
   ElseIf Instr(strDevice, "Macintosh") > 0 And InStr(strDevice, "Safari") > 0 Then
      GetDeviceType = "Mac - Safari"   
   ElseIf Instr(strDevice, "Macintosh") Then
      GetDeviceType = "Macintosh"
   
   'Nintendo
   ElseIf InStr(strDevice,"Nintendo WiiU") Then
      GetDeviceType = "Nintendo Wii u"
   ElseIf InStr(strDevice,"Nintendo 3DS") Then
      GetDeviceType = "Nintendo 3DS"
   ElseIf InStr(strDevice, "Nintendo DSi") Then
      GetDeviceType = "Nintendo DSi"
   ElseIf InStr(strDevice,"Nintendo") Then
      GetDeviceType = "Nintendo"
   
   'Playstation
   ElseIf InStr(strDevice,"PlayStation Vita") Then
      GetDeviceType = "Playstation Vita"
   ElseIf InStr(strDevice,"PLAYSTATION 3") Then
      GetDeviceType = "Playstation 3"
   ElseIf Instr(strDevice,"PlayStation 4") Then
      GetDeviceType = "Playstation 4"
   
   'Other Devices
   ElseIf InStr(strDevice,"CrOS") Then
      GetDeviceType = "ChromeBook"
   ElseIf InStr(strDevice,"BlackBerry") Then
      GetDeviceType = "BlackBerry"
   ElseIf InStr(strDevice,"Linux") Then
      GetDeviceType = "Linux"
   Else
      GetDeviceType = "Unknown"
   End If 
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