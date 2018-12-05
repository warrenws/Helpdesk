<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/23/04
'Last Updated 6/16/14

'This page will allow the help desk administrators change some of the settings.  At this
'point once the settings are saved the World Wide Web Service needs to be restarted.  This
'can easily be done by running IISReset from the command prompt on the server.

Option Explicit

On Error Resume Next

Dim strSQL, objRecordSet, objColorSet, strColorScheme, strSendMailTo, strSendMailFrom
Dim strBlindCopyTo, strAdminURL, strSchoolName, strEMailSuffix, strMainPageText
Dim strMessage, strBGColor, strTxtColor, strLnkColor, strInfoColor, strBlindCopyToTemp
Dim strWarningColor, objRegExp, strSendMailToTemp, strSendMailFromTemp, objNetwork
Dim strAdminURLTemp, strSchoolNameTemp, strEMailSuffixTemp, strMainPageTextTemp, strUser
Dim objNameCheckSet, strUserCanViewCallStatus, strSMTPPickupFolder, strCustom1, strCustom2
Dim strUseAD, strUseCustom1, strUseCustom2, strIconLocation, objFSO, objThemesFolder
Dim strTheme, colThemes, objFolder, strUserAgent, strRole, strUseTaskList, strUseDocumentation
Dim strUseStats, strUseUpload, strDomainController, strADUsername, strADPassword
Dim bolShowLogout, strShowUserStats, strShowUserButtons, strSendReminder

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

'Build the SQL string
strSQL = "Select Username, UserLevel, Active, Theme" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"

Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)

'See if the user has the rights to visit this page
If objNameCheckSet(1) = "Administrator" AND objNameCheckSet(2) Then

   'An error would be generated if the user has NTFS rights to get in but is not found
   'in the database.  In this case the user is denied access.
   If Err Then
      Err.Clear
      Call AccessDenied
   Else
      Call AccessGranted
   End If
Else
   Call AccessDenied
End If

Sub AccessGranted     

   'The user has a userlevel of administrator and is allowed to access this site.

   On Error Resume Next
   
   'Build the SQL string to get the current settings from the database
   strSQL = "Select SendMailTo,SendMailFrom,BBC,AdminURL,SchoolName,EMailSuffix,MainPageText,UserCanViewCallStatus,SMTPPickupFolder,UseAD,Custom1Text,Custom2Text,IconLocation,Theme,UseTaskList,UseDocumentation,UseStats,UseUpload,ADUsername,ADPassword,DomainController,ShowUserStats,ShowUserButtons,SendReminder" & vbCRLF
   strSQL = strSQL & "From Settings;"     
   
   'Execute the SQL query to get the current settings from the database and assign them to the 
   'Record Set
   Set objRecordSet = Application("Connection").Execute(strSQL) 
   
   'Check and see if the form was already submitted.  If so the write the new settings to
   'the database, if not get the current settings from the database.
   If (Request("cmdSave") = "Save") Then
      strSendMailTo = Request.Form("SendMailTo")
      strSendMailFrom = Request.Form("SendMailFrom")
      strBlindCopyTo = Request.Form("BlindCopyTo")
      strAdminURL = Request.Form("AdminURL")   
      strSchoolName = Request.Form("SchoolName")
      strEMailSuffix = Request.Form("EMailSuffix")
      strMainPageText = Request.Form("MainPageText")
      strTheme = Request.Form("Theme")
      strUserCanViewCallStatus = Request.Form("UserCanViewCallStatus")
      strSMTPPickupFolder = Request.Form("SMTPPickupFolder")
	   strUseAD = Request.Form("UseAD")
	   strCustom1 = Request.Form("Custom1")
	   strCustom2 = Request.Form("Custom2")
      strIconLocation = Request.Form("IconLocation")
      strUseTaskList = Request.Form("UseTaskList")
      strUseDocumentation = Request.Form("UseDocumentation")
      strUseStats = Request.Form("UseStats")
      strUseUpload = Request.Form("UseUpload")
      strADUsername = Request.Form("ADUsername")
      strADPassword = Request.Form("ADPassword")
      strDomainController = Request.Form("DomainController")
      strShowUserStats = Request.Form("ShowUserStats")
      strShowUserButtons = Request.Form("ShowUserButtons")
      strSendReminder = Request.Form("SendReminder")
      strMessage = "Settings Saved"
      
      'Set strUserCanViewCallStatus to True or False
      If strUserCanViewCallStatus = "Yes" Then
         strUserCanViewCallStatus = True
      Else
         strUserCanViewCallStatus = False
      End If
      
      'Create the Regular Expression object and set it's properties.
      Set objRegExp = New RegExp
      objRegExp.Pattern = "'"
      objRegExp.Global = True
      
      'Use the regular expression to change a ' to a '' so the SQL Insert command will work.
      'The value will be assigned to a new variable so the old one can still be displayed
      strSendMailToTemp = objRegExp.Replace(strSendMailTo,"''")
      strSendMailFromTemp = objRegExp.Replace(strSendMailFrom,"''")
      strBlindCopyToTemp = objRegExp.Replace(strBlindCopyTo,"''")
      strAdminURLTemp = objRegExp.Replace(strAdminURL,"''")
      strSchoolNameTemp = objRegExp.Replace(strSchoolName,"''")
      strEMailSuffixTemp = objRegExp.Replace(strEMailSuffix,"''")
      strMainPageTextTemp = objRegExp.Replace(strMainPageText,"''")
      strSMTPPickupFolder = objRegExp.Replace(strSMTPPickupFolder,"''")
	   strCustom1 = objRegExp.Replace(strCustom1,"''")
	   strCustom2 = objRegExp.Replace(strCustom2,"''")
      strADUserName = objRegExp.Replace(strADUserName,"''")
      strADPassword = objRegExp.Replace(strADPassword,"''")
      strDomainController = objRegExp.Replace(strDomainController,"''")
	  
      'Turn on or off Custom1 and Custom2
      If strCustom1 = "" Then
        strUseCustom1 = False
      Else 
        strUseCustom1 = True
      End If

      If strCustom2 = "" Then
        strUseCustom2 = False
      Else 
        strUseCustom2 = True
      End If
     
      'Build the SQL string that will insert the new data into the database
      strSQL = "Update Settings" & vbCRLF
      strSQL = strSQL & "Set SendMailTo = '" & strSendMailToTemp & "',SendMailFrom = '" & strSendMailFromTemp & "',BBC = '" & _
      strBlindCopyToTemp & "',AdminURL = '" & strAdminURLTemp & "',SchoolName = '" & strSchoolNameTemp & "',EMailSuffix = '" & strEMailSuffixTemp & _
      "',MainPageText = '" & strMainPageTextTemp & "',UserCanViewCallStatus = " & strUserCanViewCallStatus & _
      ",SMTPPickupFolder ='" & strSMTPPickupFolder & "',Custom1Text ='" & strCustom1 & "',Custom2Text = '" & strCustom2 & "'" & _
      ",UseCustom1 = " & strUseCustom1 & ",UseCustom2 = " & strUseCustom2 & ",UseAD = " & strUseAD & ",IconLocation = '" & strIconLocation & "',Theme = '" & strTheme & "'"& vbCRLF & _
      ",UseTaskList = " & strUseTaskList & ",UseDocumentation = " & strUseDocumentation & ",UseStats = " & strUseStats & ",UseUpload = " & strUseUpload & _
      ",ADUserName = '" & strADUserName & "',ADPassword = '" & strADPassword & "',DomainController = '" & strDomainController & "'" & _
      ",ShowUserStats = " & strShowUserStats & ",ShowUserButtons = " & strShowUserButtons & ",SendReminder = " & strSendReminder & vbCRLF
      strSQL = strSQL & "WHERE (((Settings.ID)=1));"

      'Execute the query
      Application("Connection").Execute(strSQL)
      
      'Run the main sub
      Call Main
   Else
      'Get the current settings from the database
      strSendMailTo = objRecordSet(0)
      strSendMailFrom = objRecordSet(1)
      strBlindCopyTo = objRecordSet(2)
      strAdminURL = objRecordSet(3)   
      strSchoolName = objRecordSet(4)
      strEMailSuffix = objRecordSet(5)
      strMainPageText = objRecordSet(6)
      strUserCanViewCallStatus = objRecordSet(7)
      strSMTPPickupFolder = objRecordSet(8)
	   strUseAD = objRecordSet(9)
	   strCustom1 = objRecordSet(10)
	   strCustom2 = objRecordSet(11)
      strIconLocation = objRecordSet(12)
      strTheme = objRecordSet(13)
      strUseTaskList = objRecordSet(14)
      strUseDocumentation = objRecordSet(15)
      strUseStats = objRecordSet(16)    
      strUseUpload = objRecordSet(17)
      strADUserName = objRecordSet(18)
      strADPassword = objRecordSet(19)
      strDomainController = objRecordSet(20)
      strShowUserStats = objRecordSet(21)
      strShowUserButtons = objRecordSet(22)
      strSendReminder = objRecordSet(23)
      
      'Run the main sub
      Call Main
   End If
End Sub

Sub Main

   Dim intInputSize, intTextBoxSize

   If InStr(strUserAgent,"Android") Or InStr(strUserAgent,"Silk") Then
      intInputSize = 58
      intTextBoxSize = 59
   Else
      intInputSize = 85
      intTextBoxSize = 64
   End If
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
      <%=Application("SchoolName")%> Help Desk Admin
   </div>
   
   <div class="version">
      Version <%=Application("Version")%>
   </div>
   
   <hr class="admintopbar" />
   <div class="admintopbar">
      <ul class="topbar">
			<li class="topbar">Setup<font class="separator"> | </font></li> 
         <li class="topbar"><a href="users.asp">Users</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="messages.asp">Messages</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="dbtools.asp">Database Tools</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="index.asp">User Mode</a>
      <% If bolShowLogout Then %>
         <font class="separator"> | </font></li>
         <li class="topbar"><a href="login.asp?action=logout">Log Out</a></li>
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

   <div align="center">
   
   <table border="0" width="750" id="table1" cellspacing="0" cellpadding="0">
   	<tr>
   		<td width="22%" valign="top">
      <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>      
   		<img border="0" src="../themes/<%=Application("Theme")%>/images/admin.gif"><p>
      <% Else %>
         <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/admin.gif"><p>
      <% End If %>

   <%    'Display a message when the form is submitted
         If strMessage <> "" Then %>
            <p align="center"><font class="information"><%=strMessage%></font></p>
   <%    Else%>
   	      &nbsp;</p>
   <%    End If%>
   
         </td>
   		<td width="2%" valign="top">&nbsp;</td>
   		<td width="71%" valign="top">
   		<p style="margin-top: 0; margin-bottom: 0">
   		Modify different help desk settings.</p>
   		<hr />
   	<table border="0" width="99%" id="table5" cellspacing="0" cellpadding="0">
   	   <form method="POST" action="setup.asp">
   		<tr>
   			<td height="19" colspan="3">
			E-Mail: Help Desk ticket recipients, separate each value with a semicolon.</td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3">
   			<input type="text" name="SendMailTo" value="<%=strSendMailTo%>" size="<%=intInputSize%>">
   			</td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3"><hr /></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3">Admin E-Mail: All mail will come from this e-mail address.</td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3">
   			<input type="text" name="SendMailFrom" value="<%=strSendMailFrom%>" size="40"></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3"><hr /></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3">Blind Copy To:&nbsp; A blind copy of all mail goes to this email address.</td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3"> 
   			<input type="text" name="BlindCopyTo" value="<%=strBlindCopyTo%>" size="40"></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3"><hr /></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3">Admin URL:&nbsp; The admin section address, used 
			in email to helpdesk admins.</td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3"> 
   			<input type="text" name="AdminURL" value="<%=strAdminURL%>" size="<%=intInputSize%>"></td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3"><hr /></td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3">
			Organization Name:&nbsp; The name at the top of the screen.</td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3">
   			<input type="text" name="SchoolName" value="<%=strSchoolName%>" size="40"></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3"><hr></td>
   		</tr>
   		<tr>
   			<td height="19" colspan="3">E-Mail Suffix:&nbsp; In the format @domain.com</td>
   		</tr>
   		<tr>
   			<td height="22" colspan="3"> 
   			<input type="text" name="EMailSuffix" value="<%=strEMailSuffix%>" size="40"></td>
   		</tr>
   		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3"><hr /></td>
   		</tr>
   		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
			SMTP Pickup Folder (Default = C:\Inetpub\mailroot\Pickup) 
			</td>
   		</tr>
   		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
   			<input type="text" name="SMTPPickupFolder" value="<%=strSMTPPickupFolder%>" size="40"></td>
   		</tr>
         <tr>
   			<td width="99%" height="19" valign="bottom" colspan="3"><hr /></td>
   		</tr>
         <tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
			Icon Location (Default = https://helpdesk.lkgeorge.org/icons)</td>
   		</tr>
   		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
   			<input type="text" name="IconLocation" value="<%=strIconLocation%>" size="<%=intInputSize%>"></td>
   		</tr>
   		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3"><hr></td>
   		</tr>
   		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">Main Page Text:&nbsp; The text seen when a user 
			goes to the help desk page.</td>
   		</tr>
   		<tr>
   			<td width="99%" colspan="3">
			<textarea rows="6" name="MainPageText" cols="<%=intTextBoxSize%>"><%=strMainPageText%></textarea></td>
   		</tr>
   		<tr>
   			<td width="99%" height="16" colspan="3"><hr /></td>
   		</tr>
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Allow the user to view their tickets?
   		<select size="1" name="UserCanViewCallStatus">

<%       'Set the value from the database as the default in the pulldown list
         If strUserCanViewCallStatus = True Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>
         
         
   		<tr>
   			<td height="29" colspan="3">
   			Allow the user to track, request updates, or close their open tickets?
   		<select size="1" name="ShowUserButtons">

<%       'Set the value from the database as the default in the pulldown list
         If strShowUserButtons = True Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			
   		</tr>
         
         
   		<tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Do you want to use Active Directory to attempt to  </br>
			look up the user's information?&nbsp;&nbsp;
   		<select size="1" name="UseAD">

<%       'Set the value from the database as the default in the pulldown list
         If strUseAD = True or strUseAD = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>
         <tr>
            <td colspan="3">Domain Controller</td>
         </tr>
         <tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
   			<input type="text" name="DomainController" value="<%=strDomainController%>" size="40"></td>
   		</tr>
         <tr>
            <td colspan="3">User name (domain\username)</td>
         </tr>
         <tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
   			<input type="text" name="ADUsername" value="<%=strADUsername%>" size="40"></td>
   		</tr>
         <tr>
            <td colspan="3">Password</td>
         </tr> 
         <tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
   			<input type="password" name="ADPassword" value="<%=strADPassword%>" size="40"></td>
   		</tr>
         <tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Do you want to display the administrator's statistics?
   		<select size="1" name="UseStats">

<%       'Set the value from the database as the default in the pulldown list
         If strUseStats = True or strUseStats = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>

   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Do you want to display the user's statistics?
   		<select size="1" name="ShowUserStats">

<%       'Set the value from the database as the default in the pulldown list
         If strShowUserStats = True or strShowUserStats = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>         
         
         
         <tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Enable the task list feature?
   		<select size="1" name="UseTaskList">

<%       'Set the value from the database as the default in the pulldown list
         If strUseTaskList = True or strUseTaskList = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>  

         <tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Enable the documentation feature?
   		<select size="1" name="UseDocumentation">

<%       'Set the value from the database as the default in the pulldown list
         If strUseDocumentation = True or strUseDocumentation = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>          

   		<tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
         
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Allow users to upload files?
   		<select size="1" name="UseUpload">

<%       'Set the value from the database as the default in the pulldown list
         If strUseUpload = True or strUseUpload = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>  
         
   		<tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>       
   		<tr>
   			<td width="82%" height="29" colspan="2">
   			Send a reminder email if someone else enters a ticket for a user?
   		<select size="1" name="SendReminder">

<%       'Set the value from the database as the default in the pulldown list
         If strSendReminder = True or strSendReminder = "Yes" Then %>   		
            <option>Yes</option>
			   <option>No</option>
<%       Else %>
			   <option>No</option>
			   <option>Yes</option>
<%       End If %>

			</select></td>
   			<td width="18%" height="29">&nbsp;</td>
   		</tr>  
         
   		<tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
		<tr>
   			<td width="99%" height="19" valign="bottom" colspan="3">
			Below you can add up to two additional fields that will be mandatory.
		   </td>
   		</tr>
		<tr>
   			<td width="45%" colspan="1">
			<input type="text" name="Custom1" value="<%=strCustom1%>" size="15">
			<td width="45%" colspan="1">
			<input type="text" name="Custom2" value="<%=strCustom2%>" size="15">
   		</tr>
   		<tr>
   			<td width="99%" height="16" colspan="3"><hr></td>
   		</tr>
		
   		<tr>
   			<td width="21%" height="22">Choose the default theme:</td>
   			<td width="61%" height="22">

			<select size="1" name="Theme">
         
<%       'Populate the color scheme pulldown list
         Set objFSO = CreateObject("Scripting.FileSystemObject")
         Set objThemesFolder = objFSO.GetFolder(Application("ThemeLocation"))
         Set colThemes = objThemesFolder.Subfolders

         For Each objFolder in colThemes
            If Trim(Ucase(objFolder.Name)) <> Trim(Ucase(strTheme)) Then %>
               <option value="<%=objFolder.Name%>"><%=objFolder.Name%></option>
<%          Else %>
               <option value="<%=objFolder.Name%>" selected="selected"><%=objFolder.Name%></option>
<%          End If
         Next%>
         
			</select></td>
   			<td width="18%" height="22">
   			<input type="submit" value="Save" name="cmdSave" style="float: right"></td>
   		</tr>
   		<tr>
   			<td width="99%" colspan="3"><hr /></td>
   		</tr>
   	</table>
   </form>   
   	<p style="margin-top: 0; margin-bottom: 0">   
   		</td>
   	</tr>
   	<tr>
   		<td width="22%" valign="top" height="23">
   		&nbsp;</td>
   		<td width="2%" valign="top" height="23">&nbsp;</td>
   		<td width="71%" valign="top" height="23">
   		<p>&nbsp;</td>
   	</tr>
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