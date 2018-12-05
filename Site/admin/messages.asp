<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 12/28/11
'Last Updated 6/16/14

'This page allows you to post a message to the help desk site.

Option Explicit

On Error Resume Next

Dim objNetwork, strUser, strSQL, strRole, objNameCheckSet, strUserAgent, strEMailSubject
Dim strEMailMessage, strEMailTitle, bolShowLogout

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

'Build the SQL string, this will check the userlevel of the user.
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

   Dim objMessage, strOldMessage, strOldRecipient, strOldType, strOldEnabled
   Dim strNewMessage, strNewRecipient, strNewType, strNewEnabled
   Dim strMessage, strRecipient, strType, strEnabled, strAlertSelected
   Dim strTechsSelected, strUsersSelected, strBothSelected, strNormalSelected
   Dim strEnabledYes, strEnabledNo, strMessagetoUser, strTopSelected
   Dim strBottomSelected, strBothPositionsSelected, strOldPosition
   Dim strNewPosition, strPosition, strMessageName, objEMail
   
   strSQL = "SELECT Message,Recipient,Type,PositionOnPage,Enabled" & vbCRLF
   strSQL = strSQL & "FROM Message" & vbCRLF
   strSQL = strSQL & "WHERE ID=1"
   
   Set objMessage = Application("Connection").Execute(strSQL)
   
   strOldMessage = objMessage(0)
   strOldRecipient = objMessage(1)
   strOldType = objMessage(2)
   strOldPosition = objMessage(3)
   strOldEnabled = objMessage(4)
   
   strNewMessage = Request.Form("Message")
   strNewRecipient = Request.Form("Recipient")
   strNewType = Request.Form("Type")
   strNewPosition = Request.Form("Position")
   strNewEnabled = Request.Form("Enable")

   If (strOldMessage <> strNewMessage) And (strNewMessage <> "") Then
      strMessage = strNewMessage
   Else   
      strMessage = strOldMessage
   End If
   
   If (strOldRecipient <> strNewRecipient) And (strNewRecipient <> "") Then
      strRecipient = strNewRecipient
   Else   
      strRecipient = strOldRecipient
   End If
   
   If (strOldType <> strNewType) And (strNewType <> "") Then
      strType = strNewType
   Else   
      strType = strOldType
   End If   
   
   If (strOldPosition <> strNewPosition) And (strNewPosition <> "") Then
      strPosition = strNewPosition
   Else   
      strPosition = strOldPosition
   End If   
   
   If (strOldEnabled <> strNewEnabled) And (strNewEnabled <> "") Then
      strEnabled = strNewEnabled
   Else   
      strEnabled = strOldEnabled
   End If  
 
   Select Case strRecipient
      Case "Techs"
         strTechsSelected = "selected=""selected"""
      Case "Users"
         strUsersSelected = "selected=""selected"""
      Case "Both"
         strBothSelected = "selected=""selected"""
   End Select
   
   Select Case strType
      Case "Normal"
         strNormalSelected = "selected=""selected"""
      Case "Alert"
         strAlertSelected = "selected=""selected"""
   End Select
   
   Select Case strPosition
      Case "Top"
         strTopSelected = "selected=""selected"""
      Case "Bottom"
         strBottomSelected = "selected=""selected"""
      Case "Both"
         strBothPositionsSelected = "selected=""selected"""
   End Select
   
   If strEnabled = "True" Then
      strEnabledYes = "selected=""selected"""
   Else
      strEnabledNo = "selected=""selected"""
   End If
   
   If Request.Form("cmdsubmit") = "Save" Then
      strMessagetoUser = "Message Updated"
      
      strSQL = "UPDATE Message" & vbCRLF
      strSQL = strSQL & "SET Message='" & Replace(strMessage,"'","''") & "',"
      strSQL = strSQL & "Recipient='" & strRecipient & "',"
      strSQL = strSQL & "Type='" & strType & "',"
      strSQL = strSQL & "PositionOnPage='" & strPosition & "',"
      strSQL = strSQL & "Enabled=" & strEnabled & vbCRLF
      strSQL = strSQL & "WHERE ID=1"

      Application("Connection").Execute(strSQL)
   End If 
   
   If Request.Form("cmdEMail") = "Save" Then
      strEMailTitle = Request.Form("Title")
      strEMailSubject = Request.Form("Subject")
      strEMailMessage = Request.Form("EMailMessage")
      
      If strEMailSubject <> "" And strEMailMessage <> "" Then
         strSQL = "UPDATE EMail" & vbCRLF
         strSQL = strSQL & "SET Subject='" & Replace(strEMailSubject,"'","''") & "',Message='" & Replace(strEMailMessage,"'","''") & "'" & vbCRLF
         strSQL = strSQL & "WHERE Title='" & strEMailTitle & "'"
         Application("Connection").Execute(strSQL)
         strMessagetoUser = "EMail Updated"
      End If
   End If
   
   strMessageName = Request.Form("MessageName")
   If strMessageName <> "" Then
      strSQL = "SELECT Title, Subject, Message FROM EMail WHERE Title='" & strMessageName & "'"
      Set objEMail = Application("Connection").Execute(strSQL)
   Else
      strSQL = "SELECT Title FROM EMail ORDER BY Title"
      Set objEMail = Application("Connection").Execute(strSQL)
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
   <SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
   <!--
   function popitup(url)
   {
   	newwindow=window.open(url,'name','height=370,width=500,scrollbars=yes,top=100,resizable=yes,status=no');
   	if (window.focus) {newwindow.focus()}
   	return false;
   }
   
   // -->
   </SCRIPT> 
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
			<li class="topbar"><a href="setup.asp">Setup</a><font class="separator"> | </font></li> 
         <li class="topbar"><a href="users.asp">Users</a><font class="separator"> | </font></li>
         <li class="topbar">Messages<font class="separator"> | </font></li>
         <li class="topbar"><a href="dbtools.asp">Database Tools</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="index.asp">User Mode</a></li>
      <% If bolShowLogout Then %>
         <font class="separator"> | </font></li>
         <li class="topbar"><a href="login.asp?action=logout">Log Out</a></li>
      <% Else %>
         </li>
      <% End If %>
      </ul>
      </ul>
   </div>
   
<% If InStr(strUserAgent,"MSIE") Then %>
      <hr class="adminbottombarIE"/>
<% Else %>   
      <hr class="adminbottombar"/>
<% End If %>
   <div class="mainarea">
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
      <img class="mainareaimage" src="../themes/<%=Application("Theme")%>/images/admin.gif"/>
   <% Else %>
      <img class="mainareaimage" src="../themes/<%=objNameCheckSet(3)%>/images/admin.gif"/>
   <% End If %>
      <div class="mainareatext">  
   <center>
   <table>
   <form method="POST" action="messages.asp">
      <tr><td>
         Display a message to users of the help desk.  You can choose to have the message visiable by
         users, techs, or both.  Also select the type of message you want to display and where 
         it will be displayed.
      </td></tr>
      <tr><td>
         Enter Your Message:
      </td></tr>
      <tr><td>
         <textarea rows="6" name="Message" style="width: 525px;"><%=strMessage%></textarea>
      </td></tr>
      <tr><td>
         Who will see the message? 
         <select size="1" name="Recipient">
            <option value="Techs" <%=strTechsSelected%>>Techs Only</option>
            <option value="Users" <%=strUsersSelected%>>Users Only</option>
            <option value="Both" <%=strBothSelected%>>Both</option>
         </select>
      </td></tr>
      <tr><td>
         What type of message is this? 
         <select size="1" name="Type">
            <option value="Normal" <%=strNormalSelected%>>Normal</option>
            <option value="Alert" <%=strAlertSelected%>>Alert</option>           
         </select>
      </td></tr>
      <tr><td>
         Where do you want the message displayed? 
         <select size="1" name="Position">
            <option value="Top" <%=strTopSelected%>>Top of Page</option>
            <option value="Bottom" <%=strBottomSelected%>>Bottom of Page</option> 
            <option value="Both" <%=strBothPositionsSelected%>>Both</option> 
         </select>
      </td></tr>
      <tr>
		<td>
			<table width="100%">
				<tr>
					<td>
						Enable Message? 
						<select size="1" name="Enable">
                     <option value="True" <%=strEnabledYes%>>Yes</option>
							<option value="False" <%=strEnabledNo%>>No</option>
						</select>
					</td>
					<td>
						<input type="submit" value="Save" name="cmdsubmit" style="float: right">
					</td>
				</tr>
			</table>
		</td>
	  </tr>
	  
      </form>  
      <tr><td><hr /></td></tr>
      <tr><td>
         Modify one of the default email messages sent from the Help Desk.
      </td></tr>
<%    If strMessageName = "" Then %>      
      <form method="POST" action="messages.asp">
      <tr>
         <td>
            <table width="100%">
               <tr>
                  <td>
                     Choose a message to modify: 
                     <select size="1" name="MessageName">
                        <option></option>
                  <%    Do Until objEMail.EOF %>            
                           <option><%=objEMail(0)%></option>
                  <%       objEMail.MoveNext
                        Loop %>
                     </select>
                  </td>
                  <td>
                     <input type="submit" value="Select" name="cmdSelect" style="float: right">
                  </td>
               </tr>
            </table>
         </td>
		</tr>
      </form>
      <tr><td><hr /></td></tr>
<%    Else %>
      <form method="POST" action="messages.asp">
      <tr><td>
         Title: <%=objEMail(0)%> <input type="hidden" name="title" value="<%=objEMail(0)%>">
      </td></tr>
      <tr><td>
         Subject: <input type="text" name="Subject" value="<%=objEMail(1)%>" size="71">
      </td></tr>
      <tr><td>
         Message: 
      </td></tr>
      <tr><td>
         <textarea rows="12" name="EMailMessage" style="width: 525px;"><%=objEMail(2)%></textarea>
      </td></tr>
      <tr><td align="right">
         <a href="popup.asp?Item=Location" onClick="return popitup('variables.asp')">View Variables</a>
         <input type="submit" value="Cancel" name="cmdEMail">
         <input type="submit" value="Save" name="cmdEMail">
      </td></tr>
      <tr><td><hr /></td></tr> 
      </form>
<%    End If %>
   </table>
   </center>
   
<%    'If a task was performed display the results
      If strMessagetoUser <> "" Then %>
         <font class="information"><%=strMessagetoUser%></font></center>
<%    End If%>
   </div>
   </div>
   </body>
   </html>

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