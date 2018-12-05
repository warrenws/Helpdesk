<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 12/14/11
'Last Updated 6/16/14

'This page contains some tools that can help minor things to the database.

'If the database and the website are not the same version then let them know
If Application("VersionError") Then
   VersionProblem
End If

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
strSQL = "Select Username, UserLevel, Active, Theme" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"

Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)

'See if the user has the rights to visit this page
If objNameCheckSet(1) = "Administrator" And objNameCheckSet(2) Then

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
%>

<%Sub AccessGranted 

   intID = Request.Form("Ticket")
   strCMD = Request.Form("Submit")
   
   'Remove all old sessions
   strSQL = "DELETE FROM Sessions WHERE Date() > ExpirationDate"
   Application("Connection").Execute(strSQL)
   
   'See if there is any sort string in the URL for the active sessions list
   Select Case LCase(Request.QueryString("Sort"))
      Case "date"
         strSQLSort = "ORDER BY LoginDate DESC, LoginTime DESC"
      Case "name"
         strSQLSort = "ORDER BY Username"
      Case "device"
         strSQLSort = "ORDER BY UserAgent"
      Case Else
         strSQLSort = "ORDER BY LoginDate DESC, LoginTime DESC"
   End Select
   
   'Get the list of active sessions
   strSQL = "SELECT ID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate,UserName "
   strSQL = strSQL & "FROM Sessions " & strSQLSort
   Set objSessions = Application("Connection").Execute(strSQL)
   
   'Get the selected sessions and split them into an array
   strSelectedSessions = Request.Form("chkSession")
   objSelectedSessions = Split(strSelectedSessions,",")
   
   Select Case strCMD
      Case "Mark Viewed"
         If IsNumeric(intID) Then
            strViewedMessage = SetViewed
         Else
            strViewedMessage = "Invalid Ticket Number"
         End If
      
      Case "Fix Username"
         If IsNumeric(intID) Then
            strUsernameFixedMessage = FixUserName
         Else
            strUsernameFixedMessage = "Invalid Ticket Number"
         End If
      
      Case "Fix Completed Tickets"
         strAllViewedMessage = MarkCompletedAsViewed
      
      Case "Change Usernames"
         strChangeUsernameMessage = ChangeUsername
      
      Case "Fix Tech Names"
         strFixTechNameMessage = FixTechNames
         
      Case "Delete Selected"
         For Each strSession in objSelectedSessions
            strSQL = "DELETE FROM Sessions" & vbCRLF
            strSQL = strSQL & "Where ID = " & Trim(strSession)
            Application("Connection").Execute(strSQL)
         Next   
         GetUser
         
      Case "Delete All"
         strSQL = "DELETE FROM Sessions" & vbCRLF
         Application("Connection").Execute(strSQL) 
         GetUser
         
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
         <li class="topbar"><a href="messages.asp">Messages</a><font class="separator"> | </font></li>
         <li class="topbar">Database Tools<font class="separator"> | </font></li>
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
      <tr><td>
      This page contains a few tools you can use to make changes to data in the database.
      </td></tr>
      <tr><td><hr /></td></tr>
      <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td colspan="4">Mark a ticket as viewed by the tech.  
            This will also log an event in the event log.</td></tr>
            <tr>
               <td>Enter the ticket number.
               <input type="text" name="Ticket" size="10">
               <input type="submit" name ="submit" value="Mark Viewed"> 
         <% If strViewedMessage = "Ticket viewed" Then %>
               <font class="information"><%=strViewedMessage%></font></td>
         <% Else %>
               <font class="missing"><%=strViewedMessage%></font></td>
         <% End If%>
            </tr>
            </form>
         </table>
      </td></tr>
   <!--   
      <tr><td><hr /></td></tr>
       <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td colspan="4">Mark completed tickets as viewed
               <input type="submit" name ="submit" value="Fix Completed Tickets"> 
         <% If strAllViewedMessage = "Complete" Then %>
               <font class="information"><%=strAllViewedMessage%></font></td>
         <% End If%>
            </tr>
            </form>
         </table>
      </td></tr>    
   -->
      <tr><td><hr /></td></tr>
      <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td colspan="4">Sometimes when a ticket is entered manually the wrong username is used.
            This tool will let you fix the ticket.  This will correct the Username, Display Name and Email address
            fields in the database.</td></tr>
            <tr>
               <td>
                  Enter the ticket number. <input type="text" name="Ticket" size="10">
                  <% If strUsernameFixedMessage <> "Updated" Then %>
                     <font class="missing"><%=strUsernameFixedMessage%></font></td>
                  <% End If %>
               </td>
            </tr>
            <tr>
               <td>
                  Enter the correct username. <input type="text" name="Username" size="10">
                  <input type="submit" name ="submit" value="Fix Username">
            <% If strUsernameFixedMessage = "Updated" Then %>
                  <font class="information"><%=strUsernameFixedMessage%></font></td>
            <% End If %>
               </td>
            </tr>
            </form>
         </table>
      </td></tr>
      <tr><td><hr /></td></tr>
      <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td colspan="4">If you want to change all the tickets from one person to another use this tool.</td></tr>
            <tr>
               <td>
                  Enter the old username <input type="text" name="OldUsername" size="10">
                  <% If strChangeUsernameMessage <> "Updated" Then %>
                     <font class="missing"><%=strChangeUsernameMessage%></font></td>
                  <% End If %>
               </td>
            </tr>
            <tr>
               <td>
                  Enter the new username. <input type="text" name="NewUsername" size="10">
                  <input type="submit" name ="submit" value="Change Usernames">
            <% If strChangeUsernameMessage = "Updated" Then %>
                  <font class="information"><%=strChangeUsernameMessage%></font></td>
            <% End If %>
               </td>
            </tr>
            </form>
         </table>
      </td></tr>
      <tr><td><hr /></td></tr>      
      <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td>If you are using Active Directory to look up user's information then your tech's name's
            need to be in the format Firstname LastName.  This button will fix your techs and update tickets.
            (Note this button is disabled if your not using Active Directory lookups.)
      <% If Application("UseAD") Then %>      
            <input type="submit" name ="submit" value="Fix Tech Names">
      <% Else %>
            <input type="submit" name ="submit" disabled="disabled" value="Fix Tech Names">
      <%End If %>
      <% If strFixTechNameMessage <> "" Then %>
            <font class="information"><%=strFixTechNameMessage%></font></td>
      <% End If %>
            </td></tr>
            <tr><td><hr /></td></tr>
      <% If Not objSessions.EOF Then %>
            <tr><td align="center"><a name="sort">Active Sessions</a></td></tr>
            <tr><td align="center">
               <table>
                  <tr>
                     <th>&nbsp;</th>
                     <th>&nbsp;<a href="dbtools.asp?sort=date#sort">Date</a>&nbsp;</th>
                     <th>&nbsp;Time&nbsp;</th>
                     <th>&nbsp;<a href="dbtools.asp?sort=name#sort">Name</a>&nbsp;</th>
                     <th>&nbsp;<a href="dbtools.asp?sort=device#sort">Device Type</a>&nbsp;</th>
                  </tr>
            <% Do Until objSessions.EOF 'ID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate,UserName
            %>
                  <tr>
                     <td class="showborders" align="center" width="1%">
                        <input type="checkbox" name="chkSession" value="<%=objSessions(0)%>">
                     </td>
                     <td class="showborders">&nbsp;<%=Left(objSessions(3), Len(objSessions(3)) - 5)%>&nbsp;</td>
                     <td class="showborders">&nbsp;<%=Left(objSessions(4), Len(objSessions(4)) - 6) & Right(objSessions(4),2)%>&nbsp;</td>
                     <td class="showborders">&nbsp;<%=GetFirstandLastName(objSessions(6))%>&nbsp;</td>
                     <td class="showborders">&nbsp;<%=GetDeviceType(objSessions(2))%>&nbsp;</td>
                  </tr>
            <%    objSessions.MoveNext
               Loop %>
                  <tr><td colspan="5">
                     <input type="submit" value="Delete Selected" name="submit" style="float: right">
                     <input type="submit" value="Delete All" name="submit" style="float: right">
                  </td></tr>
               </table>
            </td></tr>
            <tr><td><hr /></td></tr>
      <% End If %>

            </form>
         </table>
      </td></tr>
   <!-- 
      <tr><td><hr /></td></tr>
      <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td>Use this button to show all tickets with an improper DisplayName.
            <input type="submit" name ="submit" value="Show Bad Display Names">
            </td></tr>
            </form>
         </table>
      </td></tr>
      <tr><td><hr /></td></tr>
      <tr><td>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td>Use this button to fix all the display names in the database.
            <input type="submit" name ="submit" value="Fix All Display Names">
            </td></tr>
            </form>
         </table>
      </td></tr>
      <tr><td><hr /></td></tr>
<%    If strCMD = "Fix All Display Names" Then %>
         <table>
            <form method="POST" action="dbtools.asp">
            <tr><td>Are you sure you want to fix all display name?  Type in YES:
            <input type="text" name="Yes">
            <input type="submit" name ="submit" value="Confirm">
            
            </td></tr>
            </form>
         </table>
<%    End If%>

<%    If strCMD = "Confirm" And Request.Form("Yes") = "YES" Then
         FixAllNames
      End If%>
      
<%    If strCMD = "Show Bad Display Names" Then
         ShowBadDisplayNames
      End If%>
   -->
   </table>
   </center>
   </div>
   </div>
   </body>
      
<%End Sub%>


<%Function SetViewed
   'Created by Matthew Hull on 11/9/11
   
   intID = Request.Form("Ticket")
   
   strSQL = "SELECT Tech,TicketViewed" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF 
   strSQL = strSQL & "WHERE ID=" & intID
   
   Set objTech = Application("Connection").Execute(strSQL)
   
   If objTech.EOF Then
      strMessage = "Invalid Ticket Number"
   Else
      If Not objTech(1) Then
         If objTech(0) <> "" Then
            strSQL = "UPDATE Main SET TicketViewed = True" & vbCRLF
            strSQL = strSQL & "WHERE ID=" & intID

            Application("Connection").Execute(strSQL)
            
            strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
            strSQL = strSQL & "VALUES (" & intID & ",'Ticket Viewed','" & objTech(0) & "','" & Date() & "','" & Time() & "');"
            Application("Connection").Execute(strSQL)
            strMessage = "Ticket viewed"
         Else
            strMessage = "No tech assigned"
         End If
      Else
         strMessage = "Already viewed"
      End If
   End If
   
   SetViewed = strMessage

End Function%>

<%Function FixUserName
   'Created by Matthew Hull on 12/14/11
   
   intID = Request.Form("Ticket")
   strDisplayName = Replace(GetFirstandLastName(Request("UserName")),"'","''")
   strUserName = Replace(Request("UserName"),"'","''")

   If strUserName <> strDisplayName Then
      strSQL = "UPDATE Main" & vbCRLF
      strSQL = strSQL & "SET Name='" & strUserName & "',DisplayName='" & strDisplayName & "',EMail='" & strUserName & Application("EMailSuffix") & "'" & vbCRLF
      strSQL = strSQL & "WHERE ID=" & intID

      Application("Connection").Execute(strSQL)
      strMessage = "Updated"
   Else
      strMessage = "Bad username"
   End If
   
   FixUserName = strMessage

End Function%>

<%Function ChangeUsername
   strOldUserName = Replace(Request("OldUserName"),"'","''")
   strNewUserName = Replace(Request("NewUserName"),"'","''")
   strOldDisplayName = Replace(GetFirstandLastName(Request("OldUsername")),"'","''")
   strNewDisplayName = Replace(GetFirstandLastName(Request("NewUsername")),"'","''")
   
   strSQL = "UPDATE Main" & vbCRLF
   strSQL = strSQL & "SET Name='" & strNewUserName & "',DisplayName='" & strNewDisplayName & "',EMail='" & strNewUserName & Application("EMailSuffix") & "'" & vbCRLF
   strSQL = strSQL & "WHERE Name='" & strOldUserName & "'"
   Application("Connection").Execute(strSQL)
   
   ChangeUsername = "Updated"
End Function %>

<%Function MarkCompletedAsViewed
   'Created by Matthew Hull on 11/9/11

   strSQL = "UPDATE Main SET TicketViewed = True" & vbCRLF
   strSQL = strSQL & "WHERE Status=""Complete"""

   Application("Connection").Execute(strSQL)

   MarkCompletedAsViewed = "Complete"
   
End Function%>

<%Function FixTechNames

   strSQL = "SELECT ID,UserName,Tech FROM Tech"
   Set objTechList = Application("Connection").Execute(strSQL)

   If Not objTechList.EOF Then
      Do Until objTechList.EOF
      
         If objTechList(1) <> "Tech Services" Or objTechList(1) <> "Heat Help Desk" Then
      
            intID = objTechList(0)
            strTechUserName = objTechList(1)
            strOldName = objTechList(2)
            strNewName = GetFirstandLastName(strTechUserName)

            strSQL = "UPDATE Tech" & vbCRLF
            strSQL = strSQL & "SET Tech='" & strNewName & "'" & vbCRLF
            strSQL = strSQL & "Where ID=" & intID
            Application("Connection").Execute(strSQL)
            
            strSQL = "UPDATE Main" & vbCRLF
            strSQL = strSQL & "SET Tech='" & strNewName & "'" & vbCRLF
            strSQL = strSQL & "WHERE Tech='" & strOldName & "'"
            Application("Connection").Execute(strSQL)
            
            strSQL = "UPDATE Log" & vbCRLF
            strSQL = strSQL & "SET OldValue='" & strNewName & "'" & vbCRLF
            strSQL = strSQL & "WHERE OldValue='" & strOldName & "'"
            Application("Connection").Execute(strSQL)
            
            strSQL = "UPDATE Log" & vbCRLF
            strSQL = strSQL & "SET NewValue='" & strNewName & "'" & vbCRLF
            strSQL = strSQL & "WHERE NewValue='" & strOldName & "'"
            Application("Connection").Execute(strSQL)
         
         End If
      
         objTechList.MoveNext
      Loop
   End If
   
   FixTechNames = "Complete"
End Function%>

<%Function FixAllNames
   strSQL = "Select ID,Email,DisplayName" & vbCRLF
   strSQL = strSQL & "From Main" & vbCRLF
   strSQL = strSQL & "ORDER BY ID"

   Set objRecordSet = Application("Connection").Execute(strSQL)

   Set objRegExp = New RegExp
   objRegExp.Pattern = "'"
   objRegExp.Global = True %>
   
   <table border="1">
	<tr>
		<th>&nbsp;&nbsp;&nbsp;Ticket&nbsp;&nbsp;&nbsp;</th>
		<th>&nbsp;&nbsp;&nbsp;User Name&nbsp;&nbsp;&nbsp;</th>
		<th>&nbsp;&nbsp;&nbsp;Old Name&nbsp;&nbsp;&nbsp;</th>
		<th>&nbsp;&nbsp;&nbsp;New Name&nbsp;&nbsp;&nbsp;</th>
	</tr>
   
<% Do  Until objRecordSet.EOF
      intID = objRecordSet(0)
      strEMail = objRecordSet(1)
      strUserName = objRegExp.Replace(Left(strEmail,InStr(strEMail,"@")-1),"''")
      strDisplayName = objRegExp.Replace(GetFirstandLastName(strUserName),"''")
      strDisplayName = GetFirstandLastName(strUserName)
      strOldDisplayName = objRecordSet(2)
   %>
      <tr>
         <td class="showborders"><%=intID%></td>
         <td class="showborders"><%=strUserName%></td>
         <td class="showborders"><%=strOldDisplayName%></td>
         <td class="showborders"><%=strDisplayName%></td>
      </tr>
   <%
      strSQL = "UPDATE Main" & vbCRLF & "SET DisplayName='" & Replace(strDisplayName,"'","''") & "'" & vbCRLF & _
      "WHERE ID=" & objRecordSet(0)

      Application("Connection").Execute(strSQL)

      objRecordSet.MoveNext

   Loop
End Function%>

<%Function ShowBadDisplayNames
   strSQL = "Select ID,Name,DisplayName" & vbCRLF
   strSQL = strSQL & "From Main" & vbCRLF
   strSQL = strSQL & "ORDER BY ID"

   Set objRecordSet = Application("Connection").Execute(strSQL)
 %>
   <table border="1">
	<tr>
		<th>&nbsp;&nbsp;&nbsp;Ticket&nbsp;&nbsp;&nbsp;</th>
		<th>&nbsp;&nbsp;&nbsp;User Name&nbsp;&nbsp;&nbsp;</th>
		<th>&nbsp;&nbsp;&nbsp;Display Name&nbsp;&nbsp;&nbsp;</th>
	</tr>
   
<% Do  Until objRecordSet.EOF
      If objRecordSet(1) = objRecordSet(2) Then
   %>
         <tr>
            <td class="showborders"><%=objRecordSet(0)%></td>
            <td class="showborders"><%=objRecordSet(1)%></td>
            <td class="showborders"><%=objRecordSet(2)%></td>
         </tr>
   <% End If
      objRecordSet.MoveNext
   Loop
End Function%>

<%
Function GetFirstandLastName(strUserName)

   On Error Resume Next

   Dim objConnection, objCommand, objRootDSE, objRecordSet,strDNSDomain

   If Application("UseAD") Then
      'Create objects required to connect to AD
      Set objConnection = CreateObject("ADODB.Connection")
      Set objCommand = CreateObject("ADODB.Command")
      Set objRootDSE = GetObject("LDAP://" & Application("Domain") & "/rootDSE")

      'Create a connection to AD
      objConnection.Provider = "ADSDSOObject"

      objConnection.Open "Active Directory Provider", Application("ADUsername"), Application("ADPassword")
      objCommand.ActiveConnection = objConnection
      strDNSDomain = objRootDSE.Get("DefaultNamingContext")
      objCommand.CommandText = "<LDAP://" & Application("DomainController") & "/" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user)(samaccountname=" & strUserName & ")); GivenName,SN,name ;subtree"

      'Initiate the LDAP query and return results to a RecordSet object.
      Set objRecordSet = objCommand.Execute

      If NOT objRecordSet.EOF Then
         If objRecordSet(0) = "" Then
            GetFirstandLastName = strUserName
         Else
            GetFirstandLastName = objRecordSet(0) & " " & objRecordSet(1)
         End If
      Else
         GetFirstandLastName= strUserName
      End If

   Else
      GetFirstandLastName= strUserName
   End If

End Function
%>

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