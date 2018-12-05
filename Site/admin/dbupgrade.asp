<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 12/14/11
'Last Updated 6/16/14

'This page contains some tools that can help minor things to the database.

strVersion = "1.02"

On Error Resume Next

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

   strCMD = Request.Form("Submit")
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
   
   <center>
   <form method="POST" action="dbupgrade.asp">
   <table>
      <tr><td>Use the button below to upgrade the database to the latest version.</td></tr>
      <tr><td align="center"><input type="submit" name ="submit" value="Upgrade Database"></td></tr>
<% Select Case strCMD
      Case "Upgrade Database"
         UpgradeDatabase
      Case "Save"
         strADUserName = Request.Form("ADUserName")
         strADPassword = Request.Form("ADPassword")
         strDomainController = Request.Form("DomainController")
         
         'Update the user account if a password is supplied, otherwise only update the domain controller
         If strADPassword = "" Then
            strSQL = "UPDATE Settings SET DomainController = '" & strDomainController & "' WHERE ID=1"
         ElseIf strDomainController <> "" Then
            strSQL = "UPDATE Settings SET ADUserName = '" & strADUserName & "'," & _
               "ADPassword = '" & strADPassword & "',DomainController = '" & strDomainController & "' " & _
               "WHERE ID=1"
         End If

         Application("Connection").Execute(strSQL)
      
   End Select 
   
   'Get the AD values from the database
   strSQL = "SELECT ADUserName,DomainController FROM Settings WHERE ID=1"
   Set objADSettings = Application("Connection").Execute(strSQL)
   
   If Not objADSettings.EOF Then
      strADUserName = objADSettings(0)
      strDomainController = objADSettings(1)
   End If

   %>
      <tr><td><hr /></td></tr>
      <tr>
         <th>Active Directory Connection</th>
      </tr>
      <tr>
         <td>Domain Controller</td>
      </tr>
      <tr>
         <td width="99%" height="19" valign="bottom">
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
      <tr><td><input type="submit" value="Save" name="submit" style="float: right"></td></tr>
   
   </table>
   </form>
   </center>
   </body>
      
<%End Sub%>

<%
Sub UpgradeDatabase

   'Created by Matthew Hull 12/28/11

   'This script will upgrade the help desk database.

   On Error Resume Next

   'Connect to the database
   Set objCatalog = CreateObject("ADOX.Catalog")
   objCatalog.ActiveConnection = Application("Connection")

   'Check and see if the new tables already exist
   bolMessageTableFound = False
   bolLogTableFound = False
   bolColorsFound = False
   bolTrackingFound = False
   bolEMail = False
   bolFeedbackTableFound = False
   bolTaskListTableFound = False
   bolCounterFound = False
   bolListsTableFound = False
   bolTaskListRoles = False
   bolDocRoles = False
   bolCheckIns = False
   bolSessions = False
   bolSubnets = False
   For Each Table in objCatalog.Tables
      Select Case LCase(Table.Name)
         Case "message"
            bolMessageTableFound = True
         Case "log"
            bolLogTableFound = True
         Case "colors"
            bolColorsFound = True
         Case "tracking"
            bolTrackingFound = True
         Case "email"
            bolEMail = True
         Case "feedback"
            bolFeedbackTableFound = True
         Case "counters"
            bolCounterFound = True
         Case "tasklist"
            bolTaskListTableFound = True
         Case "tasklistroles"
            bolTaskListRoles = True
         Case "docroles"
            bolDocRoles = True
         Case "checkins"
            bolCheckIns = True
         Case "lists"
            bolListsTableFound = True
         Case "sessions"
            bolSessions = True
         Case "subnets"
            bolSubnets = True
      End Select
   Next
   '******************************************************************************************************
   'Remove the colors table
   If bolColorsFound Then
      strSQL = "DROP TABLE Colors;"
      Application("Connection").Execute(strSQL)  
      Response.Write("<tr><td align=""center"">Colors Table: Dropped </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Colors Table: Already Removed </td></tr>")
   End If
   '******************************************************************************************************
   'Create the Message table.  If a record with an ID of 1 doesn't exist then recreate the table.
   If NOT bolMessageTableFound Then
      strSQL = "CREATE TABLE Message(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Message LONGTEXT WITH COMPRESSION,"
      strSQL = strSQL & "Recipient TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Type TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "PositionOnPage TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Enabled BIT);"
      Application("Connection").Execute(strSQL)
      
      strSQL = "INSERT INTO Message(Message,Recipient,Type,PositionOnPage,Enabled)" & vbCRLF
      strSQL = strSQL & "VALUES (' ','Techs','Normal','Top',False)"
      Application("Connection").Execute(strSQL)
      'Response.Write("<tr><td align=""center"">Message Table: Created </td></tr>")
   Else
      strSQL = "SELECT Message FROM Message WHERE ID=1" ' Change ID to 10 to force the table to be recreated.
      Set objMessageTableCheck = Application("Connection").Execute(strSQL)
      
      If objMessageTableCheck.EOF Then
      
         Set objMessageTableCheck = Nothing
      
         strSQL = "DROP TABLE Message;"
         Application("Connection").Execute(strSQL)

         strSQL = "CREATE TABLE Message(" & vbCRLF
         strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
         strSQL = strSQL & "Message LONGTEXT WITH COMPRESSION,"
         strSQL = strSQL & "Recipient TEXT(255) WITH COMPRESSION,"
         strSQL = strSQL & "Type TEXT(255) WITH COMPRESSION,"
         strSQL = strSQL & "PositionOnPage TEXT(255) WITH COMPRESSION,"
         strSQL = strSQL & "Enabled BIT);"
         Application("Connection").Execute(strSQL)
         
         strSQL = "INSERT INTO Message(Message,Recipient,Type,PositionOnPage,Enabled)" & vbCRLF
         strSQL = strSQL & "VALUES (' ','Techs','Normal','Top',False)"
         Application("Connection").Execute(strSQL)
         Response.Write("<tr><td align=""center"">Message Table: Recreated </td></tr>")
      Else
         'Response.Write("<tr><td align=""center"">Message Table: Already Exists </td></tr>")
      End If
   End If
   '******************************************************************************************************
   'Create the Log table
   If NOT bolLogTableFound Then
      strSQL = "CREATE TABLE Log(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Ticket INTEGER,"
      strSQL = strSQL & "Type TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "ChangedBy TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "OldValue TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "NewValue TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "UpdateDate DATETIME,"
      strSQL = strSQL & "UpdateTime DATETIME,"
      strSQL = strSQL & "TaskTime TEXT(255) WITH COMPRESSION);"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Log Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Log Table: Already Exists </td></tr>")
   End If

   '******************************************************************************************************
   'Create the Tracking table
   If Not bolTrackingFound Then
      strSQL = "CREATE TABLE Tracking(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Ticket INTEGER,"
      strSQL = strSQL & "Type TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "TrackedBy TEXT(255) WITH COMPRESSION);"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Tracking Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Tracking Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the EMail table
   If Not bolEMail Then
      strSQL = "CREATE TABLE EMail(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Title TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Subject TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Message LONGTEXT WITH COMPRESSION);"
      Application("Connection").Execute(strSQL) 
      Response.Write("<tr><td align=""center"">EMail Table: Created </td></tr>")
      
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Update Tech'," 
      strSQL = strSQL & "'Help Desk - Ticket ##TICKET# - Update',"
      strSQL = strSQL & "'#CURRENTUSER# sent you an update from the help desk." & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & "Problem: #PROBLEM#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: #NOTES#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" 
      Application("Connection").Execute(strSQL)
      
      strSQL = "INSERT INTO Email (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Update User'," 
      strSQL = strSQL & "'Help Desk - Ticket ##TICKET# - Update',"
      strSQL = strSQL & "'#CURRENTUSER# sent you an update from the help desk." & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: #NOTES#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Problem: #PROBLEM#')"
      Application("Connection").Execute(strSQL)
      
      strSQL = "INSERT INTO Email (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Request for Update'," 
      strSQL = strSQL & "'Help Desk - Ticket ##TICKET# - Update Requested',"
      strSQL = strSQL & "'#CURRENTUSER# is requesting an update on Ticket ##TICKET#.  Use the link below to update the ticket." & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF 
      strSQL = strSQL & "#PROBLEM#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Current Tech Notes: #NOTES#')" & vbCRLF & vbCRLF
      Application("Connection").Execute(strSQL)
      
      strSQL = "INSERT INTO Email (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Ticket Assigned'," 
      strSQL = strSQL & "'Help Desk Assignment - Ticket ##TICKET#',"
      strSQL = strSQL & "'Name: #USER#" & vbCRLF
      strSQL = strSQL & "#PROBLEM#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: #NOTES#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" & vbCRLF & vbCRLF
      Application("Connection").Execute(strSQL)
      
      strSQL = "INSERT INTO Email (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Ticket Reassigned'," 
      strSQL = strSQL & "'Help Desk Ticket ##TICKET# Reassigned',"
      strSQL = strSQL & "'This ticket has been reassigned to #TECH#.  The current status is #STATUS#." & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#PROBLEM#" & vbCRLF
      strSQL = strSQL & "Tech Notes: #NOTES#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" & vbCRLF & vbCRLF
      Application("Connection").Execute(strSQL)
      
      strSQL = "INSERT INTO Email (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Ticket Closed'," 
      strSQL = strSQL & "'Help Desk - Ticket ##TICKET# Has Been Closed',"
      strSQL = strSQL & "'Ticket##TICKET# is complete." & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Below is a description of the problem: " & vbCRLF & """#PROBLEM#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "Technician: #TECH#" & vbCRLF
      strSQL = strSQL & "Tech Notes: ""#NOTES#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "Please do not respond to this email.')" & vbCRLF & vbCRLF
      Application("Connection").Execute(strSQL)   
         
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Requested Update'," 
      strSQL = strSQL & "'Help Desk Requested Update - Ticket ##TICKET#',"
      strSQL = strSQL & "'You are receiving this message because you requested an update." & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & """#PROBLEM#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: ""#NOTES#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" 
      Application("Connection").Execute(strSQL)     

      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Tracking Update'," 
      strSQL = strSQL & "'Help Desk Tracking Update - Ticket ##TICKET#',"
      strSQL = strSQL & "'You are receiving this message because you are tracking ticket ##TICKET#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & """#PROBLEM#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: ""#NOTES#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" 
      Application("Connection").Execute(strSQL)  

      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Send Ticket'," 
      strSQL = strSQL & "'Help Desk - Ticket ##TICKET#',"
      strSQL = strSQL & "'#CURRENTUSER# sent you a ticket from the help desk." & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & "Problem: #PROBLEM#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Status: #STATUS#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: #NOTES#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" 
      Application("Connection").Execute(strSQL)
      
   Else
      'Response.Write("<tr><td align=""center"">EMail Table: Already Exists </td></tr>")
   End If
   
   'Check and see if the New Ticket Admin message is in the EMail table
   strSQL = "SELECT ID FROM EMail WHERE Title = 'New Ticket Admin'"
   Set objEMailCheck = Application("Connection").Execute(strSQL)
   
   'Add the message to the database if it's missing
   If objEMailCheck.EOF Then
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'New Ticket Admin'," 
      strSQL = strSQL & "'Help Desk Ticket ##TICKET#',"
      strSQL = strSQL & "'#USER# has reported a problem." & vbCRLF & vbCRLF
      strSQL = strSQL & "Ticket: #TICKET#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Below is a description of the problem:" & vbCRLF
      strSQL = strSQL & """#PROBLEM#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "This ticket was automatically assigned to #TECH#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" 
      Application("Connection").Execute(strSQL)   
   
      Response.Write("<tr><td align=""center"">EMail Table: Added New Ticket Admin</td></tr>")
   End If
   
   'Check and see if the New Ticket User message is in the EMail table
   strSQL = "SELECT ID FROM EMail WHERE Title = 'New Ticket User'"
   Set objEMailCheck = Application("Connection").Execute(strSQL)
   
   'Add the message to the database if it's missing
   If objEMailCheck.EOF Then
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'New Ticket User'," 
      strSQL = strSQL & "'Help Desk Confirmation - Ticket ##TICKET#',"
      strSQL = strSQL & "'Your help desk request has been processed.  Your ticket number is #TICKET#." & vbCRLF & vbCRLF
      strSQL = strSQL & "Name: #USER#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Below is a description of the problem:" & vbCRLF
      strSQL = strSQL & """#PROBLEM#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "This ticket was automatically assigned to #TECH#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Please do not respond to this email.')" 
      Application("Connection").Execute(strSQL)
   
      Response.Write("<tr><td align=""center"">EMail Table: Added New Ticket User</td></tr>")
   End If
   
   'Check and see if the New Ticket Assigned message is in the EMail table
   strSQL = "SELECT ID FROM EMail WHERE Title = 'New Ticket Assigned'"
   Set objEMailCheck = Application("Connection").Execute(strSQL)
   
   'Add the message to the database if it's missing
   If objEMailCheck.EOF Then
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'New Ticket Assigned'," 
      strSQL = strSQL & "'Help Desk Assignment - ##TICKET#',"
      strSQL = strSQL & "'Name: #USER#" & vbCRLF
      strSQL = strSQL & """#PROBLEM#""" & vbCRLF & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')" 
      Application("Connection").Execute(strSQL)
   
      Response.Write("<tr><td align=""center"">EMail Table: Added New Ticket Assigned</td></tr>")
   End If
   
   'Check and see if the Ticket Closed By User message is in the EMail table
   strSQL = "SELECT ID FROM EMail WHERE Title = 'Ticket Closed By User'"
   Set objEMailCheck = Application("Connection").Execute(strSQL)
   
   'Add the message to the database if it's missing
   If objEMailCheck.EOF Then
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Ticket Closed By User'," 
      strSQL = strSQL & "'Help Desk - Ticket ##TICKET# Has Been Closed',"
      strSQL = strSQL & "'Name: #USER#" & vbCRLF
      strSQL = strSQL & "Problem: #PROBLEM#" & vbCRLF
      strSQL = strSQL & "Email: #USEREMAIL#" & vbCRLF
      strSQL = strSQL & "Location: #LOCATION#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Room: #CUSTOM1#" & vbCRLF
      strSQL = strSQL & "Phone: #CUSTOM2#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Status: #STATUS#" & vbCRLF & vbCRLF
      strSQL = strSQL & "Tech Notes: #NOTES#" & vbCRLF & vbCRLF
      strSQL = strSQL & "#LINK#')"  
      Application("Connection").Execute(strSQL)
   
      Response.Write("<tr><td align=""center"">EMail Table: Added Ticket Closed By User</td></tr>")
   End If   
   
   'Check and see if the Reminder message is in the EMail table
   strSQL = "SELECT ID FROM EMail WHERE Title = 'Reminder'"
   Set objEMailCheck = Application("Connection").Execute(strSQL)
   
   'Add the message to the database if it's missing
   If objEMailCheck.EOF Then
      strSQL = "INSERT INTO EMail (Title,Subject,Message)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'Reminder'," 
      strSQL = strSQL & "'Help Desk Reminder',"
      strSQL = strSQL & "'Hello, " & vbCRLF & vbCRLF
      strSQL = strSQL & "We noticed a new Help Desk ticket was generated for you by someone else.  We wanted to remind "
      strSQL = strSQL & "you that you can enter your own help desk tickets.  Entering the tickets yourself will expedite "  
      strSQL = strSQL & "the process.  Entering a ticket is fast and easy." & vbCRLF & vbCRLF
      strSQL = strSQL & "The address for the system is " & Left(Application("AdminURL"),Len(Application("AdminURL"))-6) & vbCRLF & vbCRLF
      strSQL = strSQL & "Thank you." & vbCRLF & vbCRLF
      strSQL = strSQL & "Please do not respond to this email.')"
      Application("Connection").Execute(strSQL)
   
      Response.Write("<tr><td align=""center"">EMail Table: Added Reminder</td></tr>")
   End If 
   
   
   '******************************************************************************************************
   'Create the Feedback table
   If NOT bolFeedbackTableFound Then
      strSQL = "CREATE TABLE Feedback(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Ticket INTEGER,"
      strSQL = strSQL & "Rating INTEGER,"
      strSQL = strSQL & "Tech TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Location TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Comment LONGTEXT WITH COMPRESSION,"
      strSQL = strSQL & "DateSubmitted DATETIME,"
      strSQL = strSQL & "TimeSubmitted DATETIME);"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Feedback Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Feedback Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the Counter table
   If NOT bolCounterFound Then
      strSQL = "CREATE TABLE Counters(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Feedback INTEGER);"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO Counters (Feedback)" & vbCRLF
      strSQL = strSQL & "VALUES (0)" 
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Counter Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Counter Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the TaskList table
   If NOT bolTaskListTableFound Then
      strSQL = "CREATE TABLE TaskList(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Title TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "List TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Priority TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Tech TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "EnteredBy TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Status TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Rank INTEGER,"
      strSQL = strSQL & "Notes LONGTEXT WITH COMPRESSION,"
      strSQL = strSQL & "DueDate DATETIME,"
      strSQL = strSQL & "DateSubmitted DATETIME,"
      strSQL = strSQL & "TimeSubmitted DATETIME);"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">TaskList Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">TaskList Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the List table
   If NOT bolListsTableFound Then
      strSQL = "CREATE TABLE Lists(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "ListName TEXT(255) WITH COMPRESSION);"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Lists Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Lists Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the TaskListRoles table
   If NOT bolTaskListRoles Then
      strSQL = "CREATE TABLE TaskListRoles(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Role TEXT(255) WITH COMPRESSION);"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO TaskListRoles (Role)" & vbCRLF
      strSQL = strSQL & "VALUES ('User')"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO TaskListRoles (Role)" & vbCRLF
      strSQL = strSQL & "VALUES ('Viewer')"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO TaskListRoles (Role)" & vbCRLF
      strSQL = strSQL & "VALUES ('Deny')"
      Application("Connection").Execute(strSQL)
      
      Response.Write("<tr><td align=""center"">TaskListRoles Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">TaskListRoles Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the DocRoles table
   If NOT bolDocRoles Then
      strSQL = "CREATE TABLE DocRoles(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Role TEXT(255) WITH COMPRESSION);"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO DocRoles (Role)" & vbCRLF
      strSQL = strSQL & "VALUES ('Full')"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO DocRoles (Role)" & vbCRLF
      strSQL = strSQL & "VALUES ('Read Only')"
      Application("Connection").Execute(strSQL)
      strSQL = "INSERT INTO DocRoles (Role)" & vbCRLF
      strSQL = strSQL & "VALUES ('Deny')"
      Application("Connection").Execute(strSQL)
      
      Response.Write("<tr><td align=""center"">DocRoles Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">DocRoles Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the CheckIn table
   If NOT bolCheckIns Then
      strSQL = "CREATE TABLE CheckIns(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Tech TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Location TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "CheckInDate DATETIME,"
      strSQL = strSQL & "CheckInTime DATETIME);"
      Application("Connection").Execute(strSQL)
      
      Response.Write("<tr><td align=""center"">CheckIns Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">CheckIns Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Create the Sessions table
   If NOT bolSessions Then
      strSQL = "CREATE TABLE Sessions(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Username TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "SessionID TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "IPAddress TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "UserAgent TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "LoginDate DATETIME,"
      strSQL = strSQL & "LoginTime DATETIME,"
      strSQL = strSQL & "ExpirationDate DATETIME);"
      Application("Connection").Execute(strSQL)
      
      Response.Write("<tr><td align=""center"">Sessions Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Sessions Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************   
   'Create the Subnets table
   If NOT bolSubnets Then
      strSQL = "CREATE TABLE Subnets(" & vbCRLF
      strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
      strSQL = strSQL & "Subnet TEXT(255) WITH COMPRESSION,"
      strSQL = strSQL & "Location TEXT(255) WITH COMPRESSION);"
      Application("Connection").Execute(strSQL)
      
      Response.Write("<tr><td align=""center"">Subnets Table: Created </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Subnets Table: Already Exists </td></tr>")
   End If
   '******************************************************************************************************     
   'Add Data Viewer as a Role
   strSQL = "SELECT UserLevel FROM UserLevel WHERE UserLevel='Data Viewer'"
   Set objUserLevelCheck = Application("Connection").Execute(strSQL)
   If objUserLevelCheck.EOF Then
      strSQL = "INSERT INTO UserLevel (UserLevel)" & vbCRLF
      strSQL = strSQL & "VALUES ('Data Viewer')"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">UserLevel Table: Added Data Viewer </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">UserLevel Table: Data Viewer Already Exists </td></tr>")
   End If   
   '******************************************************************************************************
   'Check the Main table for required fields
   bolTicketViewedFound = False
   bolTicketTrackedFound = False
   bolUpdateRequestedFound = False
   Set objMainTable = objCatalog.Tables("Main")
   For Each Column in objMainTable.Columns
      Select Case LCase(Column.Name)
         Case "ticketviewed"
            bolTicketViewedFound = True
         Case "tickettracked"
            bolTicketTrackedFound = True
         Case "updaterequested"
            bolUpdateRequestedFound = True
         Case "sourceapi"
            bolSourceAPI = True
         Case "sourceticketnumber"
            bolSourceTicketNumber = True
      End Select
   Next

   'Add the TicketViewed column to the Main table
   If NOT bolTicketViewedFound Then
      strSQL = "ALTER TABLE Main" & vbCRLF
      strSQL = strSQL & "Add TicketViewed BIT"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Main Table: Added TicketViewed </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Main Table: TicketViewed Already Exists </td></tr>")
   End If   

   'Remove the TicketTracked column from the Main table
   If bolTicketTrackedFound Then
      strSQL = "ALTER TABLE Main" & vbCRLF
      strSQL = strSQL & "DROP COLUMN TicketTracked"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Main Table: Removed TicketTracked </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Main Table: TicketTracked Already Removed </td></tr>")
   End If 

   'Remove the UpdateRequested column from the Main table
   If bolUpdateRequestedFound Then
      strSQL = "ALTER TABLE Main" & vbCRLF
      strSQL = strSQL & "DROP COLUMN UpdateRequested"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Main Table: Removed UpdateRequested </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Main Table: UpdateRequested Already Removed </td></tr>")
   End If

   'Add the SourceAPI column to the Main table
   If NOT bolSourceAPI Then
      strSQL = "ALTER TABLE Main" & vbCRLF
      strSQL = strSQL & "Add SourceAPI TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Main Table: Added SourceAPI </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Main Table: SourceAPI Already Exists </td></tr>")
   End If  

   'Add the SourceTicketNumber column to the Main table
   If NOT bolSourceTicketNumber Then
      strSQL = "ALTER TABLE Main" & vbCRLF
      strSQL = strSQL & "Add SourceTicketNumber TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Main Table: Added SourceTicketNumber </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Main Table: SourceTicketNumber Already Exists </td></tr>")
   End If 
   '******************************************************************************************************
   'Check the Settings table for required fields
   bolADUserName = False
   bolADPassword = False
   bolTxtColorFound = False
   bolBGColorFound = False
   bolLnkColorFound = False
   bolInfoColorFound = False
   bolWarningColorFound = False
   bolThemeFound = False
   bolIconLocation = False
   bolUseTaskList = False
   bolUseDocumentation = False
   bolUseStats = False
   bolUseUpload = False
   bolDomainController = False
   bolShowUserStats = False
   bolShowUserButtons = False
   bolVersion = False
   bolSendReminder = False

   Set objMainTable = objCatalog.Tables("Settings")
   For Each Column in objMainTable.Columns
      Select Case LCase(Column.Name)
         Case "adusername"
            bolADUserName = True
         Case "adpassword"
            bolADPassword = True
         Case "txtcolor"
            bolTxtColorFound = True
         Case "bgcolor"
            bolBGColorFound = True
         Case "lnkcolor"
            bolLnkColorFound = True
         Case "infocolor"
            bolInfoColorFound = True
         Case "warningcolor"
            bolWarningColorFound = True
         Case "theme"
            bolThemeFound = True
         Case "iconlocation"
            bolIconLocation = True
         Case "usetasklist"
            bolUseTaskList = True
         Case "usedocumentation"
            bolUseDocumentation = True
         Case "usestats"
            bolUseStats = True
         Case "useupload"
            bolUseUpload = True
         Case "domaincontroller"
            bolDomainController = True
         Case "showuserstats"
            bolShowUserStats = True
         Case "showuserbuttons"
            bolShowUserButtons = True
         Case "version"
            bolVersion = True
         Case "sendreminder"
            bolSendReminder = True
      End Select
   Next    

   'Add the ADUSerName column from the Settings table
   If Not bolADUSerName Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add ADUserName TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If

   'Add the ADPassword column from the Settings table
   If Not bolADPassword Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add ADPassword TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If

   'Remove the TxtColor column from the Settings table
   If bolTxtColorFound Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "DROP COLUMN TxtColor"
      Application("Connection").Execute(strSQL)
   End If

   'Remove the BGColor column from the Settings table
   If bolBGColorFound Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "DROP COLUMN BGColor"
      Application("Connection").Execute(strSQL)
   End If

   'Remove the LnkColor column from the Settings table
   If bolLnkColorFound Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "DROP COLUMN LnkColor"
      Application("Connection").Execute(strSQL)
   End If

   'Remove the InfoColor column from the Settings table
   If bolInfoColorFound Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "DROP COLUMN InfoColor"
      Application("Connection").Execute(strSQL)
   End If

   'Remove the InfoColor column from the Settings table
   If bolWarningColorFound Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "DROP COLUMN WarningColor"
      Application("Connection").Execute(strSQL)
   End If

   If Not bolADUserName Then
      Response.Write("<tr><td align=""center"">Settings Table: Added Active Directory Authentication </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: Active Directory Authentication Already Exists </td></tr>")
   End If

   If bolTxtColorFound Or bolBGColorFound Or bolLnkColorFound Or bolInfoColorFound Or bolWarningColorFound Then
      Response.Write("<tr><td align=""center"">Settings Table: Removed Colors </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: Colors Already Removed </td></tr>")
   End If

   If NOT bolThemeFound Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add Theme TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings" & vbCRLF
      strSQL = strSQL & "SET Theme='Dark Blue'" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Settings Table: Added Theme </td></tr>")
   Else
      strSQL = "UPDATE Settings" & vbCRLF
      strSQL = strSQL & "SET Theme='Dark Blue'" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      'Application("Connection").Execute(strSQL)
      'Response.Write("<tr><td align=""center"">Settings Table: Theme Already Exists </td></tr>")
   End If

   If NOT bolIconLocation Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add IconLocation TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings" & vbCRLF
      strSQL = strSQL & "SET IconLocation='http://help.wswheboces.org/icons'" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Settings Table: Added IconLocation </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: IconLocation Already Exists </td></tr>")
   End If

   If NOT bolUseTaskList Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add UseTaskList BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET UseTaskList = True" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added UseTaskList </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: UseTaskList Already Exists </td></tr>")
   End If

   If NOT bolUseDocumentation Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add UseDocumentation BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET UseDocumentation = True" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added UseDocumentation </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: UseDocumentation Already Exists </td></tr>")
   End If

   If NOT bolUseStats Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add UseStats BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET UseStats = True" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added UseStats </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: UseStats Already Exists </td></tr>")
   End If

   If NOT bolUseUpload Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add UseUpload BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET UseUpload = True" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added UseUpload </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: UseUpload Already Exists </td></tr>")
   End If
   
   If NOT bolDomainController Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add DomainController TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)  
      Response.Write("<tr><td align=""center"">Settings Table: Added DomainController </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: DomainController Already Exists </td></tr>")
   End If
   
   If NOT bolShowUserStats Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add ShowUserStats BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET ShowUserStats = True" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added ShowUserStats </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: ShowUserStats Already Exists </td></tr>")
   End If
   
   If NOT bolShowUserButtons Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add ShowUserButtons BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET ShowUserButtons = True" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added ShowUserButtons </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: ShowUserButtons Already Exists </td></tr>")
   End If
   
   If NOT bolVersion Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add Version TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Settings Table: Added Version </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: Version Already Exists </td></tr>")
   End If
   
   If NOT bolSendReminder Then
      strSQL = "ALTER TABLE Settings" & vbCRLF
      strSQL = strSQL & "Add SendReminder BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Settings SET SendReminder = False" & vbCRLF
      strSQL = strSQL & "WHERE ID=1"
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Settings Table: Added SendReminder </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Settings Table: SendReminder Already Exists </td></tr>")
   End If 
   
   '******************************************************************************************************
   'Check the Location table for required fields
   bolTechLocation = False
   Set objLocationTable = objCatalog.Tables("Location")
   For Each Column in objLocationTable.Columns
      Select Case Column
         Case "Tech"
            bolTechLocation = True
      End Select
   Next
   
   If NOT bolTechLocation Then
      strSQL = "ALTER TABLE Location" & vbCRLF
      strSQL = strSQL & "ADD Tech TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Location Table: Added Tech </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Location Table: Tech Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Check the Tech table for required fields
   bolTechThemeFound = False
   bolTechMobileVersion = False
   bolTechTaskListRole = False
   bolTechDocRole = False
   Set objTechTable = objCatalog.Tables("Tech")
   For Each Column in objTechTable.Columns
      Select Case Column
         Case "Theme"
            bolTechThemeFound = True
         Case "MobileVersion"
            bolTechMobileVersion = True
         Case "TaskListRole"
            bolTechTaskListRole = True
         Case "DocumentationRole"
            bolTechDocRole = True
      End Select
   Next

   If NOT bolTechThemeFound Then
      strSQL = "ALTER TABLE Tech" & vbCRLF
      strSQL = strSQL & "ADD Theme TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Tech Table: Added Theme </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Tech Table: Theme Already Exists </td></tr>")
   End If

   If NOT bolTechMobileVersion Then
      strSQL = "ALTER TABLE Tech" & vbCRLF
      strSQL = strSQL & "ADD MobileVersion BIT"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Tech SET MobileVersion = True" & vbCRLF
      Application("Connection").Execute(strSQL)   
      Response.Write("<tr><td align=""center"">Tech Table: Mobile Version Added </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Tech Table: Mobile Version Already Exists </td></tr>")
   End If

   If NOT bolTechTaskListRole Then
      strSQL = "ALTER TABLE Tech" & vbCRLF
      strSQL = strSQL & "ADD TaskListRole TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Tech SET TaskListRole = 'User'"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Tech Table: Added TaskListRole </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Tech Table: TaskListRole Already Exists </td></tr>")
   End If

   If NOT bolTechDocRole Then
      strSQL = "ALTER TABLE Tech" & vbCRLF
      strSQL = strSQL & "ADD DocumentationRole TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
      strSQL = "UPDATE Tech SET DocumentationRole = 'Full'"
      Application("Connection").Execute(strSQL)
      Response.Write("<tr><td align=""center"">Tech Table: Added DocumentationRole </td></tr>")
   Else
      'Response.Write("<tr><td align=""center"">Tech Table: DocumentationRole Already Exists </td></tr>")
   End If
   '******************************************************************************************************
   'Mark Completed tickets as viewed
   strSQL = "UPDATE Main SET TicketViewed = True" & vbCRLF
   strSQL = strSQL & "WHERE Status=""Complete"""
   Application("Connection").Execute(strSQL)
   'Response.Write("<tr><td align=""center"">Completed Tickets Marked as Viewed </td></tr>")

   'Disable tracking on closed tickets 
   strSQL = "DELETE Tracking.*" & vbCRLF
   strSQL = strSQL & "FROM Main INNER JOIN Tracking ON Main.ID = Tracking.Ticket" & vbCRLF
   strSQL = strSQL & "WHERE (Main.Status='Complete')"
   Application("Connection").Execute(strSQL)
   'Response.Write("<tr><td align=""center"">Fixed Tracking Bug </td></tr>")
   
   strSQL = "UPDATE Settings" & vbCRLF
   strSQL = strSQL & "SET Version='" & strVersion & "'" & vbCRLF
   strSQL = strSQL & "WHERE ID=1"
   Application("Connection").Execute(strSQL)
   Response.Write("<tr><td align=""center"">Database Upgraded to Version " & strVersion & " </td></tr>")
   
End Sub
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

   Dim objAdminAccount

   strSQL = "SELECT Username FROM Tech WHERE UserLevel='Administrator' And Active"
   Set objAdminAccount = Application("Connection").Execute(strSQL)
   
   'Grab any admin account and perform the upgrade as them
   If objAdminAccount.EOF Then
      GetUser = ""
   Else
      GetUser = objAdminAccount(0)
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