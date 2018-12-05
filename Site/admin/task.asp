<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 2/13/12
'Last Updated 6/16/14

'This is the task list page.

Option Explicit

On Error Resume Next

Dim objNetwork, strUserAgent, strSQL, objNameCheckSet, strRole, intID, intInputSize, objTask, objProrities
Dim objLists, objTechs, strSubmit, strTitle, strList, strPriority, strTech, datDueDate, strNotes
Dim strMessage, strMessageType, bolShowLogout, strUser

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
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole" & vbCRLF
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
      If objNameCheckSet(5) <> "Deny" Then
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
      intInputSize = 110
   End If

   'Get the Task ID from the URL
   intID = Request.QueryString("ID")
   If intID = "" Then
      intID = Request.Form("ID")
   End If
   
   'If the ID provided isn't a number then send them back to the tasklist page
   If Not IsNumeric(intID) or intID = "" Then
      Response.Redirect("tasklist.asp")
   End If

   strSubmit = Request.Form("cmdSubmit")
   Select Case strSubmit
   
      Case "Save"
         
         'Get information from the form
         strTitle = Replace(Request.Form("Title"),"'","''")
         strList = Replace(Request.Form("List"),"'","''")
         strPriority = Replace(Request.Form("Priority"),"'","''")
         strTech = Replace(Request.Form("Tech"),"'","''")
         strNotes = Replace(Request.Form("Notes"),"'","''")

         If Request.Form("DueDate") <> "" Then
            Err.Clear
            datDueDate = "#" & CDate(Request.Form("DueDate")) & "#"
            If Err Then
               strMessage = "Invalid Date Format"
               strMessageType = "missing"
               Err.Clear
               datDueDate = "Null"
            End If
         Else
            datDueDate = "Null"
         End If
         
         strSQL = "UPDATE TaskList SET "
         strSQL = strSQL & "Title='" & strTitle & "',"
         strSQL = strSQL & "List='" & strList & "',"
         strSQL = strSQL & "Priority='" & strPriority & "',"
         strSQL = strSQL & "Tech='" & strTech & "',"
         strSQL = strSQL & "DueDate=" & datDueDate & ","
         strSQL = strSQL & "Notes='" & strNotes & "' "
         strSQL = strSQL & "WHERE ID=" & intID

         Application("Connection").Execute(strSQL)
         
         strMessage = "Saved"
         strMessageType = "information"
   
      Case "Close"
         Response.Redirect("tasklist.asp")
   End Select
   
   'Get the details about the task from the database
   strSQL = "SELECT Title,List,Priority,Tech,EnteredBy,Status,Rank,Notes,DueDate,DateSubmitted,TimeSubmitted" & vbCRLF
   strSQL = strSQL & "FROM TaskList" & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intID
   Set objTask = Application("Connection").Execute(strSQL)
   
   'If an invalid number is given then send them back to the task list
   If objTask.EOF Then
      Response.Redirect("tasklist.asp")
   End If
   
   'If the notes are a null value set them to an empty string
   If IsNull(objTask(7)) Then
      strNotes = ""
   Else
      strNotes = objTask(7)
   End If
   
   'List of Priorities
   Set objProrities = CreateObject("ADODB.RecordSet")
   With objProrities
      Set .ActiveConnection = Nothing
      .CursorLocation = 3
      .LockType = 4
      With .Fields
         .Append "Priority", 8, 30
      End With
      .Open
   End With
   objProrities.AddNew "Priority","High"
   objProrities.AddNew "Priority","Normal"
   objProrities.AddNew "Priority","Low"
   objProrities.Update
   objProrities.MoveFirst
   
   'Get the list of lists :)
   strSQL = "SELECT ListName FROM Lists"
   Set objLists = Application("Connection").Execute(strSQL)
   
   'Get the list of techs
   strSQL = "SELECT Tech,UserName FROM Tech"
   Set objTechs = Application("Connection").Execute(strSQL)
   
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
			<li class="topbar"><a href="stats.asp">Stats</a><font class="separator"> | </font></li> 
      <% If Application("UseDocs") Then %>
         <li class="topbar"><a href="docs.asp">Docs</a><font class="separator"> | </font></li>
      <% End If%>
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
   <form method="POST" action="task.asp">
      <input type="hidden" name="ID" value="<%=intID%>" />
      <tr>
         <th colspan="4">Task #<%=intID%></th>
      </tr>
      <tr>
         <td class="ShowBorders">
            Title:&nbsp;
         </td>
         <td class="ShowBorders" colspan="3">
            <input type="textbox" name="Title" value="<%=objTask(0)%>" size="<%=intInputSize%>" />
         </td>
      </tr>
      
      <tr>
         <td class="ShowBorders" width="50px">
            List:&nbsp;
         </td>
         <td class="ShowBorders" width="325px">
            <select name="List">
               <option></option>
            <% If NOT objLists.EOF Then
                  Do Until objLists.EOF
                     If LCase(objLists(0)) = LCase(objTask(1)) Then %>   
                        <option selected="selected"><%=objLists(0)%></option>
                  <% Else %>
                        <option><%=objLists(0)%></option>
                  <% End If %>
                  <% objLists.MoveNext 
                  Loop 
               End If%>
            </select>
         </td>
         <td class="ShowBorders" width="90px">
            Priority:&nbsp;
         </td>
         <td class="ShowBorders" width="285px">
            <select name="Priority">
            <% Do Until objProrities.EOF 
                  If objProrities(0) = objTask(2) Then %>   
                     <option selected="selected"><%=objProrities(0)%></option>
               <% Else %>
                     <option><%=objProrities(0)%></option>
               <% End If
                  objProrities.MoveNext
               Loop %>
            </select>
         </td>
      </tr>
      
      <tr>
         <td class="ShowBorders">
            Tech:&nbsp;
         </td>
         <td class="ShowBorders">
            <select name="Tech">
               <option></option>
            <% Do Until objTechs.EOF 
                  If LCase(objTask(3)) = LCase(objTechs(1)) Then %>
                     <option value="<%=objTechs(1)%>" selected="selected"><%=objTechs(0)%></option>
               <% Else %>
                     <option value="<%=objTechs(1)%>"><%=objTechs(0)%></option>
               <% End If 
                  objTechs.MoveNext
               Loop %>
            </select>
         </td>
         <td class="ShowBorders">
            Due Date:&nbsp;
         </td>
         <td class="ShowBorders">
            <input type="textbox" name="DueDate" value="<%=objTask(8)%>" />
         </td>
      </tr>
         
      <tr>
         <td class="ShowBorders" colspan="4">
            Notes:
         </td>
      </tr>
      <tr>
         <td class="ShowBorders" colspan="4">
            <textarea rows="5" style="width: 100%;" name="Notes"><%=strNotes%></textarea>
         </td>
      </tr>
      <tr>
         <td class="ShowBorders" colspan="4" align="right">
         <% If strMessage <> "" Then %>
               <font class="<%=strMessageType%>"><%=strMessage%></font>
         <% End If %>
            <input type="submit" name="cmdSubmit" value="Close" />
         <% If objNameCheckSet(5) = "User" Then %>   
               <input type="submit" name="cmdSubmit" value="Save" />
         <% End If %>
         </td>
      </tr>
   </form>
   </table>
   <center>

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