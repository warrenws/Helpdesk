<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 2/13/12
'Last Updated 6/16/14

'This is the task list page.

Option Explicit

On Error Resume Next

Dim objNetwork, strUserAgent, strSQL, objNameCheckSet, strRole, strCMD, strNewTask
Dim objTaskList, strSelectedTasks, objSelectedTasks, strTask, intInputSize, objLists
Dim strList, strTitle, intColumns, strNewList, strAddTask, strAddList, strFixedList
Dim bolShowLogout, strUser

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

   If InStr(strUserAgent,"Android") Or InStr(strUserAgent,"Silk") Then
      intInputSize = 65
   Else
      intInputSize = 75
   End If

   strCMD = Request.Form("cmdSubmit")
   strNewTask = Replace(Request.Form("NewTask"),"'","''")
   strNewList = Replace(Request.Form("NewList"),"'","''")
   strSelectedTasks = Request.Form("chkComplete")
   strList = Request.Form("List")
   
   objSelectedTasks = Split(strSelectedTasks,",")
   
   Select Case strCMD
      Case "Mark Complete"
         For Each strTask in objSelectedTasks
            strSQL = "UPDATE TaskList" & vbCRLF
            strSQL = strSQL & "SET Status='Complete'" & vbCRLF
            strSQL = strSQL & "Where ID = " & Trim(strTask)
            Application("Connection").Execute(strSQL)
         Next
         
      Case "Delete"
         For Each strTask in objSelectedTasks
            strSQL = "DELETE FROM TaskList" & vbCRLF
            strSQL = strSQL & "Where ID = " & Trim(strTask)
            Application("Connection").Execute(strSQL)
            strList = "Complete"
         Next
         
      Case "Not Complete"
         For Each strTask in objSelectedTasks
            strSQL = "UPDATE TaskList" & vbCRLF
            strSQL = strSQL & "SET Status='Open'" & vbCRLF
            strSQL = strSQL & "Where ID = " & Trim(strTask)
            Application("Connection").Execute(strSQL)
            strList = "Complete"
         Next
         
      Case "Delete List"
         strSQL = "DELETE FROM Lists" & vbCRLF
         strSQL = strSQL & "WHERE ListName = '" & Replace(strList,"'","''") & "'"
         Application("Connection").Execute(strSQL)
         strList = "Open"
   End Select
   
   If strCMD = "Add Task" And strNewTask <> "" Then
      If strList = "Complete" Or strList = "Open" Or strList = "Unassigned" Then
         strFixedList = ""
      Else
         strFixedList = strList
      End If
      strSQL = "INSERT INTO TaskList (Title,EnteredBy,Status,DateSubmitted,TimeSubmitted,Priority,List)" & vbCRLF
      strSQL = strSQL & "VALUES ('" & strNewTask & "','" & objNameCheckSet(0) & "','Open','" & Date() & "','" & Time() & "','Normal','" & Replace(strFixedList,"'","''") & "')"
      Application("Connection").Execute(strSQL)
      strAddTask = "Added"
   End If
   
   If strCMD = "Add List" And strNewList <> "" Then
      strSQL = "INSERT INTO Lists (ListName)" & vbCRLF
      strSQL = strSQL & "VALUES ('" & strNewList & "')"
      Application("Connection").Execute(strSQL)
      strAddList = "Added"
   End If
   
   Select Case strList
      Case "Complete"
         strSQL = "SELECT ID,Title,DateSubmitted FROM TaskList WHERE Status='Complete'"
         strTitle = "Completed Tasks"
      Case "", "Open"
         strSQL = "SELECT ID,Title,DateSubmitted FROM TaskList WHERE Status='Open'"
         strTitle = "All Tasks"
      Case "Unassigned"
         strSQL = "SELECT ID,Title,DateSubmitted FROM TaskList WHERE (List='' OR List Is Null) AND Status <> 'Complete'"
         strTitle = "Tasks Not Assigned to Lists"
      Case Else
         strSQL = "SELECT ID,Title,DateSubmitted FROM TaskList WHERE List='" & Replace(strList,"'","''") & "' AND Status <> 'Complete'"
         strTitle = strList
   End Select
   Set objTaskList = Application("Connection").Execute(strSQL)
   
   strSQL = "SELECT DISTINCT ListName" & vbCRLF
   strSQL = strSQL & "FROM Lists"
   Set objLists =  Application("Connection").Execute(strSQL)
   
   If objNameCheckSet(5) = "User" Then
      intColumns = 4
   Else
      intColumns = 3
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
         <li class="topbar">Tasks<font class="separator"> | </font></li>
      <% End If %>
		<% If Application("UseStats") Then %>
			<li class="topbar"><a class="linkbar" href="stats.asp">Stats</a><font class="separator"> | </font></li> 
      <% End If %>
      <% If Application("UseDocs") And objNameCheckSet(6) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="docs.asp">Docs</a><font class="separator"> | </font></li>
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
      <form method="POST" action="tasklist.asp">
      <tr><td valign="top" align="center">
         Choose a List: 
         <select name="List">	
            <option></option>
            <option value="Open">All Open Tasks</option>
            <option value="Complete">Completed Tasks</option>  
            <option value="Unassigned">Not Assigned to List</option>
      <% Do Until objLists.EOF 
            If Not IsNull(objLists(0)) Or objLists(0) <> "" Then%>
            <option value="<%=objLists(0)%>"><%=objLists(0)%></option>
         <% End If   
            objLists.MoveNext
         Loop %>
         </select>
         <input type="submit" name="cmdSubmit" value="Select">
      </td></tr>
      </form>
      <tr><td><hr /></td></tr>
<% If objNameCheckSet(5) = "User" Then %>     
      <tr><td valign="top" align="center">
      <table width="750">
         <form method="POST" action="tasklist.asp">
         <tr>
            <td align="center">
               New Task
               <input type="text" name="NewTask" size="<%=intInputSize%>" />
               <input type="submit" name="cmdSubmit" value="Add Task" />
            </td>
         </tr>
         <tr>
            <td align="center">
               New List
               <input type="text" name="NewList" size="<%=intInputSize/2%>" />
               <input type="submit" name="cmdSubmit" value="Add List" />
            </td>
         <% If strAddList <> "" Then %>
               <tr><td align ="center"><font class="information"><%=strAddList%></font></td></tr>
         <% End If %>
         </tr>
         <input type="hidden" name="List" value="<%=strList%>" />
         </form>
      </table>
      </td></tr>
      <tr><td><hr /></td></tr>
<% End If %>
      <% If NOT objTaskList.EOF Then %>
      <table width="750">
         <form method="POST" action="tasklist.asp">
         <tr><th colspan="<%=intColumns%>"><%=strTitle%></th></tr>
         <tr>
            <th>&nbsp;</th>
      <% If objNameCheckSet(5) = "User" Then %>
            <th>&nbsp;</th>
      <% End If %>
            <th>Task #</th>
            <th>Task</th>
         </tr>
         <% Do Until objTaskList.EOF %>
         <tr>
            <td class="showborders" align="center" width="1%">
            <% If strList <> "Complete" Then %>
                  <a href="task.asp?ID=<%=objTaskList(0)%>">
            <% End If %>
            
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
                  <img border="0" src="../themes/<%=Application("Theme")%>/images/task.gif"></a></center>
            <% Else %>
                  <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/task.gif"></a></center>
            <% End If %>
            
            <% If strList <> "Complete" Then %>
                  </a>
            <% End If %>
            </td>
         <% If objNameCheckSet(5) = "User" Then %> 
               <td class="showborders" align="center" width="1%">
                  <input type="checkbox" name="chkComplete" value="<%=objTaskList(0)%>">
               </td>
         <% End If %>
            <td class="showborders" align="center" width="10%">
               <%=objTaskList(0)%>
            </td>
            <td class="showborders"width="650px"> 
            <% If strList <> "Complete" Then %>
                  <a href="task.asp?ID=<%=objTaskList(0)%>"><%=objTaskList(1)%></a>
            <% Else %>
                  <%=objTaskList(1)%>
            <% End If %>
            </td>
         </tr>
         <%    objTaskList.MoveNext
            Loop %>
         <tr><td colspan="4" align="right">   
      <% If objNameCheckSet(5) = "User" Then %>
         <% If strTitle = "Completed Tasks" Then %>
            <input type="submit" name="cmdSubmit" value="Not Complete">
            <input type="submit" name="cmdSubmit" value="Delete">
         <% Else %>
            <input type="submit" name="cmdSubmit" value="Mark Complete">
         <% End If %>
      <% End If %>   
         </td></tr>
         <input type="hidden" name="List" value="<%=strList%>" />
         </form>   
      </table>
      <% Else %>
            <table width="750">
               <form method="POST" action="tasklist.asp">
               <tr><th><%=strTitle%></th></tr>
               <tr><td align="center">No Tasks Found</td></tr>
            <% If strList <> "" And strList <> "Complete" And strList <> "Open" And strList <> "Complete" Then %>   
                  <tr><td align="center"><input type="submit" name="cmdSubmit" value="Delete List"></td></tr>
            <% End If %>
               <input type="hidden" name="List" value="<%=strList%>" />
               </form>
            </table>
      <% End If %>
   </table>
   </center>
   </body>

<% End Sub %>

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