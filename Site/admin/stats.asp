<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 3/24/14
'Last Updated 6/16/14

'This is the summary stats page.

Option Explicit

On Error Resume Next

Dim strUserAgent, strUser, bolShowLogout, strRole, objNameCheckSet, strSQL, objNetwork

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

strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

If UserAuthenticated Then
   AccessGranted
Else
   AccessDenied
End If
%>

<%
Sub AccessGranted 

   Dim strSQL, objCounters, objTicketCount, intTicketCount, objTotalOpenTickets, intTotalOpenTickets
   Dim objTodaysNewTickets, intTodaysNewTickets, objCompleteTicketCount, intCompleteCount
   Dim objAvgTicketTime, strDays, strHours, strMinutes, strAvgTicketTime
   Dim objNumOpenTicketsPerTech, objNumOpenTicketsPerLocation
   
   'Find out if there is any feedback.  This is used to show the feedback link if needed
   strSQL = "SELECT Feedback FROM Counters WHERE ID=1"
   Set objCounters = Application("Connection").Execute(strSQL)   
   
   'Find the total number of tickets in the system
   strSQL = "SELECT Count(ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Main"
   Set objTicketCount = Application("Connection").Execute(strSQL)
   
   If objTicketCount.EOF Then
      intTicketCount = 0
   Else
      intTicketCount = objTicketCount(0)
   End If
   
   'Find the number of open tickets in the system
   strSQL = "SELECT Count(Name) AS CountOfName" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete"""
   Set objTotalOpenTickets = Application("Connection").Execute(strSQL) 
   
   If objTotalOpenTickets.EOF Then
      intTotalOpenTickets = 0
   Else
      intTotalOpenTickets = objTotalOpenTickets(0)
   End If
   
   'Find the number of tickets submitted today
   strSQL = "SELECT Count(Main.ID) AS Tickets" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "GROUP BY SubmitDate" & vbCRLF
   strSQL = strSQL & "HAVING SubmitDate=Date()"
   Set objTodaysNewTickets = Application("Connection").Execute(strSQL) 
   
   If objTodaysNewTickets.EOF Then
      intTodaysNewTickets = 0
   Else
      intTodaysNewTickets = objTodaysNewTickets(0)
   End If
   
   'Find the number of tickets completed today
   strSQL = "SELECT Count(ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "GROUP BY LastUpdatedDate, Status" & vbCRLF
   strSQL = strSQL & "HAVING LastUpdatedDate=Date() AND Status='Complete'"
   Set objCompleteTicketCount = Application("Connection").Execute(strSQL)
   
   If objCompleteTicketCount.EOF Then
      intCompleteCount = 0
   Else
      intCompleteCount = objCompleteTicketCount(0)
   End If   
   
   'Find the average ticket time.
   strSQL = "SELECT Avg(OpenTime) AS AvgOfOpenTime" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "Where Main.OpenTime<>''"
   Set objAvgTicketTime = Application("Connection").Execute(strSQL) 

   strDays = Int(objAvgTicketTime(0)/1440)
   strHours = Int((objAvgTicketTime(0)-strDays*1440)/60)
   strMinutes = (objAvgTicketTime(0)-(strDays*1440)-(strHours*60))
   strAvgTicketTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m"    
   
   'Find the number of open tickets for each tech
   strSQL = "SELECT Tech,Count(Tech) AS OpenTicketsperTech" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete""" & vbCRLF
   strSQL = strSQL & "GROUP BY Tech" & vbCRLF
   strSQL = strSQL & "ORDER BY Count(Tech) DESC"
   Set objNumOpenTicketsPerTech = Server.CreateObject("ADODB.Recordset")
   objNumOpenTicketsPerTech.CursorLocation = 3
   objNumOpenTicketsPerTech.Open strSQL, Application("Connection")

   'Find the number of open tickets for each location
   strSQL = "SELECT Location, Count(Location) AS OpenTicketsperSite" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete""" & vbCRLF
   strSQL = strSQL & "GROUP BY Location" & vbCRLF
   strSQL = strSQL & "ORDER BY Count(Location) DESC"
   Set objNumOpenTicketsPerLocation = Server.CreateObject("ADODB.Recordset")
   objNumOpenTicketsPerLocation.CursorLocation = 3
   objNumOpenTicketsPerLocation.Open strSQL, Application("Connection")   

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
      <% If Application("UseTaskList") And objNameCheckSet(5) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="tasklist.asp">Tasks</a><font class="separator"> | </font></li>
      <% End If %>
	   <% If Application("UseStats") Then %>
			<li class="topbar">Stats<font class="separator"> | </font></li> 
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

   Summary<font class="separator"> | </font> 
   <a href="detailed.asp">Detailed</a><font class="separator"> | </font> 
   <a href="date.asp">By Date</a> <font class="separator"> | </font> 
   <a href="charts.asp">Charts</a>
<% If objCounters(0) > 0 Then %>
   <font class="separator"> | </font>
   <a href="feedback.asp">Feedback</a>
<% End If %>     
   <hr class="adminbottombar"/>
   
   <center>
   <table width="750">
      <tr><td valign="top" align="center">  
         <table>
            <tr>
               <th colspan="2">Database Stats</th>
            <tr>
               <td class="showborders">Total tickets</td>
               <td class="showborders" align="center">&nbsp;&nbsp;<%=intTicketCount%>&nbsp;&nbsp;</td>
            </tr>
            <tr>
               <td class="showborders"><a href="view.asp?Filter=AllOpenTickets">Open tickets</a></td>
               <td class="showborders" align="center"><%=intTotalOpenTickets%></td>
            </tr>
            <tr>
               <td class="showborders"><a href="view.asp?date=<%=Date()%>&type=submitted">Submitted today</a></td>
               <td class="showborders" align="center"><%=intTodaysNewTickets%></td>
            </tr>
            <tr>
               <td class="showborders"><a href="view.asp?date=<%=Date()%>&type=completed">Closed today</a>&nbsp;&nbsp;</td>
               <td class="showborders" align="center"><%=intCompleteCount%></td>
            </tr>                 
            <tr>
               <td class="showborders" align="center" colspan="2">Avg Time: <%=strAvgTicketTime%></td>
            </tr>
         </table> 
      </td>
      <td valign="top" align="center">
         <table>
            <tr><th colspan="2">Tickets Per Tech</th></tr>
            <% Do Until objNumOpenTicketsPerTech.EOF %>
               <tr>
                  <% If objNumOpenTicketsPerTech(0) = "Jacqueline Chromczak" Then %>
                        <td class="showborders"><a href="view.asp?Tech=<%=objNumOpenTicketsPerTech(0)%>&Status=Any%20Open%20Ticket">Jackie Chromczak</a></td>
                  <% ElseIf objNumOpenTicketsPerTech(0) = "" Or objNumOpenTicketsPerTech(0) = " " Then %>
                        <td class="showborders"><a href="view.asp?Tech=Nobody&Status=Any%20Open%20Ticket">Unassigned</a></td>
                  <% Else %>
                        <td class="showborders"><a href="view.asp?Tech=<%=objNumOpenTicketsPerTech(0)%>&Status=Any%20Open%20Ticket"><%=objNumOpenTicketsPerTech(0)%></a></td>
                  <% End If %>
                  <td class="showborders" align="center">&nbsp;<%=objNumOpenTicketsPerTech(1)%>&nbsp;</td>
               </tr>
            <%    objNumOpenTicketsPerTech.MoveNext
               Loop%>
         </table>
      </td>
      <td valign="top" align="center">
         <table>
            <tr><th colspan="2">Tickets Per Location</th></tr>
            <% Do Until objNumOpenTicketsPerLocation.EOF %>
               <tr>
                  <td class="showborders"><a href="view.asp?Location=<%=objNumOpenTicketsPerLocation(0)%>&Status=Any%20Open%20Ticket"><%=objNumOpenTicketsPerLocation(0)%></a></td>
                  <td class="showborders" align="center"><%=objNumOpenTicketsPerLocation(1)%></td>
               </tr>
            <%    objNumOpenTicketsPerLocation.MoveNext
               Loop%>
         </table> 
      </td>
      </tr>
      <tr><td colspan="3"><hr /></td></tr> 
   </table>
     
   
<%
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
Function UserAuthenticated

   On Error Resume Next

   'Get the users logon name
   Set objNetwork = CreateObject("WSCRIPT.Network")   
   strUser = objNetwork.UserName 'This variable should be global for the whole page
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
   
   'Check and see if anonymous access is enabled
   If LCase(Left(strUser,4)) = "iusr" Then
      strUser = GetUser
      bolShowLogout = True 'This variable should be global for the whole page
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
         UserAuthenticated = True
      Else
         If Application("UseStats") Then
            UserAuthenticated = True
         Else
            UserAuthenticated = False
         End If
      End If
   Else
      UserAuthenticated = False
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