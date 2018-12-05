<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/11
'Last Updated 6/16/14

'This is the detailed stats page.

Option Explicit

On Error Resume Next

Dim objNumOpenTicketsPerTech, objNumOpenTicketsPerLocation,objAvgTimePerTech
Dim objAvgTimePerLocation, strSQL, strAvgOpenTime, strDays, strHours, strMinutes
Dim strTech, intOpenTickets, strLocation, objNetwork, objNameCheckSet
Dim objNumTicketsPerTech, objNumTicketsPerLocation, intTotalTickets, objTopUsers
Dim objTotalTickets, objTotalOpenTickets, objAvgTicketTime, strAvgTicketTime
Dim objOldestTicket, strOldestTicket, objAvgTaskTime, strAvgTaskTime, intPercentOpen
Dim objAvgOpened, objAvgClosed, strRole, strUserAgent, objCounters, strTotalNumberOfTickets
Dim strFixedTechName, intTotalTicktets, intTechCount, intTopUserCount, strStartDate, strEndDate
Dim intAvgOpenedPerDay, intAvgClosedPerDay, objFirstTicket, strUser, bolShowLogout

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
      If Application("UseStats") Then
         AccessGranted
      Else
         AccessDenied
      End If
   End If
Else
   AccessDenied
End If

Sub AccessGranted
   
   On Error Resume Next

   strSQL = "SELECT SubmitDate FROM Main"
   Set objFirstTicket = Application("Connection").Execute(strSQL)   
   
   'Get the date of the first ticket.
   If Not objFirstTicket.EOF Then
      strStartDate = objFirstTicket(0)
   Else
      strStartDate = Date
   End If
   
   'Set the end date to today
   strEndDate = Date
   
   'Get the start date from the form
   If IsDate(Request.Form("StartDate")) Then
      strStartDate = Request.Form("StartDate")
   End If
   
   'Get the end date from the form
   If IsDate(Request.Form("EndDate")) Then
      strEndDate = Request.Form("EndDate")
   End If
   
   'Get the start date from the URL
   If IsDate(Request.QueryString("StartDate")) Then
      strStartDate = Request.QueryString("StartDate")
   End If
   
   'Get the end date from the URL
   If IsDate(Request.QueryString("EndDate")) Then
      strEndDate = Request.QueryString("EndDate")
   End If
   
   strSQL = "SELECT Feedback FROM Counters WHERE ID=1"
   Set objCounters = Application("Connection").Execute(strSQL)   
   
   strSQL = "SELECT Tech,Count(Tech) AS OpenTicketsperTech" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete""" & vbCRLF
   strSQL = strSQL & "GROUP BY Tech"
   Set objNumOpenTicketsPerTech = Application("Connection").Execute(strSQL)

   strSQL = "SELECT Location, Count(Location) AS OpenTicketsperSite" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete""" & vbCRLF
   strSQL = strSQL & "GROUP BY Location"
   Set objNumOpenTicketsPerLocation = Application("Connection").Execute(strSQL)

   strSQL = "SELECT Log.NewValue, Avg(SumOfTaskTime) AS AvgOfTaskTime" & vbCRLF
   strSQL = strSQL & "FROM (SELECT Log.Ticket, Log.NewValue, Sum(Log.TaskTime) AS SumOfTaskTime" & vbCRLF
   strSQL = strSQL & "FROM (Tech INNER JOIN Log ON Tech.Tech = Log.NewValue) INNER JOIN Main ON Log.Ticket = Main.ID" & vbCRLF
   strSQL = strSQL & "WHERE SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   strSQL = strSQL & "GROUP BY Log.Ticket, Log.NewValue)" & vbCRLF
   strSQL = strSQL & "Group By Log.NewValue;"
   Set objAvgTimePerTech = Application("Connection").Execute(strSQL)

   strSQL = "SELECT Location, Avg(OpenTime) AS AvgOpenTimeperLocation" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status=""Complete"" AND SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   strSQL = strSQL & "GROUP BY Location"
   Set objAvgTimePerLocation = Application("Connection").Execute(strSQL) 
   
   strSQL = "SELECT Tech,Count(Tech) AS CountOfTech" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   strSQL = strSQL & "GROUP BY Tech;" & vbCRLF
   Set objNumTicketsPerTech = Application("Connection").Execute(strSQL) 

   strSQL = "SELECT Location,Count(Location) AS CountOfLocation" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   strSQL = strSQL & "GROUP BY Location;"
   Set objNumTicketsPerLocation = Application("Connection").Execute(strSQL) 
   
   strSQL = "SELECT Count(Name) AS CountOfName" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   Set objTotalTickets = Application("Connection").Execute(strSQL) 
  
   strTotalNumberOfTickets = "Total Number of Tickets:"
      
   strSQL = "SELECT Count(Name) AS CountOfName" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete"""
   Set objTotalOpenTickets = Application("Connection").Execute(strSQL) 
   
   If objTotalTickets(0) <> 0 Then
      intPercentOpen = Round((objTotalOpenTickets(0)/objTotalTickets(0))*100,5)
   Else   
      intPercentOpen = 0
   End If
   
   strSQL = "SELECT Avg(OpenTime) AS AvgOfOpenTime" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Main.OpenTime<>'' AND SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   Set objAvgTicketTime = Application("Connection").Execute(strSQL) 
   
   strDays = Int(objAvgTicketTime(0)/1440)
   strHours = Int((objAvgTicketTime(0)-strDays*1440)/60)
   strMinutes = (objAvgTicketTime(0)-(strDays*1440)-(strHours*60))
   strAvgTicketTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" 
   
   strSQL = "SELECT Top 1 SubmitDate, SubmitTime, ID" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>'Complete'" & vbCRLF
   strSQL = strSQL & "ORDER BY SubmitDate"
   Set objOldestTicket = Application("Connection").Execute(strSQL) 
   
   strDays = DateDiff("d",objOldestTicket(0),Date)
   strMinutes = DateDiff("n",objOldestTicket(1),Time)
   strHours = (strMinutes / 60)
   strMinutes = strMinutes Mod 60
   If Sgn(strHours) = -1 Then
      strHours = (24 + strHours)
      strDays = strDays - 1
   End If
   If Sgn(strMinutes) = -1 Then
      strMinutes = 60 + strMinutes
   End If
   strOldestTicket = strDays & "d " & Int(strHours) & "h " & strMinutes & "m" 
   
   strSQL = "SELECT Avg(TaskTime) AS AvgOfTaskTime" & vbCRLF
   strSQL = strSQL & "FROM Log INNER JOIN Main ON Log.Ticket = Main.ID" & vbCRLF
   strSQL = strSQL & "WHERE Log.TaskTime<>'' AND SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   Set objAvgTaskTime = Application("Connection").Execute(strSQL) 

   strDays = Int(objAvgTaskTime(0)/1440)
   strHours = Int((objAvgTaskTime(0)-strDays*1440)/60)
   strMinutes = (objAvgTaskTime(0)-(strDays*1440)-(strHours*60))
   strAvgTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" 

   strSQL = "SELECT Avg(AverageofCount.CountofName) AS AvgOfCountofName" & vbCRLF
   strSQL = strSQL & "FROM (SELECT SubmitDate, Count(Name) AS CountOfName" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   strSQL = strSQL & "GROUP BY SubmitDate)  AS AverageofCount;"
   Set objAvgOpened = Application("Connection").Execute(strSQL)
   
   If NOT IsNull(objAvgOpened(0)) Then
      intAvgOpenedPerDay = Round(objAvgOpened(0),1)
   End If
   
   strSQL = "SELECT Avg(AverageofCount.CountofName) AS AvgOfCountofName" & vbCRLF
   strSQL = strSQL & "FROM (SELECT LastUpdatedDate, Count(Name) AS CountOfName" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status = 'Complete' AND SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
   
   strSQL = strSQL & "GROUP BY LastUpdatedDate)  AS AverageofCount;"
   Set objAvgClosed = Application("Connection").Execute(strSQL)
   
   If NOT IsNull(objAvgClosed(0)) Then
      intAvgClosedPerDay = Round(objAvgClosed(0),1)
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
   
   <a href="stats.asp">Summary</a><font class="separator"> | </font> 
   Detailed<font class="separator"> | </font> 
   <a href="date.asp">By Date</a> <font class="separator"> | </font> 
   <a href="charts.asp">Charts</a>
<% If objCounters(0) > 0 Then %>
   <font class="separator"> | </font>
   <a href="feedback.asp">Feedback</a>
<% End If %>    
   <hr class="adminbottombar"/>

   <center>
   <table width="750">
      <tr><td colspan="3" align="center">Date Range: <%=strStartDate%> - <%=strEndDate%></td></tr>
      <tr><td colspan="3" align="center"><hr /></td></tr>
      <tr><td valign="top" align="center">
      <table>
         <tr>
            <th>&nbsp;&nbsp;&nbsp;Technician&nbsp;&nbsp;&nbsp;</th>
            <th>&nbsp;&nbsp;&nbsp;Tickets&nbsp;&nbsp;&nbsp;</th>
            <th>&nbsp;&nbsp;&nbsp;Open&nbsp;&nbsp;&nbsp;</th>
            <th>&nbsp;&nbsp;&nbsp;Avg Time&nbsp;&nbsp;&nbsp;</th>
         </tr>
   <% strAvgOpenTime = ""
      intTechCount = 0
      Do Until objAvgTimePerTech.EOF
         intTechCount = intTechCount + 1
         intTotalTickets = 0
         strTech = objAvgTimePerTech(0)
         strAvgOpenTime = ""
         If objAvgTimePerTech(1) <> 0 Then
            strDays = Int(objAvgTimePerTech(1)/1440)
            strHours = Int((objAvgTimePerTech(1)-strDays*1440)/60)
            strMinutes = (objAvgTimePerTech(1)-(strDays*1440)-(strHours*60))
            strAvgOpenTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m"      
         Else
            strAvgOpenTime = "N/A"
         End If
         intOpenTickets = "0"
         Do Until objNumOpenTicketsPerTech.EOF 
            If objNumOpenTicketsPerTech(0) = objAvgTimePerTech(0) Then
               intOpenTickets = objNumOpenTicketsPerTech(1)
            End If 
            objNumOpenTicketsPerTech.MoveNext
         Loop
         objNumOpenTicketsPerTech.MoveFirst
         
         If Not objNumTicketsPerTech.EOF Then
            Do Until objNumTicketsPerTech.EOF
               If objNumTicketsPerTech(0) = objAvgTimePerTech(0) Then
                  intTotalTickets = objNumTicketsPerTech(1)
               End If
               objNumTicketsPerTech.MoveNext
            Loop
            objNumTicketsPerTech.MoveFirst
         Else
            intTotalTickets = 0
         End If
         
         Select Case strTech
            Case "Crystal Jones-Howe"
               strFixedTechName = "Crystal Howe"
            Case "Jacqueline Chromczak"
               strFixedTechName = "Jackie Chromczak"
            Case Else
               strFixedTechName = strTech
         End Select
         
         %>
         
         <tr>
            <td class="showborders"><a href="view.asp?Tech=<%=strTech%>&Status=Any%20Open%20Ticket"><%=strFixedTechName%></a></td>
            <td class="showborders" align="center"><%=intTotalTickets%></td>
            <td class="showborders" align="center"><%=intOpenTickets%></td>
            <td class="showborders" align="center"><%=strAvgOpenTime%></td>
         </tr>   
   <%    objAvgTimePerTech.MoveNext
      Loop 
         %>
         
      </table>
      </td>
      <td valign="top" align="center">
      <table>
         <tr>
            <th>&nbsp;&nbsp;&nbsp;<a href="topusers.asp">Top Users</a>&nbsp;&nbsp;&nbsp;</th>
            <th>&nbsp;&nbsp;&nbsp;Tickets&nbsp;&nbsp;&nbsp;</th>
         </tr>
   <%    If intTechCount <> 0 Then
            strSQL = "SELECT Top " & intTechCount & " Main.DisplayName, Count(Main.DisplayName) AS CountOfDisplayName" & vbCRLF
            strSQL = strSQL & "FROM Main" & vbCRLF
            strSQL = strSQL & "WHERE SubmitDate>=#" & strStartDate & "# And SubmitDate<=#" & strEndDate & "#" & vbCRLF
            strSQL = strSQL & "GROUP BY Main.DisplayName" & vbCRLF
            strSQL = strSQL & "ORDER BY Count(Main.DisplayName) DESC;"
            Set objTopUsers = Application("Connection").Execute(strSQL)     
      
            intTopUserCount = 0
            Do Until objTopUsers.EOF or intTechCount = intTopUserCount %>
               <tr>
                  <td class="showborders"><a href="view.asp?User=<%=Replace(objTopUsers(0)," ","%20")%>"><%=objTopUsers(0)%></td>
                  <td class="showborders" align="center"><%=objTopUsers(1)%></td>
               </tr>
      <%       intTopUserCount = intTopUserCount + 1
               objTopUsers.MoveNext
            Loop 
         
         End If %>
         
      </table>
      <td>
      
      <tr><td colspan=2><hr /></td></tr>
      <tr>
      <td valign="top" align="center">
      <table>
         <tr>
            <th class="showborders">&nbsp;&nbsp;&nbsp;Location&nbsp;&nbsp;&nbsp;</th>
            <th class="showborders">&nbsp;&nbsp;&nbsp;Tickets&nbsp;&nbsp;&nbsp;</th>
            <th class="showborders">&nbsp;&nbsp;&nbsp;Open&nbsp;&nbsp;&nbsp;</th>
            <th class="showborders">&nbsp;&nbsp;&nbsp;Average Time&nbsp;&nbsp;&nbsp;</th>
         </tr>
   <% Do Until objAvgTimePerLocation.EOF
         If objAvgTimePerLocation(1) <> 0 Then
            strDays = Int(objAvgTimePerLocation(1)/1440)
            strHours = Int((objAvgTimePerLocation(1)-strDays*1440)/60)
            strMinutes = (objAvgTimePerLocation(1)-(strDays*1440)-(strHours*60))
            strAvgOpenTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" 
         Else
            strAvgOpenTime = "N/A"
         End If
         strLocation = objAvgTimePerLocation(0)
         
		   intOpenTickets = "0"
         Do Until objNumOpenTicketsPerLocation.EOF
            If objAvgTimePerLocation(0) = objNumOpenTicketsPerLocation(0) Then
               strLocation = objNumOpenTicketsPerLocation(0)
               intOpenTickets = objNumOpenTicketsPerLocation(1)
            End If
            objNumOpenTicketsPerLocation.MoveNext
         Loop 
         objNumOpenTicketsPerLocation.MoveFirst
         
		   intTotalTicktets = 0
         Do Until objNumTicketsPerLocation.EOF
            If objAvgTimePerLocation(0) = objNumTicketsPerLocation(0) Then
               intTotalTicktets = objNumTicketsPerLocation(1)
            End If
            objNumTicketsPerLocation.MoveNext
         Loop 
         objNumTicketsPerLocation.MoveFirst%>
        
         <tr>
            <td class="showborders"><a href="view.asp?Location=<%=strLocation%>&Status=Any%20Open%20Ticket"><%=strLocation%></a></td>
            <td class="showborders" align="center"><%=intTotalTicktets%></td>
            <td class="showborders" align="center"><%=intOpenTickets%></td>
            <td class="showborders" align="center"><%=strAvgOpenTime%></td>
         </tr>
   <%    objAvgTimePerLocation.MoveNext
      Loop%>
      
      </table>
      </td>
      <td valign="top" align="center">
      <table>
         <tr><th colspan="2">Database Stats</th></tr>
         <tr>
         
            <td class="showborders"><%=strTotalNumberOfTickets%>&nbsp;</td>
            <td class="showborders" align="center"><%=objTotalTickets(0)%></td>
         </tr>
         <tr>
            <td class="showborders">Open Tickets:&nbsp;</td>
            <td class="showborders" align="center"><%=objTotalOpenTickets(0)%></td>
         </tr>
         <tr>
            <td class="showborders">Percent Open:&nbsp;</td>
            <td class="showborders" align="center">&nbsp;<%=intPercentOpen%>%&nbsp;</td>
         </tr>
         <tr>
            <td class="showborders">Average Ticket Time:&nbsp;</td>
            <td class="showborders" align="center">&nbsp;<%=strAvgTicketTime%>&nbsp;</td>
         </tr>
         <tr>
            <td class="showborders">Average Task Time:&nbsp;</td>
            <td class="showborders" align="center">&nbsp;<%=strAvgTaskTime%>&nbsp;</td>
         </tr>
         <tr>
            <td class="showborders">Average Opened Per Day:&nbsp;</td>
            <td class="showborders" align="center">&nbsp;<%=intAvgOpenedPerDay%>&nbsp;</td>
         </tr>
         <tr>
            <td class="showborders">Average Closed Per Day:&nbsp;</td>
            <td class="showborders" align="center">&nbsp;<%=intAvgClosedPerDay%>&nbsp;</td>
         </tr>
         <tr>
            <td class="showborders">Oldest Ticket: #<a href="modify.asp?ID=<%=objOldestTicket(2)%>"><%=objOldestTicket(2)%></a></td>
            <td class="showborders" align="center">&nbsp;<%=strOldestTicket%>&nbsp;</a></td>
         </tr>
      </table>
      </td>
      </tr>
      <tr><td colspan=2><hr /></td></tr>
   </table>
   <form method="POST" action="detailed.asp">
      Date Range: <input type="text" name="StartDate" value="<%=strStartDate%>" size="10"> - <input type="text" name="EndDate" value="<%=strEndDate%>" size="10">
      <input type="submit" value="Apply" name="Submit"></td>
   </form>
   </center>
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