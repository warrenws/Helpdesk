<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/11
'Last Updated 6/16/14

'This is the charts page.

Option Explicit

On Error Resume Next

Dim strSQL, strRole, strUserAgent, objNetwork, objNameCheckSet, objCounters, objTicketsPerTech
Dim intItemsToDisplay, strUser, bolShowLogout, objTicketsOpenPerTech, objTicketsOpenPerLocation
Dim intHighestValue, objTicketsPerLocation, objNumberPerDay, intCounter, strTechName

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
   intItemsToDisplay = 10
   
   strSQL = "SELECT Feedback FROM Counters WHERE ID=1"
   Set objCounters = Application("Connection").Execute(strSQL) 
   
   strSQL = "SELECT Tech, Count(ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>'Complete'" & vbCRLF
   strSQL = strSQL & "GROUP BY Tech" & vbCRLF
   strSQL = strSQL & "ORDER BY Tech"
   Set objTicketsOpenPerTech = Application("Connection").Execute(strSQL) 

   strSQL = "SELECT Location, Count(ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>'Complete'" & vbCRLF
   strSQL = strSQL & "GROUP BY Location" & vbCRLF
   strSQL = strSQL & "ORDER BY Location"
   Set objTicketsOpenPerLocation = Application("Connection").Execute(strSQL) 
   
   strSQL = "SELECT TOP " & intItemsToDisplay & " Main.Tech, Count(Main.ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Tech INNER JOIN Main ON Tech.Tech = Main.Tech" & vbCRLF
   strSQL = strSQL & "WHERE Status='Complete' AND Active=True" & vbCRLF
   strSQL = strSQL & "GROUP BY Main.Tech" & vbCRLF
   strSQL = strSQL & "ORDER BY Count(Main.ID) DESC" 
   Set objTicketsPerTech = Application("Connection").Execute(strSQL) 
   
   strSQL = "SELECT TOP 10 Main.Location, Count(Main.ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Location INNER JOIN Main ON Location.Location = Main.Location" & vbCRLF
   strSQL = strSQL & "WHERE Status='Complete' AND Active=True" & vbCRLF
   strSQL = strSQL & "GROUP BY Main.Location" & vbCRLF
   strSQL = strSQL & "ORDER BY Count(Main.ID) DESC"
   Set objTicketsPerLocation = Application("Connection").Execute(strSQL) 
   
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

      <script type="text/javascript" src="https://www.google.com/jsapi"></script>
      <script type="text/javascript">
      google.load("visualization", "1", {packages:["corechart"]});
      google.setOnLoadCallback(drawOpenPerTechChart);
      google.setOnLoadCallback(drawOpenPerLocationChart);
      google.setOnLoadCallback(drawTicketsPerTech);
      google.setOnLoadCallback(drawTicketsPerLocation);
      google.setOnLoadCallback(drawTicketsPerWeekday);
      
      function drawOpenPerTechChart() {
         var data = google.visualization.arrayToDataTable([ 
            ['Tech', 'Open Tickets'],
      <% If Not objTicketsOpenPerTech.EOF Then
            Do Until objTicketsOpenPerTech.EOF 
               If objTicketsOpenPerTech(0) = "" Or objTicketsOpenPerTech(0) = " " Then 
                  strTechName = "Unassigned"
               Else
                  strTechName = objTicketsOpenPerTech(0) 
               End If%>
               ['<%=strTechName%>', <%=objTicketsOpenPerTech(1)%>],
         <% objTicketsOpenPerTech.MoveNext 
            Loop 
         End If%>    

         ]);

         var options = {
            title: 'Number of Open Tickets Per Tech',
            is3D: true,
            pieSliceText: 'percent',
            chartArea:{left:25,top:20,width:'95%',height:'95%'},
            titleTextStyle:{fontSize: 14},
            legend:{textStyle: {fontSize: 14}, alignment: 'center'},
         };
         
         var chart = new google.visualization.PieChart(document.getElementById('openPerTech'));
         chart.draw(data, options);
      }
      
      function drawOpenPerLocationChart() {
         var data = google.visualization.arrayToDataTable([ 
            ['Location', 'Open Tickets'],
      <% If Not objTicketsOpenPerLocation.EOF Then
            Do Until objTicketsOpenPerLocation.EOF %>
               ['<%=objTicketsOpenPerLocation(0)%>', <%=objTicketsOpenPerLocation(1)%>],
         <% objTicketsOpenPerLocation.MoveNext 
            Loop 
         End If%>   

         ]);

         var options = {
            title: 'Number of Open Tickets Per Location',
            is3D: true,
            pieSliceText: 'percent',
            chartArea:{left:25,top:20,width:'95%',height:'95%'},
            titleTextStyle:{fontSize: 14},
            legend:{textStyle: {fontSize: 14}, alignment: 'center'},
         };
         
         var chart = new google.visualization.PieChart(document.getElementById('openPerLocation'));
         chart.draw(data, options);
      }
      
      function drawTicketsPerTech() {
         var data = google.visualization.arrayToDataTable([ 
            ['Tech', 'Completed Tickets'],
      <% intHighestValue = 0
         If Not objTicketsPerTech.EOF Then
            Do Until objTicketsPerTech.EOF 
               If objTicketsPerTech(1) > intHighestValue Then
                  intHighestValue = objTicketsPerTech(1)
               End If
            %>
               ['<%=Trim(Left(objTicketsPerTech(0), InStr(objTicketsPerTech(0)," ")))%>', <%=objTicketsPerTech(1)%>],
         <% objTicketsPerTech.MoveNext 
            Loop 
         End If%>   

         ]);

         var options = {
            title: 'Total Number of Tickets Per Tech - Top <%=intItemsToDisplay%>',
            bar: {groupWidth: "95%"},
            vAxis: {viewWindow: {max : <%=intHighestValue+(intHighestValue*.2)%>}},
            titleTextStyle:{fontSize: 14},
            titlePosition: 'out', 
            legend:{position: 'none'},
            chartArea:{width:'90%', height:'80%'},
         };
         
         var chart = new google.visualization.ColumnChart(document.getElementById('PerTech'));
         chart.draw(data, options);
      }
      
      function drawTicketsPerLocation() {
         var data = google.visualization.arrayToDataTable([ 
            ['Location', 'Completed Tickets'],
      <% intHighestValue = 0
         If Not objTicketsPerLocation.EOF Then
            Do Until objTicketsPerLocation.EOF 
               If objTicketsPerLocation(1) > intHighestValue Then
                  intHighestValue = objTicketsPerLocation(1)
               End If
            %>
               ['<%=objTicketsPerLocation(0)%>', <%=objTicketsPerLocation(1)%>],
         <% objTicketsPerLocation.MoveNext 
            Loop 
         End If%>   

         ]);

         var options = {
            title: 'Total Number of Tickets Per Location - Top <%=intItemsToDisplay%>',
            bar: {groupWidth: "95%"},
            vAxis: {viewWindow: {max : <%=intHighestValue%>}},
            titleTextStyle:{fontSize: 14},
            titlePosition: 'out', 
            legend:{position: 'none'},
            chartArea:{width:'90%', height:'80%'},
         };
         
         var chart = new google.visualization.ColumnChart(document.getElementById('PerLocation'));
         chart.draw(data, options);
      }
      
      function drawTicketsPerWeekday() {
         var data = google.visualization.arrayToDataTable([ 
            ['Weekday', 'Submitted Tickets'],
      <% 
      intHighestValue = 0
      For intCounter = 1 to 7 
            strSQL = "SELECT Sum(CountOfID) As SumofDay" & vbCRLF
            strSQL = strSQL & "FROM (SELECT Count(ID) AS CountOfID, SubmitDate" & vbCRLF
            strSQL = strSQL & "FROM Main" & vbCRLF
            strSQL = strSQL & "GROUP BY SubmitDate" & vbCRLF
            strSQL = strSQL & "HAVING ((Weekday([SubmitDate])=" & intCounter & ")))"
            Set objNumberPerDay = Application("Connection").Execute(strSQL)
            
            If Not objNumberPerDay.EOF Then
               Select Case intCounter
                  Case 1
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Sunday', <%=objNumberPerDay(0)%>],
                  <% End If
                  Case 2
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Monday', <%=objNumberPerDay(0)%>],
                  <% End If
                  Case 3
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Tuesday', <%=objNumberPerDay(0)%>],
                  <% End If
                  Case 4
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Wednesday', <%=objNumberPerDay(0)%>],
                  <% End If
                  Case 5
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Thursday', <%=objNumberPerDay(0)%>],
                  <% End If
                  Case 6
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Friday', <%=objNumberPerDay(0)%>],
                  <% End If
                  Case 7
                     If Not IsNull(objNumberPerDay(0)) Then %>
                        ['Saturday', <%=objNumberPerDay(0)%>],
                  <% End If
               End Select
               If objNumberPerDay(0) > intHighestValue Then
                  intHighestValue = objNumberPerDay(0)
               End If
            End If
         Next 
      
      %>   

         ]);

         var options = {
            title: 'Tickets Submitted Per Weekday',
            bar: {groupWidth: "95%"},
            vAxis: {viewWindow: {max : <%=intHighestValue%>}},
            titleTextStyle:{fontSize: 14},
            titlePosition: 'out', 
            legend:{position: 'none'},
            chartArea:{width:'90%', height:'80%'},
         };
         
         var chart = new google.visualization.ColumnChart(document.getElementById('PerWeekday'));
         chart.draw(data, options);
      }
            
      </script>
   
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
   <a href="detailed.asp">Detailed</a><font class="separator"> | </font> 
   <a href="date.asp">By Date</a> <font class="separator"> | </font> 
   Charts
<% If objCounters(0) > 0 Then %>
   <font class="separator"> | </font>
   <a href="feedback.asp">Feedback</a>
<% End If %>  
   <hr class="adminbottombar"/>
   <center>
      <div id="openPerTech" style="width: 750px; height: 250px;"></div> <br />
      <div id="openPerLocation" style="width: 750px; height: 250px;"></div> <br />
      <div id="PerTech" style="width: 750px; height: 300px;"></div> <br />
      <div id="PerLocation" style="width: 750px; height: 300px;"></div> <br />
      <div id="PerWeekday" style="width: 750px; height: 300px;"></div> <br />
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