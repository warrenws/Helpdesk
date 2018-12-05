<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/9/14
'Last Updated 6/9/14

'This is a mobile version of the stats page

Option Explicit

On Error Resume Next

Dim strUser, strSQL, objNameCheckSet, strRole, strUserAgent

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

Const USERLEVEL = 1
Const ACTIVE = 2

strUser = GetUser
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

'Get the user's information from the database
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE UserName='" & strUser & "'"
Set objNameCheckSet = Application("Connection").Execute(strSQL)

strRole = objNameCheckSet(USERLEVEL)

'See if the user has the rights to visit this page
If objNameCheckSet(ACTIVE) Then

   'An error would be generated if the user has NTFS rights to get in but is not found
   'in the database.  In this case the user is denied access.
   If Err Then
      AccessDenied
   Else
      AccessGranted
   End If
Else
   AccessDenied
End If %>

<%Sub AccessGranted %>
   
<% Select Case Request.QueryString("Report")
      Case "MonthlyTotals"
         MonthlyTotals
      Case "MonthlyAverageTimes"
         MonthlyAverageTimes
      Case "WeekdayTotalsPerLocation"
         WeekdayTotalsPerLocation
      Case "Category"
         Category
      Case Else
         MainMenu
   End Select
End Sub %>

<%Sub MainMenu %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   </head>



   Help Desk Reports <br /> <br />

   1. <a href="index.asp?Report=MonthlyTotals">Monthly Totals</a> <br />
   2. <a href="index.asp?Report=MonthlyAverageTimes">Monthly Average Times</a> <br />
   3. <a href="index.asp?Report=WeekdayTotalsPerLocation">Weekday Totals Per Location</a> <br />
   4. <a href="index.asp?Report=Category">Categories</a> <br />
   
<%End Sub %>

<%Sub MonthlyTotals 

   Dim strSQL, objFirstTicket, datDate, intMonth, intYear, datStart, datEnd, objTicketCount, intCount, strType
   
   Select Case Request.QueryString("Type")
      Case "Line"
         strType = "LineChart"
      Case "Bar"
         strType = "ColumnChart"
      Case Else
         strType = "LineChart"
   End Select

   %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
       <script type="text/javascript" src="https://www.google.com/jsapi"></script>
       <script type="text/javascript">
         google.load("visualization", "1", {packages:["corechart"]});
         google.setOnLoadCallback(drawChart);
         function drawChart() {
           var data = google.visualization.arrayToDataTable([
             ['Month', 'Completed Tickets'],
   <%

   strSQL = "SELECT SubmitDate FROM Main"
   Set objFirstTicket = Application("Connection").Execute(strSQL)

   If Not objFirstTicket.EOF Then

      datDate = objFirstTicket(0) - (DatePart("d", objFirstTicket(0)) - 1)

      Do Until datDate > Date
         
         intMonth = DatePart("m",datDate)
         intYear = DatePart("yyyy",datDate)
         
         datStart = datDate
         datDate = DateAdd("m",1,datDate)
         datEnd = datDate - 1

         strSQL = "SELECT Count(ID) AS CountOfID" & vbCRLF
         strSQL = strSQL & "FROM Main" & vbCRLF
         strSQL = strSQL & "WHERE SubmitDate>#" & datStart & "# And SubmitDate<#" & datEnd & "#" & vbCRLF
         Set objTicketCount = Application("Connection").Execute(strSQL)
         
         If objTicketCount.EOF Then
            intCount = 0
         Else
            intCount = objTicketCount(0)
         End If %>
         
         ['<%=intMonth & "/" & intYear%>',  <%=intCount%>], 
   <% Loop 
   End If %>
           ]);

            var options = {
            title: 'Help Desk Tickets Completed Per Month',
            bar: {groupWidth: "95%"},
            titleTextStyle:{fontSize: 14},
            titlePosition: 'out', 
            legend:{position: 'none'},
            chartArea:{width:'85%', height:'80%'},
            };

            var chart = new google.visualization.<%=strType%>(document.getElementById('chart_div'));
            chart.draw(data, options);
         }
       </script>
   </head>   
   <center>
   <a href="index.asp">Back</a>
   
   <div id="chart_div" style="width: 750px; height: 500px;"></div>
   <a href="index.asp?Report=MonthlyTotals&Type=Line">Line</a> | 
   <a href="index.asp?Report=MonthlyTotals&Type=Bar">Bar</a> 
   </center>

<%End Sub %>

<%Sub MonthlyAverageTimes 

   Dim strSQL, objFirstTicket, datDate, intMonth, intYear, datStart, datEnd, objTicketCount, intCount
   Dim strDays, strHours, strMinutes, strAvgTicketTime, strType
   
   Select Case Request.QueryString("Type")
      Case "Line"
         strType = "LineChart"
      Case "Bar"
         strType = "ColumnChart"
      Case Else
         strType = "LineChart"
   End Select
   
   %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
       <script type="text/javascript" src="https://www.google.com/jsapi"></script>
       <script type="text/javascript">
         google.load("visualization", "1", {packages:["corechart"]});
         google.setOnLoadCallback(drawChart);
         function drawChart() {
           var data = google.visualization.arrayToDataTable([
             ['Month', 'Average Time (In Minutes)'],
   <%

   strSQL = "SELECT SubmitDate FROM Main"
   Set objFirstTicket = Application("Connection").Execute(strSQL)

   If Not objFirstTicket.EOF Then

      datDate = objFirstTicket(0) - (DatePart("d", objFirstTicket(0)) - 1)

      Do Until datDate > Date
         
         intMonth = DatePart("m",datDate)
         intYear = DatePart("yyyy",datDate)
         
         datStart = datDate
         datDate = DateAdd("m",1,datDate)
         datEnd = datDate - 1

         strSQL = "SELECT Avg(OpenTime) AS AvgOfOpenTime" & vbCRLF
         strSQL = strSQL & "FROM Main" & vbCRLF
         strSQL = strSQL & "WHERE SubmitDate>#" & datStart & "# And SubmitDate<#" & datEnd & "# AND Status='Complete'" & vbCRLF
         Set objTicketCount = Application("Connection").Execute(strSQL)
         
         If objTicketCount.EOF Then
            intCount = 0
         Else
            If Not IsNull(objTicketCount(0)) Then
				intCount = Round(objTicketCount(0),2)

			Else
				intCount = 0
			End If
         End If
         
      %>
          ['<%=intMonth & "/" & intYear%>',  <%=intCount%>], 
      <%   
      Loop 

   End If
   %>
           ]);

            var options = {
            title: 'Monthly Average Times',
            bar: {groupWidth: "95%"},
            titleTextStyle:{fontSize: 14},
            titlePosition: 'out', 
            legend:{position: 'none'},
            chartArea:{width:'85%', height:'80%'},
           };

            var chart = new google.visualization.<%=strType%>(document.getElementById('chart_div'));
            chart.draw(data, options);
         }
       </script>
   </head>   
   <center>
   <a href="index.asp">Back</a>
   
   <div id="chart_div" style="width: 750px; height: 500px;"></div>
   <a href="index.asp?Report=MonthlyAverageTimes&Type=Line">Line</a> | 
   <a href="index.asp?Report=MonthlyAverageTimes&Type=Bar">Bar</a> 
   </center>

<%End Sub %>

<%Sub WeekdayTotalsPerLocation 

   Dim strSQL, objSites, intCounter, objNumberPerDay, intSunday, intMonday, intTuesday, intWednesday
   Dim intThursday, intFriday, intSaturday, strCode, bolFound, strType
   
   Select Case Request.QueryString("Type")
      Case "Line"
         strType = "LineChart"
      Case "Bar"
         strType = "ColumnChart"
      Case Else
         strType = "LineChart"
   End Select
   
   strSQL = "SELECT Location FROM Location ORDER BY Location"
   Set objSites = Application("Connection").Execute(strSQL)

   strCode = "['Weekday',"
   Do Until objSites.EOF
      strCode = strCode & "'" & objSites(0) & "', "
      objSites.MoveNext
   Loop
   objSites.MoveFirst
   strCode = strCode & "]," & vbCRLF
   
   %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
       <script type="text/javascript" src="https://www.google.com/jsapi"></script>
       <script type="text/javascript">
         google.load("visualization", "1", {packages:["corechart"]});
         google.setOnLoadCallback(drawChart);
         function drawChart() {
           var data = google.visualization.arrayToDataTable([ 
             <%=strCode%>
   <%

      
   For intCounter = 1 to 7 
      strSQL = "SELECT Sum(CountOfID) As SumofDay, Location" & vbCRLF
      strSQL = strSQL & "FROM (SELECT Count(ID) AS CountOfID, SubmitDate, Location" & vbCRLF
      strSQL = strSQL & "FROM (SELECT ID, SubmitDate, Location FROM Main)" & vbCRLF
      strSQL = strSQL & "GROUP BY  Location, SubmitDate" & vbCRLF
      strSQL = strSQL & "HAVING ((Weekday([SubmitDate])=" & intCounter & ")))" & vbCRLF
      strSQL = strSQL & "GROUP BY Location" & vbCRLF
      strSQL = strSQL & "ORDER BY Location" & vbCRLF
      Set objNumberPerDay = Application("Connection").Execute(strSQL)
      
      If Not objNumberPerDay.EOF Then
         Select Case intCounter
            Case 1
               strCode = "['Sunday',"
            Case 2
               strCode = "['Monday',"
            Case 3
               strCode = "['Tuesday',"
            Case 4
               strCode = "['Wednesday',"
            Case 5
               strCode = "['Thursday',"
            Case 6
               strCode = "['Friday',"
            Case 7
               strCode = "['Saturday',"
         End Select 
        
         objSites.MoveFirst
         Do Until objSites.EOF 
         
            objNumberPerDay.MoveFirst
            bolFound = False
            
            Do Until objNumberPerDay.EOF
               If objSites(0) = objNumberPerDay(1) Then
                  strCode = strCode & objNumberPerDay(0) & ","
                  bolFound = True
               End If
               objNumberPerDay.MoveNext
            Loop
            
            If Not bolFound Then
               strCode = strCode & "0,"
            End If
        
            objSites.MoveNext
         Loop
         If intCounter = 7 Then
            Response.Write strCode & "]" & vbCRLF
         Else
            Response.Write strCode & "]," & vbCRLF
         End If
      End If
   Next %>

             ]);

            var options = {
            title: 'Weekday Totals Per Location',
            bar: {groupWidth: "95%"},
            titleTextStyle:{fontSize: 14},
           };

            var chart = new google.visualization.<%=strType%>(document.getElementById('chart_div'));
            chart.draw(data, options);
         }
       </script>
   </head>   
   <center>
   <a href="index.asp">Back</a>
   
   <div id="chart_div" style="width: 750px; height: 500px;"></div>
   <a href="index.asp?Report=WeekdayTotalsPerLocation&Type=Line">Line</a> | 
   <a href="index.asp?Report=WeekdayTotalsPerLocation&Type=Bar">Bar</a> 
   </center>

<%End Sub %>

<%Sub Category 
   
   Dim strSQL, objCategories, strCategory
   
   
   %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
       <script type="text/javascript" src="https://www.google.com/jsapi"></script>
       <script type="text/javascript">
         google.load("visualization", "1", {packages:["corechart"]});
         google.setOnLoadCallback(drawChart);
         function drawChart() {
           var data = google.visualization.arrayToDataTable([
             ['Category', 'Tickets'],
   
   <% strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
      strSQL = strSQL & "FROM Main" & vbCRLF
      strSQL = strSQL & "GROUP BY Category" & vbCRLF
      strSQL = strSQL & "ORDER BY Category" & vbCRLF
      Set objCategories = Application("Connection").Execute(strSQL)
      
      If Not objCategories.EOF Then
         Do Until objCategories.EOF 
            If objCategories(0) = "" Then
               strCategory = "<No Category>"
            Else
               strCategory = objCategories(0)
            End If %>
            ['<%=strCategory%>', <%=objCategories(1)%>], 
      <%    objCategories.MoveNext
         Loop
      End If %>
      
           ]);

         var options = {
            title: 'Categories',
            is3D: true,
            pieSliceText: 'percent',
            chartArea:{left:25,top:20,width:'95%',height:'95%'},
            titleTextStyle:{fontSize: 14},
            legend:{textStyle: {fontSize: 14}, alignment: 'center'},
            };

         var chart = new google.visualization.PieChart(document.getElementById('chart_div'));
         chart.draw(data, options);
         }
       </script>
   </head>   
   <center>
   <a href="index.asp">Back</a>
   
   <div id="chart_div" style="width: 750px; height: 500px;"></div>

   </center>      

<%End Sub%>


<%Sub AccessDenied 

   If bolShowLogout Then
      SendToLogonScreen
   Else
   %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
      <link rel="stylesheet" type="text/css" href="../../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
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

<%Function BuildReturnLink(bolIncludeID)

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

End Function%>

<%Function GetUser

   Const USERNAME = 1

   Dim strUserAgent, strSessionID, objSessionLookup, objNetwork
   
   'Get the users logon name from IIS
   Set objNetwork = CreateObject("WSCRIPT.Network")   
   GetUser = objNetwork.UserName
   
   'Check and see if anonymous access is enabled in IIS
   If LCase(Left(GetUser,4)) = "iusr" Then
   
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
   End If
   
End Function %>

<%Sub SendToLogonScreen

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
   
End Sub %>

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

End Sub %>